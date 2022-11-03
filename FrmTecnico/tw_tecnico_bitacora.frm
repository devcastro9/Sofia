VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tw_tecnico_bitacora 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Técnico - Bitacora de Eventos"
   ClientHeight    =   10260
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "tw_tecnico_bitacora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10260
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      Height          =   4220
      Left            =   7320
      TabIndex        =   11
      Top             =   600
      Width           =   11175
      Begin VB.TextBox TxtContrato 
         Height          =   285
         Left            =   1500
         TabIndex        =   85
         Text            =   "0"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Txt_descripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_justificacion"
         DataSource      =   "Ado_datos"
         Height          =   555
         Left            =   1515
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   3000
         Width           =   9405
      End
      Begin VB.TextBox txt_obs 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_observaciones"
         DataSource      =   "Ado_datos"
         Height          =   525
         Left            =   1520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   3600
         Width           =   9405
      End
      Begin VB.TextBox TxtPlazo 
         DataField       =   "PlazoDias"
         Height          =   285
         Left            =   6720
         TabIndex        =   79
         Text            =   "48"
         Top             =   2040
         Width           =   520
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "tw_tecnico_bitacora.frx":0A02
         DataField       =   "TipoContratoCodigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4320
         TabIndex        =   27
         Top             =   2040
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "TipoContratoCodigo"
         BoundColumn     =   "TipoContratoCodigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox txt_obs2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "observaciones2"
         DataSource      =   "Ado_datos"
         Height          =   435
         Left            =   1520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   77
         Top             =   3720
         Visible         =   0   'False
         Width           =   9405
      End
      Begin VB.TextBox Txt_campo3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "doc_numero2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   8685
         MultiLine       =   -1  'True
         TabIndex        =   74
         Top             =   1560
         Width           =   2205
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "tw_tecnico_bitacora.frx":0A1C
         DataField       =   "subproceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4275
         TabIndex        =   20
         Top             =   3420
         Visible         =   0   'False
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "subproceso_descripcion"
         BoundColumn     =   "subproceso_codigo"
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
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   7080
         TabIndex        =   44
         Top             =   450
         Width           =   285
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   10635
         TabIndex        =   39
         Top             =   1110
         Width           =   270
      End
      Begin MSDataListLib.DataCombo dtc_codigo11 
         Bindings        =   "tw_tecnico_bitacora.frx":0A35
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6000
         TabIndex        =   31
         Top             =   1560
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux1 
         Bindings        =   "tw_tecnico_bitacora.frx":0A4F
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4800
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_sigla"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "tw_tecnico_bitacora.frx":0A68
         DataField       =   "subproceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   21
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "subproceso_codigo"
         BoundColumn     =   "subproceso_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "tw_tecnico_bitacora.frx":0A81
         DataField       =   "TipoContratoCodigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Top             =   2040
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "DescripcionTipoContrato"
         BoundColumn     =   "TipoContratoCodigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         Bindings        =   "tw_tecnico_bitacora.frx":0A9B
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2880
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "tw_tecnico_bitacora.frx":0AB4
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9000
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "tw_tecnico_bitacora.frx":0ACD
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4200
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "edif_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "tw_tecnico_bitacora.frx":0AE6
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5460
         TabIndex        =   15
         Top             =   1095
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "tw_tecnico_bitacora.frx":0AFF
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5760
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "tw_tecnico_bitacora.frx":0B18
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1695
         TabIndex        =   0
         Top             =   435
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc11 
         Bindings        =   "tw_tecnico_bitacora.frx":0B31
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1860
         TabIndex        =   1
         Top             =   1560
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "solicitud_fecha_solicitud"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   9435
         TabIndex        =   49
         Top             =   2040
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         Format          =   117702657
         CurrentDate     =   44324
         MaxDate         =   55153
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "tw_tecnico_bitacora.frx":0B4B
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   50
         Top             =   1095
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   5520
         TabIndex        =   82
         Top             =   2520
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         Format          =   117702657
         CurrentDate     =   44324
         MaxDate         =   55153
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   9440
         TabIndex        =   83
         Top             =   2520
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         Format          =   117702657
         CurrentDate     =   44324
         MaxDate         =   55153
         MinDate         =   2
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Contrato:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   180
         TabIndex        =   84
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin de Contrato:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   7320
         TabIndex        =   81
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio de Contrato:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3300
         TabIndex        =   80
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Justificación Técnica . . . . :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   180
         TabIndex        =   75
         Top             =   3600
         Width           =   1245
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo:              Hrs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   4
         Left            =   6120
         TabIndex        =   78
         Top             =   2070
         Width           =   1545
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Penúltimo Párrafo. . . . . :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   180
         TabIndex        =   76
         Top             =   3720
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cite TEC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   3
         Left            =   7560
         TabIndex        =   73
         Top             =   1590
         Width           =   1170
      End
      Begin VB.Label dtc_codigo9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "doc_codigo"
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
         Height          =   300
         Left            =   2580
         TabIndex        =   48
         Top             =   3540
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Trámite:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   47
         Top             =   2070
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cod.Adm./File.Contrato"
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
         Index           =   6
         Left            =   7680
         TabIndex        =   43
         Top             =   180
         Width           =   2070
      End
      Begin VB.Label Txt_campo2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "unidad_codigo_ant"
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
         Height          =   300
         Left            =   7800
         TabIndex        =   42
         Top             =   435
         Width           =   1815
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   180
         TabIndex        =   38
         Top             =   840
         Width           =   660
      End
      Begin VB.Label lbl_descripcion 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto (para Ref.). . :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   180
         TabIndex        =   37
         Top             =   3015
         Width           =   1275
      End
      Begin VB.Label lbl_campo9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Registro ISO"
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
         Left            =   1500
         TabIndex        =   36
         Top             =   3360
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label lbl_campo11 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable CGI:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   180
         TabIndex        =   35
         Top             =   1590
         Width           =   1815
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   5460
         TabIndex        =   34
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Ejecutora"
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
         Left            =   1725
         TabIndex        =   33
         Top             =   180
         Width           =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   10
         X2              =   11150
         Y1              =   788
         Y2              =   788
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
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
         Height          =   300
         Left            =   180
         TabIndex        =   29
         Top             =   435
         Width           =   1215
      End
      Begin VB.Label txt_campo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "doc_numero"
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
         Height          =   300
         Left            =   5460
         TabIndex        =   28
         Top             =   3420
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Doc. ISO"
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
         Index           =   13
         Left            =   3360
         TabIndex        =   23
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Registro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   12
         Left            =   7980
         TabIndex        =   22
         Top             =   2070
         Width           =   1425
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
         Height          =   300
         Left            =   10065
         TabIndex        =   4
         Top             =   435
         Width           =   855
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "#Trámite"
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
         Left            =   180
         TabIndex        =   13
         Top             =   180
         Width           =   795
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
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
         Left            =   10155
         TabIndex        =   12
         Top             =   180
         Width           =   645
      End
   End
   Begin VB.PictureBox BtnImprimir1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   590
      Left            =   220
      Picture         =   "tw_tecnico_bitacora.frx":0B64
      ScaleHeight     =   585
      ScaleWidth      =   1395
      TabIndex        =   72
      ToolTipText     =   "Cotización del Servicio"
      Top             =   8640
      Width           =   1400
   End
   Begin VB.PictureBox BtnImprimir2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   590
      Left            =   120
      Picture         =   "tw_tecnico_bitacora.frx":14C9
      ScaleHeight     =   585
      ScaleWidth      =   1395
      TabIndex        =   68
      ToolTipText     =   "Cotización del Servicio"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   580
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   20280
      TabIndex        =   57
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5280
         Picture         =   "tw_tecnico_bitacora.frx":1EAD
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.PictureBox BtnDesAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6600
         Picture         =   "tw_tecnico_bitacora.frx":26E0
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   9480
         Picture         =   "tw_tecnico_bitacora.frx":30D7
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   10
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
         Picture         =   "tw_tecnico_bitacora.frx":3519
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   64
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
         Picture         =   "tw_tecnico_bitacora.frx":3CD8
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   63
         Top             =   0
         Visible         =   0   'False
         Width           =   1430
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         DataSource      =   "Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6600
         Picture         =   "tw_tecnico_bitacora.frx":45ED
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   62
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2640
         Picture         =   "tw_tecnico_bitacora.frx":4D39
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   60
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3960
         Picture         =   "tw_tecnico_bitacora.frx":54EE
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   59
         ToolTipText     =   "Listado de Trámites Iniciados para Cotizacion"
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17520
         Picture         =   "tw_tecnico_bitacora.frx":5DBB
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   58
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
         TabIndex        =   66
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
      Height          =   580
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   20280
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "tw_tecnico_bitacora.frx":657D
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   55
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
         Picture         =   "tw_tecnico_bitacora.frx":6E69
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   54
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
         TabIndex        =   56
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.PictureBox FrmABMDet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   120
      ScaleHeight     =   2970
      ScaleWidth      =   1545
      TabIndex        =   52
      Top             =   6435
      Width           =   1575
      Begin VB.PictureBox BtnAnlDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   590
         Left            =   100
         Picture         =   "tw_tecnico_bitacora.frx":763F
         ScaleHeight     =   585
         ScaleWidth      =   1215
         TabIndex        =   71
         Top             =   1440
         Width           =   1220
      End
      Begin VB.PictureBox BtnModDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   590
         Left            =   100
         Picture         =   "tw_tecnico_bitacora.frx":7D8B
         ScaleHeight     =   585
         ScaleWidth      =   1425
         TabIndex        =   70
         Top             =   765
         Width           =   1430
      End
      Begin VB.PictureBox BtnAddDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   590
         Left            =   100
         Picture         =   "tw_tecnico_bitacora.frx":86A0
         ScaleHeight     =   585
         ScaleWidth      =   1335
         TabIndex        =   69
         Top             =   40
         Width           =   1335
      End
   End
   Begin VB.PictureBox FrmABMDet2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1430
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   1545
      TabIndex        =   51
      Top             =   4950
      Width           =   1575
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BITACORA"
      ForeColor       =   &H00800000&
      Height          =   2940
      Left            =   1770
      TabIndex        =   45
      Top             =   6480
      Width           =   16800
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "tw_tecnico_bitacora.frx":8E5F
         Height          =   2535
         Left            =   195
         TabIndex        =   46
         Top             =   225
         Width           =   16335
         _ExtentX        =   28813
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "bitacora_codigo"
            Caption         =   "Correl"
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
            DataField       =   "negocia_fecha_real"
            Caption         =   "Fecha Evento"
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
            DataField       =   "negocia_tarea_realizada"
            Caption         =   "Tema Tratado"
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
            DataField       =   "negocia_observaciones"
            Caption         =   "Conclusiones u Observaciones"
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
            DataField       =   "negocia_hora_real"
            Caption         =   "Hora Evento"
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
            DataField       =   "negocia_gasto_estimado"
            Caption         =   "Gasto Estimado"
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
         BeginProperty Column06 
            DataField       =   "negocia_forma"
            Caption         =   "Tipo.Evento"
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "Cliente Contactado"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Personal CGI"
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
               Locked          =   -1  'True
               ColumnWidth     =   524.976
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   3990.047
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   3855.118
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "EQUIPOS QUE INTERVIENEN EN EL SERVICIO"
      ForeColor       =   &H00800000&
      Height          =   1520
      Left            =   1755
      TabIndex        =   24
      Top             =   4860
      Width           =   16760
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "tw_tecnico_bitacora.frx":8E7A
         Height          =   1215
         Left            =   180
         TabIndex        =   25
         Top             =   240
         Width           =   16335
         _ExtentX        =   28813
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "bien_codigo"
            Caption         =   "Codigo "
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
            DataField       =   "bien_codigo_anterior"
            Caption         =   "Nro.Eqp."
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
            DataField       =   "bien_total_venta"
            Caption         =   "Precio.Servicio"
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
         BeginProperty Column03 
            DataField       =   "bien_cantidad_por_empaque"
            Caption         =   "Hrs.X Día"
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
            DataField       =   "marca_codigo"
            Caption         =   "Marca"
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
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo"
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
            DataField       =   "bien_cantidad"
            Caption         =   "Cantidad"
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
            DataField       =   "bien_descripcion"
            Caption         =   "Descripcion del Bien"
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
            DataField       =   "bien_descripcion_anterior"
            Caption         =   "Caracteristicas/Identificacion.Ubicacion"
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
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1544.882
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1635.024
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   7905.26
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GERENCIA GENERAL"
      ForeColor       =   &H00800000&
      Height          =   4185
      Left            =   120
      TabIndex        =   14
      Top             =   615
      Width           =   7095
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "tw_tecnico_bitacora.frx":8E95
         Height          =   3495
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   6165
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
            DataField       =   "solicitud_codigo"
            Caption         =   "#Trámite"
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
            DataField       =   "unidad_codigo"
            Caption         =   "U.Ejecutora"
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
            DataField       =   "edif_codigo"
            Caption         =   "Edificio"
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
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cite.Trámite"
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
         BeginProperty Column06 
            DataField       =   "observacion_proy"
            Caption         =   "Nombre.Edificio"
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
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contratos.Vigentes"
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
         TabIndex        =   40
         Top             =   3825
         Value           =   -1  'True
         Width           =   1935
      End
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
         Left            =   4560
         TabIndex        =   41
         Top             =   3825
         Width           =   1155
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   3765
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
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   11280
      TabIndex        =   5
      Top             =   10260
      Width           =   11280
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   10
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   9600
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
   Begin Crystal.CrystalReport CR01 
      Left            =   9840
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2400
      Top             =   9600
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
      Top             =   9600
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
      Top             =   9600
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
      Top             =   9600
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   11520
      Top             =   9600
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
      Left            =   13800
      Top             =   9600
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
      Left            =   120
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   2400
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   4680
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Ado_detalle1 
      Height          =   330
      Left            =   11520
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Ado_detalle2 
      Height          =   330
      Left            =   13800
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Ado_datos11 
      Height          =   330
      Left            =   6960
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Ado_detalle7 
      Height          =   330
      Left            =   9240
      Top             =   9960
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
      Caption         =   "Ado_detalle7"
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
   Begin MSAdodcLib.Adodc Ado_detalle3 
      Height          =   330
      Left            =   120
      Top             =   10320
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
      Caption         =   "Ado_detalle3"
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
   Begin MSAdodcLib.Adodc Ado_detalle4 
      Height          =   330
      Left            =   2400
      Top             =   10320
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
      Caption         =   "Ado_detalle4"
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
   Begin MSAdodcLib.Adodc Ado_detalle5 
      Height          =   330
      Left            =   4680
      Top             =   10320
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
      Caption         =   "Ado_detalle5"
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
   Begin MSAdodcLib.Adodc Ado_detalle6 
      Height          =   330
      Left            =   6960
      Top             =   10320
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
      Caption         =   "Ado_detalle6"
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
      Left            =   10320
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR03 
      Left            =   10800
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR00 
      Left            =   9360
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR04 
      Left            =   11280
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR05 
      Left            =   11760
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR06 
      Left            =   12240
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "tw_tecnico_bitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset
Dim rs_datos12 As New ADODB.Recordset

Dim rs_det1 As New ADODB.Recordset
Dim rs_det2 As New ADODB.Recordset
Dim rs_det3 As New ADODB.Recordset
Dim rs_det4 As New ADODB.Recordset
Dim rs_det5 As New ADODB.Recordset
Dim rs_det6 As New ADODB.Recordset
Dim rs_det7 As New ADODB.Recordset

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

Dim rsNada As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod, VAR_DET As String
Dim VAR_VAL, VAR_SUBP As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim CodBien As String
Dim VAR_UNI As String
Dim sino As String
Dim parametro As String
Dim VAR_DA, VAR_UORIGEN As String
Dim VAR_DPTO, VAR_DPTOC As String
Dim VAR_TIT, VAR_SUBT As String

Dim VAR_AUX, VAR_CONT2 As Double

Dim VAR_TIPO, VAR_SOL As Integer
Dim iResult, VAR_CITES As Integer

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAddDetalle_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  'marca1 = Ado_datos.Recordset.Bookmark
  
  If rs_datos!estado_codigo <> "ANL" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False
    VAR_SOL = Ado_datos.Recordset!solicitud_codigo
    Call ABRIR_TABLA_DET
    If Ado_datos.Recordset!unidad_codigo = "DNEME" Then
        tw_bitacora_emergencia.txt_codigo.Caption = Me.txt_codigo.Caption
        tw_bitacora_emergencia.Txt_campo1.Caption = Me.dtc_codigo1.Text
        tw_bitacora_emergencia.Txt_descripcion.Caption = Me.dtc_desc1.Text
        tw_bitacora_emergencia.Txt_Correl.Caption = 0    'rs_datos!correl_bitacora + 1
        tw_bitacora_emergencia.Txt_estado.Caption = "REG"
        tw_bitacora_emergencia.lbl_bitacora.Caption = Me.FraDet1.Caption
        Ado_detalle1.Recordset.AddNew
        tw_bitacora_emergencia.Show vbModal
    Else
        tw_solicitud_bitacora.txt_codigo.Caption = Me.txt_codigo.Caption
        tw_solicitud_bitacora.Txt_campo1.Caption = Me.dtc_codigo1.Text
        tw_solicitud_bitacora.Txt_descripcion.Caption = Me.dtc_desc1.Text
        tw_solicitud_bitacora.Txt_Correl.Caption = 0    'rs_datos!correl_bitacora + 1
        tw_solicitud_bitacora.Txt_estado.Caption = "REG"
        tw_solicitud_bitacora.lbl_bitacora.Caption = Me.FraDet1.Caption
        Ado_detalle1.Recordset.AddNew
        tw_solicitud_bitacora.Show vbModal
    End If
    
    VAR_COD2 = Ado_datos.Recordset!solicitud_codigo

     If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "solicitud_codigo = " & VAR_COD2 & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
         If rs_det1.RecordCount > 0 Then
         rs_det1.MoveLast
        End If
     Else
        rs_datos.MoveLast
     End If
    
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    'Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
  'WWWWWWWWWWWW EMERGENCIAS
 '  marca1 = Ado_datos.Recordset.Bookmark
  'If rs_datos!estado_codigo = "REG" Then
    'swnuevo = 1
    'fraOpciones.Enabled = False
    'FraNavega.Enabled = False
    'FraDet1.Enabled = False
    'FrmABMDet.Enabled = False
    'FraDet2.Enabled = False
    'FrmABMDet2.Enabled = False
    'Fra_datos.Enabled = False
    'VAR_SOL = Ado_datos.Recordset!solicitud_codigo
    'Call ABRIR_TABLA_DET
'    If Ado_datos.Recordset!unidad_codigo = "DNEME" Then
'    tw_bitacora_emergencia.txt_codigo.Caption = Me.txt_codigo.Caption
'    tw_bitacora_emergencia.Txt_campo1.Caption = Me.dtc_codigo1.Text
'    tw_bitacora_emergencia.Txt_descripcion.Caption = Me.dtc_desc1.Text
'    tw_bitacora_emergencia.Txt_Correl.Caption = 0    'rs_datos!correl_bitacora + 1
'    tw_bitacora_emergencia.Txt_estado.Caption = "REG"
'    tw_bitacora_emergencia.lbl_bitacora.Caption = Me.FraDet1.Caption
'    Ado_detalle1.Recordset.AddNew
'    Txt_campo2.Visible = False
'    tw_bitacora_emergencia.Show vbModal
'
'    Call ABRIR_TABLA_DET
'
'    swnuevo = 0
'    fraOpciones.Enabled = True
'    FraNavega.Enabled = True
'    FraDet1.Enabled = True
'    FrmABMDet.Enabled = True
'    FraDet2.Enabled = True
'    FrmABMDet2.Enabled = True
'    'Fra_datos.Enabled = True
'  Else
'    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
'  End If
  'WWWWWWWWWWWW
End Sub

Private Sub NuevoDetalle()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos!estado_codigo = "REG" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    FraDet3.Enabled = False
    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False
    Call ABRIR_TABLA_DET
            If VAR_DET = "30000" Then
                Ado_detalle3.Recordset.AddNew
                tw_solicitud_bienes3.txt_codigo.Caption = Me.txt_codigo.Caption
                tw_solicitud_bienes3.Txt_campo1.Caption = Me.dtc_codigo1.Text
                tw_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
                tw_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes3.lbl_det.Caption = VAR_DET     '"34110"
                tw_solicitud_bienes3.Txt_estado.Caption = "REG"
                tw_solicitud_bienes3.Show vbModal
            End If
            If VAR_DET = "39800" Then       'REPUESTOS
                Ado_detalle5.Recordset.AddNew
                tw_solicitud_bienes5.txt_codigo.Caption = Me.txt_codigo.Caption
                tw_solicitud_bienes5.Txt_campo1.Caption = Me.dtc_codigo1.Text
                tw_solicitud_bienes5.Txt_descripcion.Caption = Me.dtc_desc1.Text
                tw_solicitud_bienes5.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes5.lbl_det.Caption = VAR_DET     '"34110"
                tw_solicitud_bienes5.Txt_estado.Caption = "REG"
                tw_solicitud_bienes5.Show vbModal
            End If
            If VAR_DET = "34800" Then
                Ado_detalle6.Recordset.AddNew
                tw_solicitud_bienes6.txt_codigo.Caption = Me.txt_codigo.Caption
                tw_solicitud_bienes6.Txt_campo1.Caption = Me.dtc_codigo1.Text
                tw_solicitud_bienes6.Txt_descripcion.Caption = Me.dtc_desc1.Text
                tw_solicitud_bienes6.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes6.lbl_det.Caption = VAR_DET     '"34110"
                tw_solicitud_bienes6.Txt_estado.Caption = "REG"
                tw_solicitud_bienes6.Show vbModal
            End If
            If VAR_DET = "24300" Then       'SERVICIOS
                Ado_detalle7.Recordset.AddNew
                tw_solicitud_bienes7.txt_codigo.Caption = Me.txt_codigo.Caption
                tw_solicitud_bienes7.Txt_campo1.Caption = Me.dtc_codigo1.Text
                tw_solicitud_bienes7.Txt_descripcion.Caption = Me.dtc_desc1.Text
                tw_solicitud_bienes7.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes7.lbl_det.Caption = VAR_DET     '"34110"
                tw_solicitud_bienes7.Txt_estado.Caption = "REG"
                tw_solicitud_bienes7.Show vbModal
            End If
            
'        Case "TEC-04"    '8. VENTA DE SERVICIOS (EME)
'            Call ABRIR_TABLA_DET
'            Ado_detalle3.Recordset.AddNew
'            tw_solicitud_bienes3.txt_codigo.Caption = Me.txt_codigo.Caption
'            tw_solicitud_bienes3.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            tw_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            tw_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
'            tw_solicitud_bienes3.lbl_det.Caption = "34110"
'            tw_solicitud_bienes3.Txt_estado.Caption = "REG"
'            tw_solicitud_bienes3.Show vbModal
'        Case "TEC-05"    '9. VENTA DE SERVICIOS (MOD)
'            Call ABRIR_TABLA_DET
'            Ado_detalle3.Recordset.AddNew
'            tw_solicitud_bienes3.txt_codigo.Caption = Me.txt_codigo.Caption
'            tw_solicitud_bienes3.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            tw_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            tw_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
'            tw_solicitud_bienes3.lbl_det.Caption = "34110"
'            tw_solicitud_bienes3.Txt_estado.Caption = "REG"
'            tw_solicitud_bienes3.Show vbModal
'
'    End Select
    swnuevo = 0
    Call ABRIR_TABLA_DET
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    FraDet3.Enabled = True
    FrmABMDet3.Enabled = True
   
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If

End Sub

Private Sub BtnAnlDetalle_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  If Ado_detalle1.Recordset.RecordCount > 0 Then
   sino = MsgBox("Está Seguro de ANULAR el Registro Activo --> " + Str(Ado_detalle1.Recordset!bitacora_codigo), vbYesNo + vbQuestion, "Atención")
   If Ado_detalle1.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        VAR_SOL = Ado_datos.Recordset!solicitud_codigo
        parametro = Ado_datos.Recordset!unidad_codigo             'Unidad
        db.Execute "delete ao_solicitud_bienes Where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & " and bitacora_codigo = " & Ado_detalle1.Recordset!bitacora_codigo & " "
        'Ado_detalle1.Recordset.Delete 'adAffectAll
        Call ABRIR_TABLA_DET
      End If
   Else
        MsgBox "No se puede ANULAR, un registro Aprobado o Anulado, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
 Else
     MsgBox "No se puede ANULAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
 End If
End Sub

'Private Sub BtnAprobar_Click()
'  On Error GoTo UpdateErr
''  If Ado_datos.Recordset.RecordCount > 0 Then
''   If Ado_datos.Recordset!beneficiario_codigo = "0" Or Ado_datos.Recordset!beneficiario_codigo = "" Then
''        MsgBox "No se puede APROBAR, debe registrar al Propietario del Proyecto de Edificación: " + lbl_campo4.Caption, vbExclamation, "Validación de Registro"
''        Exit Sub
''   End If
''   Set rs_aux2 = New ADODB.Recordset
''   rs_aux2.Open "Select * from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
''   If rs_aux2.RecordCount > 0 Then
''        VAR_CONT2 = rs_aux2.RecordCount
''   End If
'  VAR_VAL = "OK"
'  Call valida_campos
'  If VAR_VAL = "OK" Then
'   If rs_datos!estado_codigo = "REG" Then       'And VAR_CONT2 > 0 Then
'      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'      db.Execute "update ao_solicitud_bienes set ao_solicitud_bienes.almacen_tipo = ac_bienes.almacen_tipo from ac_bienes where ac_bienes.bien_codigo = ao_solicitud_bienes.bien_codigo"
'        VAR_UNI = Ado_datos.Recordset!unidad_codigo
'        VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'        VAR_SUBP = Ado_datos.Recordset!subproceso_codigo
'        Select Case VAR_SUBP        'dtc_codigo2.Text
'            Case "1"    'SOLO COMPRAS BB y SS
'            Case "2"    'SOLO VENTA DE BIENES
'            Case "TEC-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL
'                Set rs_aux1 = New ADODB.Recordset
'                'SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & "  and edif_codigo = '" & Ado_detalle1.Recordset!edif_codigo & "'  "
'                SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
'                rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'                'If rs_aux1.RecordCount > 0 Then
'                '    MsgBox "El código ya existe, consulte con el administrador del Sistema..."
'                '    var_cod = 0
'                '    Exit Sub
'                'Else
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    'rs_aux2.Open "Select max(trafico_codigo) as Codigo from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
'                    rs_aux2.Open "Select max(trafico_codigo) as Codigo from ao_solicitud_calculo_trafico where unidad_codigo = '" & VAR_UNI & "' ", db, adOpenStatic
'                    If Not rs_aux2.EOF Then
'                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                    End If
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    rs_aux2.Open "Select edif_capacidad_min_trafico as Codigo from ao_solicitud_edificacion where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
'                    If Not rs_aux2.EOF Then
'                        VAR_AUX = rs_aux2!Codigo
'                    End If
'                    rs_aux1.AddNew
'                    'var_cod = rs_aux1.RecordCount + 1
'                    rs_aux1!ges_gestion = glGestion
'                    rs_aux1!unidad_codigo = VAR_UNI
'                    rs_aux1!solicitud_codigo = VAR_SOL
'                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'                    rs_aux1!trafico_codigo = var_cod
'                   ' rs_aux1!trafico_h_capacidad_trafico_parametro = Round(VAR_AUX, 2)
'                    rs_aux1!estado_codigo = "REG"
'                    rs_aux1!Fecha_Registro = Date
'                    rs_aux1!usr_codigo = glusuario
'                    rs_aux1.Update
'                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
'                'End If
'                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
'
'            'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'            Case "TEC-02", "TEC-03", "TEC-04", "TEC-05"     '10. SERVICIO MANTENIMIENTO Y REPARACIONES
'            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'               Set rs_aux4 = New ADODB.Recordset
'               If rs_aux4.State = 1 Then rs_aux4.Close
'               If VAR_SUBP = "TEC-02" Then
'                    rs_aux4.Open "select sum(bien_precio_venta_base) as totbs2, sum(bien_total_venta) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & VAR_UNI & "' and solicitud_codigo =" & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic
'               Else
'                    rs_aux4.Open "select sum(bien_precio_venta_base) as totbs2, sum(bien_total_venta) as totdl2, SUM(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & VAR_UNI & "' and solicitud_codigo =" & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic
'               End If
'               If IsNull(rs_aux4!totbs2) Then
'                    'If CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
'                        MsgBox "No puede Aprobar, debe registrar <" + FraDet2.Caption + "> !! Vuelva a Intentar ...", vbExclamation, "Atención"
'                        If rs_aux4.State = 1 Then rs_aux4.Close
'                        Exit Sub
'                    'End If
'               Else
'
'               Set rs_aux1 = New ADODB.Recordset
'                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'               SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "    "
'               rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'               If rs_aux1.RecordCount > 0 Then
'                    MsgBox "Una Cotización anterior ya fue Aprobada, el Registro Actual se adicionará al que fue aprobado anteriormente ..."
'                    '    var_cod = 0
'                    '    Exit Sub
'                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + rs_aux4!totdl2
'                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + rs_aux4!totdl2 / GlTipoCambioOficial
'                    db.Execute "delete ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & glGestion & "' "
'               Else
'                    'CREA VENTA CABECERA
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & var_cod & "   ", db, adOpenStatic
'                    rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
'                    If Not rs_aux2.EOF Then
'                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                    End If
'                    rs_aux1.AddNew
'                    'var_cod = rs_aux1.RecordCount + 1
'                    rs_aux1!ges_gestion = glGestion
'                    rs_aux1!unidad_codigo = VAR_UNI
'                    rs_aux1!solicitud_codigo = VAR_SOL
'                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'                    rs_aux1!depto_codigo = Left(Ado_datos.Recordset!edif_codigo, 1)
'                    rs_aux1!venta_codigo = var_cod
'                    rs_aux1!beneficiario_codigo = Ado_datos.Recordset!beneficiario_codigo
'                    rs_aux1!venta_monto_total_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_monto_total_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
'                    rs_aux1!venta_monto_cobrado_bs = 0
'                    rs_aux1!venta_monto_cobrado_dol = 0
'                    rs_aux1!venta_saldo_p_cobrar_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_saldo_p_cobrar_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
'                    rs_aux1!venta_cantidad_total = rs_aux4!cant2
'                    rs_aux1!venta_fecha = Ado_datos.Recordset!solicitud_fecha_solicitud
'                    rs_aux1!venta_fecha_inicio = Ado_datos.Recordset!solicitud_fecha_solicitud
'                    'VAR_CONT2 = 365 / 30 * rs_aux4!cant2
'                    rs_aux1!venta_plazo_dias_calendario = 0 'VAR_CONT2
'
'                    rs_aux1!correl_cobro_prog = 0
'                    rs_aux1!venta_fecha_fin = FormatDateTime(Ado_datos.Recordset!solicitud_fecha_solicitud + VAR_CONT2, vbGeneralDate)
'                    rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
'                    rs_aux1!estado_codigo = "REG"
'                    rs_aux1!Fecha_Registro = Date
'                    rs_aux1!usr_codigo = glusuario
'                    rs_aux1.Update
''                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
'               End If
'                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
'               If var_cod = "" Then
'                    var_cod = rs_aux1!venta_codigo
'               End If
'                'GRABA VENTA DETALLE
'                'wwwwwwwwwwwwwwwwwww
'               Set rs_aux5 = New ADODB.Recordset
'               If rs_aux5.State = 1 Then rs_aux5.Close
'               rs_aux5.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenKeyset, adLockBatchOptimistic   'and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'
'               'Set AdoAux.Recordset = rsAuxDetalle
'               If rs_aux5.RecordCount > 0 Then
'                   'AdoAux.Recordset.MoveFirst
'                  rs_aux5.MoveFirst
'                  While Not rs_aux5.EOF   ' AdoAux.Recordset.EOF
'
'                    Set rs_aux3 = New ADODB.Recordset
'                    If rs_aux3.State = 1 Then rs_aux3.Close
'                    'rs_aux3.Open "Select * from ao_ventas_detalle where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
'                    rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & glGestion & "'   ", db, adOpenKeyset, adLockOptimistic
'                    'If rs_aux3.RecordCount > 0 Then
'                        'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                    'Else
'                        'db.Execute "delete ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & glGestion & "' "
'                        VAR_AUX = rs_aux3.RecordCount + 1
'                        rs_aux3.AddNew
'                        rs_aux3!ges_gestion = glGestion         'glGestion
'                        rs_aux3!venta_codigo = var_cod
'                        rs_aux3!venta_codigo_det = VAR_AUX
'                        rs_aux3!bien_codigo = rs_aux5!bien_codigo
'                        rs_aux3!venta_det_cantidad = rs_aux5!bien_cantidad
'                        rs_aux3!venta_precio_unitario_bs = rs_aux5!bien_precio_venta_base
'                        rs_aux3!venta_descuento_bs = 0
'                        rs_aux3!venta_precio_total_bs = rs_aux5!bien_total_venta
'                        rs_aux3!venta_precio_unitario_dol = rs_aux5!bien_precio_venta_base / GlTipoCambioOficial
'                        rs_aux3!venta_descuento_dol = 0
'                        rs_aux3!venta_precio_total_dol = rs_aux5!bien_total_venta / GlTipoCambioOficial
'                        'rs_aux3!concepto_venta = dtc_desc2.Text + " - " + Trim(dtc_desc3.Text)
'                        Set rs_aux6 = New ADODB.Recordset
'                        If rs_aux6.State = 1 Then rs_aux6.Close
'                        rs_aux6.Open "Select * from ac_bienes where bien_codigo = '" & rs_aux3!bien_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'                        If rs_aux6.RecordCount > 0 Then
'                            rs_aux3!concepto_venta = rs_aux6!bien_descripcion '+ " - " + Trim(dtc_desc3.Text)
'                        Else
'                            rs_aux3!concepto_venta = "NA1"
'                        End If
'                        rs_aux3!modelo_codigo = rs_aux5!modelo_codigo
'                        rs_aux3!grupo_codigo = rs_aux5!grupo_codigo
'                        rs_aux3!subgrupo_codigo = rs_aux5!subgrupo_codigo
'                        rs_aux3!par_codigo = rs_aux5!par_codigo
'                        'ok
'                        rs_aux3!bien_cantidad_por_empaque = rs_aux5!bien_cantidad_por_empaque
'                        'If rs_aux5!par_codigo = "43340" Or rs_aux5!par_codigo = "99990" Then
'                        If rs_aux5!par_codigo = "43340" Then
'                            db.Execute "update ao_ventas_cabecera set unimed_codigo = '" & rs_aux5!unimed_codigo & "' WHERE venta_codigo = " & var_cod & ""
'                        End If
'                        rs_aux3!bien_codigo_padre = rs_aux5!bien_codigo_padre
'                        rs_aux3!tipo_descuento = 0
'                        rs_aux3!almacen_codigo = 0
'                        rs_aux3!modelo_codigo1 = rs_aux5!modelo_codigo 'do_datos.Recordset!modelo_codigo
'                        rs_aux3!modelo_codigo_h = "S/M" 'Ado_datos.Recordset!modelo_codigo_h
'                        rs_aux3!modelo_codigo_x = "S/M" 'Ado_datos.Recordset!modelo_codigo_x
'                        rs_aux3!modelo_elegido = "N"
'                        rs_aux3!modelo_elegido_h = "N"
'                        rs_aux3!modelo_elegido_x = "N"
'                        rs_aux3!estado_codigo = "REG"
'                        rs_aux3!Fecha_Registro = Date
'                        rs_aux3!usr_codigo = glusuario
'                        rs_aux3.Update
'                        rs_aux5.MoveNext
'                  Wend
'               Else
'                    MsgBox "Error Verifique los datos de Bienes..."
'               End If
'              End If
'            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'        Case "COM-03"    '3. SERVICIO INSTALACION
'            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'               Set rs_aux4 = New ADODB.Recordset
'               If rs_aux4.State = 1 Then rs_aux4.Close
'               'rs_aux4.Open "select sum(bien_precio_compra) as totbs2, sum(bien_total_compra) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & VAR_UNI & "' and solicitud_codigo =" & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic
'               rs_aux4.Open "select sum(bien_precio_venta_base) as totbs2, sum(bien_total_venta) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & VAR_UNI & "' and solicitud_codigo =" & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic
'               If IsNull(rs_aux4!totbs2) Then
'                    'If CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
'                        MsgBox "No puede Aprobar, debe registrar <" + FraDet2.Caption + "> !! Vuelva a Intentar ...", vbExclamation, "Atención"
'                        If rs_aux4.State = 1 Then rs_aux4.Close
'                        Exit Sub
'                    'End If
'               Else
'
'               Set rs_aux1 = New ADODB.Recordset
'                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'               SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "    "
'               rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'               If rs_aux1.RecordCount > 0 Then
'                    MsgBox "Una Cotización anterior ya fue Aprobada, el Registro Actual se adicionará al que fue aprobado anteriormente ..."
'                    '    var_cod = 0
'                    '    Exit Sub
'                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + rs_aux4!totdl2
'                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + rs_aux4!totdl2 / GlTipoCambioOficial
'               Else
'                    'CREA VENTA CABECERA
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
'                    rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
'                    If Not rs_aux2.EOF Then
'                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                    End If
'                    rs_aux1.AddNew
'                    'var_cod = rs_aux1.RecordCount + 1
'                    rs_aux1!ges_gestion = glGestion
'                    rs_aux1!unidad_codigo = VAR_UNI
'                    rs_aux1!solicitud_codigo = VAR_SOL
'                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'                    rs_aux1!venta_codigo = var_cod
'                    rs_aux1!beneficiario_codigo = Ado_datos.Recordset!beneficiario_codigo
'                    rs_aux1!venta_monto_total_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_monto_total_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
'                    rs_aux1!venta_monto_cobrado_bs = 0
'                    rs_aux1!venta_monto_cobrado_dol = 0
'                    rs_aux1!venta_saldo_p_cobrar_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_saldo_p_cobrar_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
'                    rs_aux1!venta_cantidad_total = rs_aux4!cant2
'                    rs_aux1!venta_fecha = Ado_datos.Recordset!solicitud_fecha_solicitud
'                    rs_aux1!venta_fecha_inicio = Ado_datos.Recordset!solicitud_fecha_solicitud
'                    'VAR_CONT2 = 365 / 30 * rs_aux4!cant2
'                    rs_aux1!venta_plazo_dias_calendario = 0 'VAR_CONT2
'                    rs_aux1!correl_cobro_prog = 0
'                    rs_aux1!venta_fecha_fin = FormatDateTime(Ado_datos.Recordset!solicitud_fecha_solicitud + VAR_CONT2, vbGeneralDate)
'                    rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
'                    rs_aux1!estado_codigo = "REG"
'                    rs_aux1!Fecha_Registro = Date
'                    rs_aux1!usr_codigo = glusuario
'                    rs_aux1.Update
''                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
'               End If
'                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
'               If var_cod = "" Then
'                    var_cod = rs_aux1!venta_codigo
'               End If
'                'GRABA VENTA DETALLE
'                'wwwwwwwwwwwwwwwwwww
'               Set rs_aux5 = New ADODB.Recordset
'               If rs_aux5.State = 1 Then rs_aux5.Close
'               rs_aux5.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenKeyset, adLockBatchOptimistic   'and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'
'               'Set AdoAux.Recordset = rsAuxDetalle
'               If rs_aux5.RecordCount > 0 Then
'                   'AdoAux.Recordset.MoveFirst
'                  rs_aux5.MoveFirst
'                  While Not rs_aux5.EOF   ' AdoAux.Recordset.EOF
'
'                    Set rs_aux3 = New ADODB.Recordset
'                    If rs_aux3.State = 1 Then rs_aux3.Close
'                    'rs_aux3.Open "Select * from ao_ventas_detalle where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
'                    rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & glGestion & "'   ", db, adOpenKeyset, adLockOptimistic
'                    'If rs_aux3.RecordCount > 0 Then
'                        'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                    'Else
'                        VAR_AUX = rs_aux3.RecordCount + 1
'                        rs_aux3.AddNew
'                        rs_aux3!ges_gestion = glGestion
'                        rs_aux3!venta_codigo = var_cod
'                        rs_aux3!venta_codigo_det = VAR_AUX
'                        rs_aux3!bien_codigo = rs_aux5!bien_codigo
'                        rs_aux3!venta_det_cantidad = rs_aux5!bien_cantidad
'                        rs_aux3!venta_precio_unitario_bs = rs_aux5!bien_precio_venta_base
'                        rs_aux3!venta_descuento_bs = 0
'                        rs_aux3!venta_precio_total_bs = rs_aux5!bien_total_venta
'                        rs_aux3!venta_precio_unitario_dol = rs_aux5!bien_precio_venta_base / GlTipoCambioOficial
'                        rs_aux3!venta_descuento_dol = 0
'                        rs_aux3!venta_precio_total_dol = rs_aux5!bien_total_venta / GlTipoCambioOficial
'                        'rs_aux3!concepto_venta = dtc_desc2.Text + " - " + Trim(dtc_desc3.Text)
'                        Set rs_aux6 = New ADODB.Recordset
'                        If rs_aux6.State = 1 Then rs_aux6.Close
'                        rs_aux6.Open "Select * from ac_bienes where bien_codigo = '" & rs_aux3!bien_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'                        If rs_aux6.RecordCount > 0 Then
'                            rs_aux3!concepto_venta = rs_aux6!bien_descripcion '+ " - " + Trim(dtc_desc3.Text)
'                        Else
'                            rs_aux3!concepto_venta = "NA1"
'                        End If
'                        rs_aux3!modelo_codigo = rs_aux5!modelo_codigo
'                        rs_aux3!grupo_codigo = rs_aux5!grupo_codigo
'                        rs_aux3!subgrupo_codigo = rs_aux5!subgrupo_codigo
'                        rs_aux3!par_codigo = rs_aux5!par_codigo
'                        'ok
'                        rs_aux3!bien_cantidad_por_empaque = rs_aux5!bien_cantidad_por_empaque
'                        'If rs_aux5!par_codigo = "43340" Or rs_aux5!par_codigo = "99990" Then
'                        If rs_aux5!par_codigo = "43340" Then
'                            db.Execute "update ao_ventas_cabecera set unimed_codigo = '" & rs_aux5!unimed_codigo & "' WHERE venta_codigo = " & var_cod & ""
'                        End If
'                        rs_aux3!tipo_descuento = 0
'                        rs_aux3!almacen_codigo = 0
'                        rs_aux3!modelo_codigo1 = rs_aux5!modelo_codigo 'do_datos.Recordset!modelo_codigo
'                        rs_aux3!modelo_codigo_h = "S/M" 'Ado_datos.Recordset!modelo_codigo_h
'                        rs_aux3!modelo_codigo_x = "S/M" 'Ado_datos.Recordset!modelo_codigo_x
'                        rs_aux3!modelo_elegido = "N"
'                        rs_aux3!modelo_elegido_h = "N"
'                        rs_aux3!modelo_elegido_x = "N"
'                        rs_aux3!estado_codigo = "REG"
'                        rs_aux3!Fecha_Registro = Date
'                        rs_aux3!usr_codigo = glusuario
'                        rs_aux3.Update
'                     rs_aux5.MoveNext
'                  Wend
'               Else
'                    MsgBox "Error Verifique la Venta de Productos..."
'               End If
'              End If
'            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'        Case "COM-04"    '4. SERVICIO AJUSTE
'            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'               Set rs_aux4 = New ADODB.Recordset
'               If rs_aux4.State = 1 Then rs_aux4.Close
'               'rs_aux4.Open "select sum(bien_precio_compra) as totbs2, sum(bien_total_compra) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & VAR_UNI & "' and solicitud_codigo =" & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic
'               rs_aux4.Open "select sum(bien_precio_venta_base) as totbs2, sum(bien_total_venta) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & VAR_UNI & "' and solicitud_codigo =" & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic
'               If IsNull(rs_aux4!totbs2) Then
'                    'If CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
'                        MsgBox "No puede Aprobar, debe registrar <" + FraDet2.Caption + "> !! Vuelva a Intentar ...", vbExclamation, "Atención"
'                        If rs_aux4.State = 1 Then rs_aux4.Close
'                        Exit Sub
'                    'End If
'               Else
'
'               Set rs_aux1 = New ADODB.Recordset
'                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'               SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "    "
'               rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'               If rs_aux1.RecordCount > 0 Then
'                    MsgBox "Una Cotización anterior ya fue Aprobada, el Registro Actual se adicionará al que fue aprobado anteriormente ..."
'                    '    var_cod = 0
'                    '    Exit Sub
'                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + rs_aux4!totdl2
'                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + rs_aux4!totdl2 / GlTipoCambioOficial
'               Else
'                    'CREA VENTA CABECERA
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
'                    rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
'                    If Not rs_aux2.EOF Then
'                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                    End If
'                    rs_aux1.AddNew
'                    'var_cod = rs_aux1.RecordCount + 1
'                    rs_aux1!ges_gestion = glGestion
'                    rs_aux1!unidad_codigo = VAR_UNI
'                    rs_aux1!solicitud_codigo = VAR_SOL
'                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'                    rs_aux1!venta_codigo = var_cod
'                    rs_aux1!beneficiario_codigo = Ado_datos.Recordset!beneficiario_codigo
'                    rs_aux1!venta_monto_total_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_monto_total_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
'                    rs_aux1!venta_monto_cobrado_bs = 0
'                    rs_aux1!venta_monto_cobrado_dol = 0
'                    rs_aux1!venta_saldo_p_cobrar_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_saldo_p_cobrar_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
'                    rs_aux1!venta_cantidad_total = rs_aux4!cant2
'                    rs_aux1!venta_fecha = Ado_datos.Recordset!solicitud_fecha_solicitud
'                    rs_aux1!venta_fecha_inicio = Ado_datos.Recordset!solicitud_fecha_solicitud
'                    'VAR_CONT2 = 365 / 30 * rs_aux4!cant2
'                    rs_aux1!venta_plazo_dias_calendario = 0 'VAR_CONT2
'                    rs_aux1!correl_cobro_prog = 0
'                    rs_aux1!venta_fecha_fin = FormatDateTime(Ado_datos.Recordset!solicitud_fecha_solicitud + VAR_CONT2, vbGeneralDate)
'                    rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
'                    rs_aux1!estado_codigo = "REG"
'                    rs_aux1!Fecha_Registro = Date
'                    rs_aux1!usr_codigo = glusuario
'                    rs_aux1.Update
''                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
'               End If
'                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
'               If var_cod = "" Then
'                    var_cod = rs_aux1!venta_codigo
'               End If
'                'GRABA VENTA DETALLE
'                'wwwwwwwwwwwwwwwwwww
'               Set rs_aux5 = New ADODB.Recordset
'               If rs_aux5.State = 1 Then rs_aux5.Close
'               rs_aux5.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenKeyset, adLockBatchOptimistic   'and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'
'               'Set AdoAux.Recordset = rsAuxDetalle
'               If rs_aux5.RecordCount > 0 Then
'                   'AdoAux.Recordset.MoveFirst
'                  rs_aux5.MoveFirst
'                  While Not rs_aux5.EOF   ' AdoAux.Recordset.EOF
'
'                    Set rs_aux3 = New ADODB.Recordset
'                    If rs_aux3.State = 1 Then rs_aux3.Close
'                    'rs_aux3.Open "Select * from ao_ventas_detalle where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
'                    rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & glGestion & "'   ", db, adOpenKeyset, adLockOptimistic
'                    'If rs_aux3.RecordCount > 0 Then
'                        'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                    'Else
'                        VAR_AUX = rs_aux3.RecordCount + 1
'                        rs_aux3.AddNew
'                        rs_aux3!ges_gestion = glGestion
'                        rs_aux3!venta_codigo = var_cod
'                        rs_aux3!venta_codigo_det = VAR_AUX
'                        rs_aux3!bien_codigo = rs_aux5!bien_codigo
'                        rs_aux3!venta_det_cantidad = rs_aux5!bien_cantidad
'                        rs_aux3!venta_precio_unitario_bs = rs_aux5!bien_precio_venta_base
'                        rs_aux3!venta_descuento_bs = 0
'                        rs_aux3!venta_precio_total_bs = rs_aux5!bien_total_venta
'                        rs_aux3!venta_precio_unitario_dol = rs_aux5!bien_precio_venta_base / GlTipoCambioOficial
'                        rs_aux3!venta_descuento_dol = 0
'                        rs_aux3!venta_precio_total_dol = rs_aux5!bien_total_venta / GlTipoCambioOficial
'                        'rs_aux3!concepto_venta = dtc_desc2.Text + " - " + Trim(dtc_desc3.Text)
'                        Set rs_aux6 = New ADODB.Recordset
'                        If rs_aux6.State = 1 Then rs_aux6.Close
'                        rs_aux6.Open "Select * from ac_bienes where bien_codigo = '" & rs_aux3!bien_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'                        If rs_aux6.RecordCount > 0 Then
'                            rs_aux3!concepto_venta = rs_aux6!bien_descripcion '+ " - " + Trim(dtc_desc3.Text)
'                        Else
'                            rs_aux3!concepto_venta = "NA1"
'                        End If
'                        rs_aux3!modelo_codigo = rs_aux5!modelo_codigo
'                        rs_aux3!grupo_codigo = rs_aux5!grupo_codigo
'                        rs_aux3!subgrupo_codigo = rs_aux5!subgrupo_codigo
'                        rs_aux3!par_codigo = rs_aux5!par_codigo
'                        'ok
'                        rs_aux3!bien_cantidad_por_empaque = rs_aux5!bien_cantidad_por_empaque
'                        'If rs_aux5!par_codigo = "43340" Or rs_aux5!par_codigo = "99990" Then
'                        If rs_aux5!par_codigo = "43340" Then
'                            db.Execute "update ao_ventas_cabecera set unimed_codigo = '" & rs_aux5!unimed_codigo & "' WHERE venta_codigo = " & var_cod & ""
'                        End If
'                        rs_aux3!tipo_descuento = 0
'                        rs_aux3!almacen_codigo = 0
'                        rs_aux3!modelo_codigo1 = rs_aux5!modelo_codigo 'do_datos.Recordset!modelo_codigo
'                        rs_aux3!modelo_codigo_h = "S/M" 'Ado_datos.Recordset!modelo_codigo_h
'                        rs_aux3!modelo_codigo_x = "S/M" 'Ado_datos.Recordset!modelo_codigo_x
'                        rs_aux3!modelo_elegido = "N"
'                        rs_aux3!modelo_elegido_h = "N"
'                        rs_aux3!modelo_elegido_x = "N"
'                        rs_aux3!estado_codigo = "REG"
'                        rs_aux3!Fecha_Registro = Date
'                        rs_aux3!usr_codigo = glusuario
'                        rs_aux3.Update
'                     rs_aux5.MoveNext
'                  Wend
'               Else
'                    MsgBox "Error Verifique la Venta de Productos..."
'               End If
'              End If
'            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'        End Select
'        If rs_datos!unidad_codigo = "DNMAN" Then
'            db.Execute "update ao_solicitud set estado_cotiza = 'APR' where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
'            'rs_datos!estado_cotiza = "APR"
'        End If
'        Set rs_aux2 = New ADODB.Recordset
'        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9 & "'  "
'        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'            txt_campo1.Caption = rs_aux2!correl_doc
'            rs_aux2.Update
'        End If
'        db.Execute "update ao_solicitud set doc_numero = " & txt_campo1.Caption & " where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
'        'rs_datos!doc_numero = txt_campo1.Caption
'        'REVISAR !!! JQA 2014_07_08
'        'VAR_ARCH = RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
'        VAR_ARCH = "TEC_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
'        db.Execute "update ao_solicitud set archivo_respaldo = '" & VAR_ARCH & "' + '.PDF' where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
'        db.Execute "update ao_solicitud set archivo_respaldo_cargado = 'N' where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
'        db.Execute "update ao_solicitud set estado_codigo = 'APR' where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
'        db.Execute "update ao_solicitud set fecha_aprueba = '" & Date & "'  where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
'        db.Execute "update ao_solicitud set usr_codigo_aprueba = '" & glusuario & "'  where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
'        'rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
'        'rs_datos!archivo_respaldo_cargado = "N"
'        'rs_datos!estado_codigo = "APR"
'        'rs_datos!fecha_aprueba = Date
'        'rs_datos!usr_codigo_aprueba = glusuario
'        'rs_datos.UpdateBatch adAffectAll
'        db.Execute "update ao_ventas_detalle set ao_ventas_detalle.almacen_tipo = ac_bienes.almacen_tipo from ac_bienes where ac_bienes.bien_codigo = ao_ventas_detalle.bien_codigo "
'
'
'    VAR_COD2 = Ado_datos.Recordset!solicitud_codigo
'    OptFilGral2_Click
'
'     If (dg_datos.SelBookmarks.Count <> 0) Then
'        dg_datos.SelBookmarks.Remove 0
'     End If
'     If Ado_datos.Recordset.RecordCount > 0 Then
'        rs_datos.Find "solicitud_codigo = " & VAR_COD2 & "   ", , , 1
'        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
'         If rs_det1.RecordCount > 0 Then
'         rs_det1.MoveLast
'        End If
'     Else
'        rs_datos.MoveLast
'     End If
'
'      End If
'   Else
'       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene DETALLE ...", vbExclamation, "Validación de Registro"
'   End If
''  Else
''      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
''  End If
'  End If
'  Exit Sub
'UpdateErr:
'  MsgBox Err.Description
'
'
'End Sub

Private Sub BtnBuscar_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        'Call OptFilGral2_Click
        buscados = 1
        OptFilGral1.Visible = False
        OptFilGral2.Visible = False
'        If OptFilGral1.Value = True Then
'            MsgBox "Esta Buscando los Registros... " + OptFilGral1.Caption, vbInformation, "Atención!"
'        Else
'            MsgBox "Esta Buscando... " + OptFilGral2.Caption + " los Registros.", vbInformation, "Atención!"
'        End If
        Set ClBuscaGrid = New ClBuscaEnGridExterno
        Set ClBuscaGrid.Conexión = db
        ClBuscaGrid.EsTdbGrid = False
        Set ClBuscaGrid.GridTrabajo = dg_datos
        ClBuscaGrid.QueryUtilizado = queryinicial
        Set ClBuscaGrid.RecordsetTrabajo = rs_datos
        'ClBuscaGrid.CamposVisibles = "11010011"
        ClBuscaGrid.Ejecutar
        '        OptFilGral1.Visible = True
'        OptFilGral2.Visible = True
    Else
      MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
      OptFilGral1.Visible = True
      OptFilGral2.Visible = True
    End If
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.Cancel
'        If mvBookMark > 0 Then
'          rs_datos.BookMark = mvBookMark
'        Else
'          rs_datos.MoveFirst
'        End If
        mbDataChanged = False
        Fra_datos.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        dg_det1.Visible = True
        dg_det2.Visible = True
        dg_det3.Visible = True
        dg_det5.Visible = True
        dg_det6.Visible = True
        dg_det7.Visible = True
        
        FrmABMDet2.Enabled = True
        FrmABMDet5.Enabled = True
        FrmABMDet3.Enabled = True
        FrmABMDet6.Enabled = True
        FrmABMDet.Enabled = True
        FrmABMDet7.Enabled = True
        'txt_codigo.Enabled = True
        If rs_datos!solicitud_codigo <> "" Then
        VAR_SOLA = rs_datos!solicitud_codigo
        End If
        
     If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 And VAR_SW = "MOD" Then
        rs_datos.Find "solicitud_codigo = " & VAR_SOLA & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
        rs_datos.MoveLast
     End If
     
        VAR_SW = ""
'        dtc_codigo9.Enabled = True
    End If
'    dtc_desc1.Visible = True
'    lbl_aux1.Visible = False
End Sub

'Private Sub BtnEliminar_Click()
'  On Error GoTo UpdateErr
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    'If ExisteReg(Ado_datos.Recordset!edif_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
'    If ExisteReg(Ado_datos.Recordset!unidad_codigo, Ado_datos.Recordset!solicitud_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
'    If rs_datos!estado_codigo = "APR" Then
'       sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
'       If sino = vbYes Then
'          rs_datos!estado_codigo = "ANL"
'          rs_datos!Fecha_Registro = Date
'          rs_datos!usr_codigo = glusuario
'          rs_datos.UpdateBatch adAffectAll
'       End If
'    Else
'        rs_datos!estado_codigo = "ERR"
'        rs_datos!Fecha_Registro = Date
'        rs_datos!usr_codigo = glusuario
'        rs_datos.UpdateBatch adAffectAll
'       'MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
'    End If
'  Else
'      MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
'  Exit Sub
'
'UpdateErr:
'  MsgBox Err.Description
'End Sub

Private Sub BtnDesAprobar_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_datos!estado_codigo = "APR" Then
      If sino = vbYes Then
         rs_datos!estado_codigo = "REG"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
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
    If VAR_SW = "ADD" Then
        VAR_UNI = dtc_codigo1.Text
        var_cod = IIf(txt_codigo.Caption = "", 0, txt_codigo.Caption)
        Set rs_aux1 = New ADODB.Recordset
        If rs_aux1.State = 1 Then rs_aux1.Close
        SQL_FOR = "Select max(solicitud_codigo) as Codigo from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        'If rs_aux1.RecordCount > 0 Then
        If Not rs_aux1.EOF Then
            var_cod = IIf(IsNull(rs_aux1!Codigo), 1, rs_aux1!Codigo + 1)
        Else
            var_cod = 1
        End If
        Set rs_aux10 = New ADODB.Recordset
        If rs_aux10.State = 1 Then rs_aux10.Close
        SQL_FOR = "Select max(doc_numero2) as Codigo from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' "
        rs_aux10.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If Not rs_aux10.EOF Then
            VAR_CITES = IIf(IsNull(rs_aux10!Codigo), 1, rs_aux10!Codigo + 1)
        Else
            VAR_CITES = 1
        End If
        'var_cod = RTrim(RTrim(dtc_codigo2.Text) + "-") + LTrim(Str(Val(dtc_aux2) + 1))
        txt_codigo.Caption = var_cod
        txt_campo3.Text = VAR_CITES
        ' Guardar con INSERT
        'ges_gestion, unidad_codigo, solicitud_codigo, solicitud_fecha_solicitud, solicitud_fecha_recepción, solicitud_tipo, edif_codigo, beneficiario_codigo,
'                    beneficiario_codigo_resp, beneficiario_codigo_resp2, unidad_codigo_sol, solicitud_justificacion, solicitud_observaciones, proceso_codigo, subproceso_codigo,
'                      etapa_codigo, etapa_codigo2, clasif_codigo, doc_codigo, doc_codigo2, doc_numero, doc_numero2, poa_codigo, ges_gestion_ant, unidad_codigo_ant,
'                      solicitud_codigo_ant, correl_detalle, correl_edificacion, correl_calculo, correl_persona, correl_cotiza, correl_bitacora, archivo_respaldo, archivo_respaldo_cargado,
'                      estado_codigo, estado_etapa2, estado_cotiza, fecha_registro, hora_registro, usr_codigo, usr_codigo_aprueba, fecha_aprueba, hora_aprueba, fecha_registro2,
'                      usr_codigo2 , observacion_proy, mes_codigo
                      
        'db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " &
        '"VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", '" & rs_aux4!bien_codigo & "', '1', " & rs_aux4!venta_precio_unitario_bs & ", '0', " & rs_aux4!venta_precio_total_bs & ", " & rs_aux4!venta_precio_unitario_dol & ", '0', " & rs_aux4!venta_precio_total_dol & ", '" & concepto_venta & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"

        rs_datos!solicitud_codigo = var_cod
        rs_datos!estado_codigo = "REG"      'no cambia
        rs_datos!ges_gestion = glGestion    ' no cambia
        rs_datos!unidad_codigo = VAR_UNI
        'Actualiza correaltivo ...
        db.Execute "Update gc_unidad_ejecutora Set correl_solicitud = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "'   "
        rs_datos!doc_numero = "0"    'txt_campo1.Caption
        'rs_datos!correl_edificacion = 0
        rs_datos!archivo_respaldo = "sin_nombre"
        rs_datos!archivo_respaldo_cargado = "N"
        rs_datos!correl_bitacora = 0
        rs_datos!observaciones2 = txt_obs2.Text
        rs_datos!doc_numero2 = IIf(txt_campo3.Text = "", "0", txt_campo3.Text)
     End If
     If VAR_SW = "MOD" Then
        VAR_UNI = rs_datos!unidad_codigo
        var_cod = rs_datos!solicitud_codigo
     End If
     rs_datos!solicitud_fecha_solicitud = DTPfecha1.Value
     'rs_datos!solicitud_tipo = dtc_codigo2.Text
     rs_datos!EDIF_CODIGO = dtc_codigo3.Text
     If dtc_codigo4.Text = "" Or dtc_codigo4.Text = "0" Then
        rs_datos!beneficiario_codigo = dtc_aux3.Text
     Else
        rs_datos!beneficiario_codigo = dtc_codigo4.Text
     End If
     rs_datos!solicitud_justificacion = Txt_descripcion.Text
     
     If var_cod < 10 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-00000" + Trim(txt_codigo)
     End If
     If var_cod > 9 And var_cod < 100 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-0000" + Trim(txt_codigo)
     End If
     If var_cod > 99 And var_cod < 1000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-000" + Trim(txt_codigo)
     End If
     If var_cod > 999 And var_cod < 10000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-00" + Trim(txt_codigo)
     End If
     If var_cod > 9999 And var_cod < 100000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-0" + Trim(txt_codigo)
     End If
     If var_cod > 99999 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-" + Trim(txt_codigo)
     End If

     'rs_datos!poa_codigo = IIf(dtc_codigo10.Text = "", "3.2.6", dtc_codigo10.Text)
     Select Case dtc_codigo2.Text
        Case "COM-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL - Case "1"    'SOLO COMPRAS BB y SS
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-01-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "COM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-234"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            rs_datos!poa_codigo = "3.1.1"
        Case "CMX-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL
            rs_datos!proceso_codigo = "CMX"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "CMX-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "CMX-01-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "CMX"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-XXX"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            rs_datos!poa_codigo = "4.1.1"
        Case "COM-02"    '3. COMPRA-VENTA BB Y SS - COMERCIAL -         Case "2"    'SOLO VENTA DE BIENES
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-01-02"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "COM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-234"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            rs_datos!poa_codigo = "3.1.1"
        Case "COM-03"    'VENTA DE SERVICIOS INSTTALACIONES
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            rs_datos!poa_codigo = "3.2.2"
        Case "COM-04" '5       'VENTA DE SERVICIOS AJUSTE
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            rs_datos!poa_codigo = "3.2.6"
        Case "TEC-01"    '6. SERVICIO MANTENIMIENTO GRATUITO
            rs_datos!proceso_codigo = "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "TEC-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "TEC-01-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            rs_datos!poa_codigo = "3.2.3"           'IIf(dtc_codigo10.Text = "", "3.2.6", dtc_codigo10.Text)
        Case "TEC-02"    '10. SERVICIO MANTENIMIENTO PREVENTIVO
            'If VAR_UNI = "DNMAN" Then
            rs_datos!solicitud_tipo = "10"
            rs_datos!proceso_codigo = Left(dtc_codigo2.Text, 3) ' "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = IIf(dtc_codigo2.Text = "", "TEC-02", dtc_codigo2.Text)
            rs_datos!etapa_codigo = Trim(dtc_codigo2.Text) + "-01"  '"TEC-02-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = Left(dtc_codigo2.Text, 3)  '"TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-355"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            rs_datos!poa_codigo = "3.2.3"           'IIf(dtc_codigo10.Text = "", "3.2.6", dtc_codigo10.Text)
            'COD.ADM. o CODIGO DE CONTRATO
            rs_datos!unidad_codigo_ant = Trim(Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)) + "-" + Trim(CStr(glGestion))
            'End If
        Case "TEC-03" '10 REPARACION    If VAR_UNI = "DNIREP" Then
                rs_datos!proceso_codigo = "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "TEC-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "TEC-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
                rs_datos!poa_codigo = "3.2.4"       'IIf(dtc_codigo10.Text = "", "3.2.4", dtc_codigo10.Text)
        Case "TEC-04" '10 EMERGENCIAS   If VAR_UNI = "DNEME" Then
                rs_datos!proceso_codigo = "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "TEC-04"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "TEC-04-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
                rs_datos!poa_codigo = "3.2.1"
        Case "TEC-05"    '5. SERVICIO MODERNIZACION -If VAR_UNI = "DNMOD" Then
                rs_datos!proceso_codigo = "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "TEC-05"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "TES-05-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
                rs_datos!poa_codigo = "3.2.7"
        Case Else   '10. SERVICIO MANTENIMIENTO PREVENTIVO
            'If VAR_UNI = "DNMAN" Then
            rs_datos!solicitud_tipo = "10"
            rs_datos!proceso_codigo = Left(dtc_codigo2.Text, 3)     ' "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = IIf(dtc_codigo2.Text = "", "TEC-02", dtc_codigo2.Text)
            rs_datos!etapa_codigo = Trim(dtc_codigo2.Text) + "-01"  '"TEC-02-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = Left(dtc_codigo2.Text, 3)      '"TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-355"                           '
            rs_datos!poa_codigo = "3.2.3"                           '
     End Select
     rs_datos!TipoContratoCodigo = IIf(dtc_codigo10.Text = "", "0", dtc_codigo10.Text)
     rs_datos!PlazoDias = IIf(TxtPlazo.Text = "", "48", TxtPlazo.Text)
     rs_datos!solicitud_observaciones = IIf(txt_obs.Text = "", "", txt_obs.Text)
     rs_datos!observaciones2 = IIf(txt_obs2.Text = "", "", txt_obs2.Text)
     rs_datos!solicitud_fecha_recepción = DTPfecha1.Value
     rs_datos!beneficiario_codigo_resp = dtc_codigo11.Text
     rs_datos!observacion_proy = dtc_desc3.Text
     rs_datos!ges_gestion_ant = glGestion       'glGestion
     rs_datos!usr_codigo_aprueba = ""
     rs_datos!fecha_aprueba = Date
     rs_datos!hora_aprueba = ""
     'rs_datos!Foto = Date
     'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
     'rs_datos!archivo_foto_cargado = "N"
     'hora_registro
     rs_datos!fecha_registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     rs_datos.Update    'Batch 'adAffectAll
     VAR_SOLA = rs_datos!solicitud_codigo
'     If Ado_datos.Recordset!estado_codigo = "REG" Then
'        Call OptFilGral1_Click
'     Else
'        Call OptFilGral2_Click
'     End If
'     rs_datos.MoveLast

         If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 And VAR_SW = "MOD" Then
     VAR_SW = ""
        rs_datos.Find "solicitud_codigo = " & VAR_SOLA & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
     VAR_SW = ""
        rs_datos.MoveLast
     End If
    
     mbDataChanged = False
      
     Fra_datos.Enabled = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
    dg_det1.Visible = True
    dg_det2.Visible = True
    dg_det3.Visible = True
    dg_det5.Visible = True
    dg_det6.Visible = True
    dg_det7.Visible = True
    
    FrmABMDet2.Enabled = True
    FrmABMDet5.Enabled = True
    FrmABMDet3.Enabled = True
    FrmABMDet6.Enabled = True
    FrmABMDet.Enabled = True
    FrmABMDet7.Enabled = True
'     dtc_desc1.BackColor = &HFFFFC0
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
  If (dtc_codigo1.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo3.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo11.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If (dtc_codigo8.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (dtc_codigo9.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If (dtc_codigo10.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo10.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        'CR00.ReportFileName = App.Path & "\Reportes\comercial\ar_solicitud_cotizacion.rpt"
        CR00.ReportFileName = App.Path & "\Reportes\tecnico\tr_lista_solicitud_tecnico.rpt"
        CR00.WindowShowPrintSetupBtn = True
        CR00.WindowShowRefreshBtn = True
        'CR00.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
        CR00.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
        Select Case Me.Ado_datos.Recordset!unidad_codigo
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
      CR00.Formulas(3) = "titulo = '" & var_titulo & "' "
      CR00.Formulas(4) = "subtitulo = '" & lbl_titulo.Caption & "' "

        iResult = CR00.PrintReport
        If iResult <> 0 Then MsgBox CR00.LastErrorNumber & " : " & CR00.LastErrorString, vbCritical, "Error de impresión"
        CR00.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
End Sub

Private Sub BtnImprimir1_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR01.ReportFileName = App.Path & "\Reportes\tecnico\tr_identificacion_cliente.rpt"
        CR01.WindowShowPrintSetupBtn = True
        CR01.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          CR01.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
          CR01.Formulas(1) = "Subtitulo = '" & FraDet1.Caption & "' "
        'Call CREAVISTAF11          'JQA JUN-2008
        CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR01.PrintReport
        If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
        CR01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos... " & FraDet1.Caption, , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
End Sub

Private Sub BtnImprimir2_Click()
    Select Case parametro
        Case "DNINS"            'INI GRABA INSTALACIONES
            'dtc_codigo2.Text = "COM-03" '4
        Case "DNAJS"            'AJUSTE
            'dtc_codigo2.Text = "COM-04" '5
        Case "DNMAN", "DMANS", "DMANB", "DMANC"            'MANTENIMIENTO PREVENTIVO
            If (Ado_datos.Recordset.RecordCount > 0) Then
              If Ado_detalle2.Recordset.RecordCount > 0 Then
                  'Dim iResult As Integer
                  'Dim co As New ADODB.Command
                  '-----------------------------------------------------------------PAGINA 1
                  CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta1.rpt"
                  'CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta.rpt"
                  CR02.WindowShowPrintSetupBtn = True
                  CR02.WindowShowRefreshBtn = True

                  VAR_TIT = "GERENCIA TECNICA"
                  VAR_SUBT = "PROPUESTA SERVICIO DE MANTENIMIENTO INTEGRAL"
                  CR02.Formulas(0) = "Titulo = '" & VAR_TIT & "' "
                  CR02.Formulas(1) = "Subtitulo = '" & VAR_SUBT & "' "
                  CR02.Formulas(2) = "Subtitulo2 = '" & lbl_titulo.Caption & "' "
                  
                  CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                  CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                  iResult = CR02.PrintReport
                  If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
                  CR02.WindowState = crptMaximized
                  '-----------------------------------------------------------------PAGINA 2
                  CR04.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta2.rpt"
                  CR04.WindowShowPrintSetupBtn = True
                  CR04.WindowShowRefreshBtn = True

                  VAR_TIT = "GERENCIA TECNICA"
                  VAR_SUBT = "PROPUESTA SERVICIO DE MANTENIMIENTO INTEGRAL"
                  CR04.Formulas(0) = "Titulo = '" & VAR_TIT & "' "
                  CR04.Formulas(1) = "Subtitulo = '" & VAR_SUBT & "' "
                  CR04.Formulas(2) = "Subtitulo2 = '" & lbl_titulo.Caption & "' "
                  
                  CR04.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CR04.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                  CR04.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                  iResult = CR04.PrintReport
                  If iResult <> 0 Then MsgBox CR04.LastErrorNumber & " : " & CR04.LastErrorString, vbCritical, "Error de impresión"
                  CR04.WindowState = crptMaximized
'                  '-----------------------------------------------------------------PAGINA 3
                  CR05.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta3.rpt"
                  'CR05.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta.rpt"
                  CR05.WindowShowPrintSetupBtn = True
                  CR05.WindowShowRefreshBtn = True

                  VAR_TIT = "GERENCIA TECNICA"
                  VAR_SUBT = "PROPUESTA SERVICIO DE MANTENIMIENTO INTEGRAL"
                  CR05.Formulas(0) = "Titulo = '" & VAR_TIT & "' "
                  CR05.Formulas(1) = "Subtitulo = '" & VAR_SUBT & "' "
                  CR05.Formulas(2) = "Subtitulo2 = '" & lbl_titulo.Caption & "' "
                  
                  CR05.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CR05.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                  CR05.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                  iResult = CR05.PrintReport
                  If iResult <> 0 Then MsgBox CR05.LastErrorNumber & " : " & CR05.LastErrorString, vbCritical, "Error de impresión"
                  CR05.WindowState = crptMaximized
                  '-----------------------------------------------------------------PAGINA 4
                  CR06.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta4.rpt"
                  CR06.WindowShowPrintSetupBtn = True
                  CR06.WindowShowRefreshBtn = True

                  VAR_TIT = "GERENCIA TECNICA"
                  VAR_SUBT = "PROPUESTA SERVICIO DE MANTENIMIENTO INTEGRAL"
                  CR06.Formulas(0) = "Titulo = '" & VAR_TIT & "' "
                  CR06.Formulas(1) = "Subtitulo = '" & VAR_SUBT & "' "
                  CR06.Formulas(2) = "Subtitulo2 = '" & lbl_titulo.Caption & "' "
                  
                  CR06.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CR06.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                  CR06.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                  iResult = CR06.PrintReport
                  If iResult <> 0 Then MsgBox CR06.LastErrorNumber & " : " & CR06.LastErrorString, vbCritical, "Error de impresión"
                  CR06.WindowState = crptMaximized

                  '-----------------------------------------------------------------
              Else
                  MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet2.Caption, , "Atención"
              End If
            Else
              MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
            End If
        Case "DNREP", "DREPS", "DREPB", "DREPC"            'MANTENIMIENTO CORRECTIVO / REPARACIONES
              If (Ado_datos.Recordset.RecordCount > 0) Then
                  If Ado_detalle2.Recordset.RecordCount > 0 Then
                      'Dim co As New ADODB.Command
                      'CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta.rpt"
                      CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_reparacion.rpt"
                      CR02.WindowShowPrintSetupBtn = True
                      CR02.WindowShowRefreshBtn = True
                      'MsgBox rs.RecordCount
                        CR02.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
                        CR02.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "
                      'Call CREAVISTAF11          'JQA JUN-2008
                      CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                      CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                      CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                      iResult = CR02.PrintReport
                      If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
                      CR02.WindowState = crptMaximized
                  Else
                      MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet2.Caption, , "Atención"
                  End If
                Else
                  MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
                End If

        Case "DNEME"            'EMERGENCIAS
            'dtc_codigo2.Text = "TEC-04" '10
            If (Ado_datos.Recordset.RecordCount > 0) Then
              If Ado_detalle2.Recordset.RecordCount > 0 Then
                  
                  'Dim co As New ADODB.Command
                  'CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_solicitud_cotizacion.rpt"
                  CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta.rpt"
                  CR02.WindowShowPrintSetupBtn = True
                  CR02.WindowShowRefreshBtn = True
                  'MsgBox rs.RecordCount
                    CR02.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
                    CR02.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "
                  'Call CREAVISTAF11          'JQA JUN-2008
                  CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                  CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                  iResult = CR02.PrintReport
                  If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
                  CR02.WindowState = crptMaximized
              Else
                  MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet2.Caption, , "Atención"
              End If
            Else
              MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
            End If
        Case "DNMOD"            'MODERNIZACION
            'dtc_codigo2.Text = "TEC-05" '10
        Case Else
            'dtc_codigo2.Text = "TEC-01"   '3
            If (Ado_datos.Recordset.RecordCount > 0) Then
              If Ado_detalle2.Recordset.RecordCount > 0 Then
                  
                  'Dim co As New ADODB.Command
                  'CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_solicitud_cotizacion.rpt"
                  CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta.rpt"
                  CR02.WindowShowPrintSetupBtn = True
                  CR02.WindowShowRefreshBtn = True
                  'MsgBox rs.RecordCount
                    CR02.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
                    CR02.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "
                  'Call CREAVISTAF11          'JQA JUN-2008
                  CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                  CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                  iResult = CR02.PrintReport
                  If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
                  CR02.WindowState = crptMaximized
              Else
                  MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet2.Caption, , "Atención"
              End If
            Else
              MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
            End If
    End Select
  
End Sub

Private Sub BtnModDetalle_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  If Ado_detalle1.Recordset.RecordCount > 0 Then
    If rs_datos.RecordCount > 0 Then            'And rs_datos!estado_codigo = "REG"
      marca1 = rs_det1.Bookmark
      swnuevo = 2
      fraOpciones.Enabled = False
      FraNavega.Enabled = False
      FraDet1.Enabled = False
      FrmABMDet.Enabled = False
      FraDet2.Enabled = False
      FrmABMDet2.Enabled = False
      Fra_datos.Enabled = False
      VAR_SOL = Ado_datos.Recordset!solicitud_codigo
      Aux = Ado_datos.Recordset!unidad_codigo  'Unidad
      If Aux = "DNEME" Then
          tw_bitacora_emergencia.Txt_campo1.Caption = Aux  'Unidad
          tw_bitacora_emergencia.txt_codigo.Caption = VAR_SOL  'Tramite
          'tw_bitacora_emergencia.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
          'tw_bitacora_emergencia.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
          tw_bitacora_emergencia.Txt_descripcion.Caption = Me.dtc_desc1.Text
          tw_bitacora_emergencia.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
          'tw_bitacora_emergencia.Txt_estado.Caption = "REG"
          'Ado_detalle1.Recordset.AddNew
           
          tw_bitacora_emergencia.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("negocia_forma")
          tw_bitacora_emergencia.DTPfecha1.Value = Me.Ado_detalle1.Recordset("negocia_fecha_real")
          tw_bitacora_emergencia.Txt_campo2.Value = IIf(IsNull(Ado_detalle1.Recordset!negocia_hora_real) Or (Me.Ado_detalle1.Recordset!negocia_hora_real = ":"), "00:00", Me.Ado_detalle1.Recordset!negocia_hora_real)
          tw_bitacora_emergencia.txt_campo6.Value = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_hora_envio), "00:00", Me.Ado_detalle1.Recordset!negocia_hora_envio)
          tw_bitacora_emergencia.txt_campo7.Value = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_hora_llegada), "00:00", Me.Ado_detalle1.Recordset!negocia_hora_llegada)
          tw_bitacora_emergencia.txt_campo8.Value = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_hora_mora), "00:00", Me.Ado_detalle1.Recordset!negocia_hora_mora)
          tw_bitacora_emergencia.txt_campo9.Value = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_hora_salida), "00:00", Me.Ado_detalle1.Recordset!negocia_hora_salida)
          tw_bitacora_emergencia.txt_campo10.Value = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_hora_trabajo), "00:00", Me.Ado_detalle1.Recordset!negocia_hora_trabajo)
          
          tw_bitacora_emergencia.dtc_codigo4.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!tipo_falla), "", Me.Ado_detalle1.Recordset!tipo_falla)
          tw_bitacora_emergencia.dtc_desc4.BoundText = tw_bitacora_emergencia.dtc_codigo4.BoundText
          
          tw_bitacora_emergencia.dtc_codigo5.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!falla_codigo), "", Me.Ado_detalle1.Recordset!falla_codigo)
          tw_bitacora_emergencia.dtc_desc5.BoundText = tw_bitacora_emergencia.dtc_codigo5.BoundText
          
          tw_bitacora_emergencia.Txt_monto1.Text = Me.Ado_detalle1.Recordset("negocia_gasto_estimado")
          tw_bitacora_emergencia.dtc_codigo2.Text = Me.Ado_detalle1.Recordset("beneficiario_codigo")
          tw_bitacora_emergencia.dtc_codigo3.Text = Me.Ado_detalle1.Recordset("beneficiario_codigo_resp")
          tw_bitacora_emergencia.txt_campo3.Text = Me.Ado_detalle1.Recordset("negocia_tarea_realizada")
          tw_bitacora_emergencia.txt_campo4.Text = Me.Ado_detalle1.Recordset("negocia_observaciones")
          tw_bitacora_emergencia.txt_campo5.Text = Me.Ado_detalle1.Recordset("bitacora_cite")
          If swnuevo = 2 Then
              tw_bitacora_emergencia.dtc_desc1.BoundText = tw_bitacora_emergencia.dtc_codigo1.BoundText
              tw_bitacora_emergencia.dtc_desc2.BoundText = tw_bitacora_emergencia.dtc_codigo2.BoundText
              tw_bitacora_emergencia.dtc_desc3.BoundText = tw_bitacora_emergencia.dtc_codigo3.BoundText
    '          If tw_bitacora_emergencia.Txt_campo2 = ":" Then
    '            tw_bitacora_emergencia.HH = "00"    'Left(tw_bitacora_emergencia.Txt_campo2, 2)
    '            tw_bitacora_emergencia.MM = "00"    'Right(tw_bitacora_emergencia.Txt_campo2, 2)
    '          Else
    '            tw_bitacora_emergencia.HH = Left(tw_bitacora_emergencia.Txt_campo2.Text, 2)
    '            tw_bitacora_emergencia.MM = Right(tw_bitacora_emergencia.Txt_campo2.Text, 2)
    '          End If
              
          End If
          tw_bitacora_emergencia.Show vbModal
      Else
          tw_solicitud_bitacora.Txt_campo1.Caption = Aux  'Unidad
          tw_solicitud_bitacora.txt_codigo.Caption = VAR_SOL  'Tramite
          'tw_solicitud_bitacora.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
          'tw_solicitud_bitacora.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
          tw_solicitud_bitacora.Txt_descripcion.Caption = Me.dtc_desc1.Text
          tw_solicitud_bitacora.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
          'tw_solicitud_bitacora.Txt_estado.Caption = "REG"
          'Ado_detalle1.Recordset.AddNew
           
          tw_solicitud_bitacora.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("negocia_forma")
          tw_solicitud_bitacora.DTPfecha1.Value = Me.Ado_detalle1.Recordset("negocia_fecha_real")
          tw_solicitud_bitacora.Txt_campo2.Value = Me.Ado_detalle1.Recordset("negocia_hora_real")
          tw_solicitud_bitacora.Txt_monto1.Text = Me.Ado_detalle1.Recordset("negocia_gasto_estimado")
          tw_solicitud_bitacora.dtc_codigo2.Text = Me.Ado_detalle1.Recordset("beneficiario_codigo")
          tw_solicitud_bitacora.dtc_codigo3.Text = Me.Ado_detalle1.Recordset("beneficiario_codigo_resp")
          tw_solicitud_bitacora.txt_campo3.Text = Me.Ado_detalle1.Recordset("negocia_tarea_realizada")
          tw_solicitud_bitacora.txt_campo4.Text = Me.Ado_detalle1.Recordset("negocia_observaciones")
          tw_solicitud_bitacora.txt_campo5.Text = Me.Ado_detalle1.Recordset("bitacora_cite")
          If swnuevo = 2 Then
              tw_solicitud_bitacora.dtc_desc1.BoundText = tw_solicitud_bitacora.dtc_codigo1.BoundText
              tw_solicitud_bitacora.dtc_desc2.BoundText = tw_solicitud_bitacora.dtc_codigo2.BoundText
              tw_solicitud_bitacora.dtc_desc3.BoundText = tw_solicitud_bitacora.dtc_codigo3.BoundText
              tw_solicitud_bitacora.dtc_desc4.BoundText = tw_solicitud_bitacora.dtc_codigo4.BoundText
    '          tw_solicitud_bitacora.HH = Left(tw_solicitud_bitacora.Txt_campo2.Text, 2)
    '          tw_solicitud_bitacora.MM = Right(tw_solicitud_bitacora.Txt_campo2.Text, 2)
            
          End If
          
          tw_solicitud_bitacora.Show vbModal
      End If
      Call ABRIR_TABLA_DET
      swnuevo = 0
      fraOpciones.Enabled = True
      FraNavega.Enabled = True
      FraDet1.Enabled = True
      FrmABMDet.Enabled = True
      FraDet2.Enabled = True
      FrmABMDet2.Enabled = True
      'Fra_datos.Enabled = True
      Call ABRIR_TABLA_DET
      Call OptFilGral1_Click
      If swnuevo = 1 Then
        rs_det1.Move marca1 - 1
      End If
      swnuevo = 0
    Else
      MsgBox "No se puede Modificar un registro APROBADO o ANULADO, Verifique por favor ...!! ", vbExclamation
    End If
  Else
     MsgBox "No se puede MODIFICAR, el registro No fue identificado o No Existe, Verifique por favor ...", vbExclamation, "Validación de Registro"
  End If
End Sub

Private Sub ModifDetalle()
  If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
    swnuevo = 2
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    FraDet3.Enabled = False
    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False

            If VAR_DET = "30000" Then
                'marca1 = Ado_detalle3.Recordset.Bookmark
                tw_solicitud_bienes3.txt_codigo.Caption = Me.Ado_detalle3.Recordset("solicitud_codigo")  'cod_cabecera
                tw_solicitud_bienes3.Txt_campo1.Caption = Me.Ado_detalle3.Recordset("unidad_codigo")  'Unidad
                tw_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
            
                tw_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes3.txt_campo5.Text = Me.Ado_detalle3.Recordset("bien_codigo")
                
                tw_solicitud_bienes3.txt_campo6.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_descripcion), "-", Me.Ado_detalle3.Recordset!bien_descripcion)
                tw_solicitud_bienes3.txt_campo7.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle3.Recordset!bien_descripcion_anterior)
                tw_solicitud_bienes3.txt_campo8.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!marca_codigo), "S/M", Me.Ado_detalle3.Recordset!marca_codigo)
                tw_solicitud_bienes3.txt_campo9.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!modelo_codigo), "S/M", Me.Ado_detalle3.Recordset!modelo_codigo)
                
                tw_solicitud_bienes3.Txt_campo16.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_cantidad), "1", Me.Ado_detalle3.Recordset!bien_cantidad)
                tw_solicitud_bienes3.txt_campo10.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_precio_venta_base), "0", Me.Ado_detalle3.Recordset!bien_precio_venta_base)
                tw_solicitud_bienes3.txt_campo11.Caption = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_total_venta), "0", Me.Ado_detalle3.Recordset!bien_total_venta)
    
                tw_solicitud_bienes3.Txt_campo14.Text = Me.Ado_detalle3.Recordset("unimed_codigo")
                tw_solicitud_bienes3.Txt_campo15.Text = Me.Ado_detalle3.Recordset("fosa_dimension_frente")

                tw_solicitud_bienes3.lbl_det.Caption = VAR_DET
                tw_solicitud_bienes3.Show vbModal
                'Ado_detalle3.Recordset.Move marca1 - 1
            End If
            If VAR_DET = "39800" Then   'REPUESTOS
                tw_solicitud_bienes5.lbl_det.Caption = VAR_DET
                tw_solicitud_bienes5.txt_codigo.Caption = Me.Ado_detalle5.Recordset("solicitud_codigo")  'cod_cabecera
                tw_solicitud_bienes5.Txt_campo1.Caption = Me.Ado_detalle5.Recordset("unidad_codigo")  'Unidad
                tw_solicitud_bienes5.Txt_descripcion.Caption = Me.dtc_desc1.Text
            
                tw_solicitud_bienes5.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes5.txt_campo5.Text = Me.Ado_detalle5.Recordset("bien_codigo")
                
                tw_solicitud_bienes5.txt_campo6.Text = IIf(IsNull(Me.Ado_detalle5.Recordset!bien_descripcion), "-", Me.Ado_detalle5.Recordset!bien_descripcion)
                tw_solicitud_bienes5.txt_campo7.Text = IIf(IsNull(Me.Ado_detalle5.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle5.Recordset!bien_descripcion_anterior)
                tw_solicitud_bienes5.txt_campo8.Text = Me.Ado_detalle5.Recordset("marca_codigo")
                tw_solicitud_bienes5.txt_campo9.Text = Me.Ado_detalle5.Recordset("modelo_codigo")
                
                tw_solicitud_bienes5.Txt_campo16.Text = Me.Ado_detalle5.Recordset("bien_cantidad")
                tw_solicitud_bienes5.txt_campo10.Text = Me.Ado_detalle5.Recordset("bien_precio_venta_base")
                tw_solicitud_bienes5.txt_campo11.Text = Me.Ado_detalle5.Recordset("bien_total_venta")
                
                tw_solicitud_bienes5.Txt_campo14.Text = Me.Ado_detalle5.Recordset("unimed_codigo")
                tw_solicitud_bienes5.Txt_campo15.Text = Me.Ado_detalle5.Recordset("fosa_dimension_frente")
                tw_solicitud_bienes5.dtc_codigo2.BoundText = Me.Ado_detalle5.Recordset("unimed_codigo")
                tw_solicitud_bienes5.dtc_desc2.BoundText = Me.Ado_detalle5.Recordset("unimed_codigo")
                GlExtension = Ado_detalle2.Recordset!bien_codigo
                tw_solicitud_bienes5.Show vbModal
            End If
            If VAR_DET = "34800" Then
                 tw_solicitud_bienes6.lbl_det.Caption = VAR_DET
                tw_solicitud_bienes6.txt_codigo.Caption = Me.Ado_detalle6.Recordset("solicitud_codigo")  'cod_cabecera
                tw_solicitud_bienes6.Txt_campo1.Caption = Me.Ado_detalle6.Recordset("unidad_codigo")  'Unidad
                tw_solicitud_bienes6.Txt_descripcion.Caption = Me.dtc_desc1.Text
            
                tw_solicitud_bienes6.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes6.txt_campo5.Text = Me.Ado_detalle6.Recordset("bien_codigo")
                
                'tw_solicitud_bienes6.Txt_campo6.Text = IIf(IsNull(Me.Ado_detalle6.Recordset!bien_descripcion), "-", Me.Ado_detalle3.Recordset!bien_descripcion)
'                tw_solicitud_bienes6.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle6.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle3.Recordset!bien_descripcion_anterior)
                tw_solicitud_bienes6.txt_campo8.Text = Me.Ado_detalle6.Recordset("marca_codigo")
                tw_solicitud_bienes6.txt_campo9.Text = Me.Ado_detalle6.Recordset("modelo_codigo")
                
                tw_solicitud_bienes6.Txt_campo16.Text = Me.Ado_detalle6.Recordset("bien_cantidad")
                tw_solicitud_bienes6.txt_campo10.Text = Me.Ado_detalle6.Recordset("bien_precio_venta_base")
                tw_solicitud_bienes6.txt_campo11.Caption = Me.Ado_detalle6.Recordset("bien_total_venta")
                
                tw_solicitud_bienes6.Txt_campo14.Text = Me.Ado_detalle6.Recordset("unimed_codigo")
                tw_solicitud_bienes6.Txt_campo15.Text = Me.Ado_detalle6.Recordset("fosa_dimension_frente")
                
                tw_solicitud_bienes6.lbl_det.Caption = VAR_DET
                tw_solicitud_bienes6.Show vbModal
            End If

            If VAR_DET = "24300" Then
                tw_solicitud_bienes7.txt_codigo.Caption = Me.Ado_detalle7.Recordset("solicitud_codigo")  'cod_cabecera
                tw_solicitud_bienes7.Txt_campo1.Caption = Me.Ado_detalle7.Recordset("unidad_codigo")  'Unidad
                tw_solicitud_bienes7.Txt_descripcion.Caption = Me.dtc_desc1.Text
            
                tw_solicitud_bienes7.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes7.txt_campo5.Text = Me.Ado_detalle7.Recordset("bien_codigo")
                
                tw_solicitud_bienes7.txt_campo6.Text = IIf(IsNull(Me.Ado_detalle7.Recordset!bien_descripcion), "-", Me.Ado_detalle7.Recordset!bien_descripcion)
                tw_solicitud_bienes7.txt_campo7.Text = IIf(IsNull(Me.Ado_detalle7.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle7.Recordset!bien_descripcion_anterior)
                tw_solicitud_bienes7.txt_campo8.Text = Me.Ado_detalle7.Recordset("marca_codigo")
                tw_solicitud_bienes7.txt_campo9.Text = Me.Ado_detalle7.Recordset("modelo_codigo")
                
                tw_solicitud_bienes7.Txt_campo16.Text = Me.Ado_detalle7.Recordset("bien_cantidad")
                tw_solicitud_bienes7.txt_campo10.Text = Me.Ado_detalle7.Recordset("bien_precio_venta_base")
                tw_solicitud_bienes7.txt_campo11.Caption = Me.Ado_detalle7.Recordset("bien_total_venta")
                
                tw_solicitud_bienes7.Txt_campo14.Text = Me.Ado_detalle7.Recordset("unimed_codigo")
                tw_solicitud_bienes7.Txt_campo15.Text = Me.Ado_detalle7.Recordset("fosa_dimension_frente")
                
                tw_solicitud_bienes7.lbl_det.Caption = VAR_DET
                tw_solicitud_bienes7.Show vbModal
            End If
'    End Select
    
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    FraDet3.Enabled = True
    FrmABMDet3.Enabled = True
'    Fra_datos.Enabled = True
    Call ABRIR_TABLA_DET
'    Ado_detalle3.Recordset.Move marca1 - 1
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If

End Sub

'Private Sub BtnModificar_Click()
'  On Error GoTo EditErr
'  If Ado_datos.Recordset.RecordCount > 0 Then
''  lblStatus.Caption = "Modificar registro"
'    If Ado_datos.Recordset!estado_codigo = "REG" Then
'        'marca1 = Ado_datos.Recordset.Bookmark
'        Fra_datos.Enabled = True
'        fraOpciones.Visible = False
'        FraGrabarCancelar.Visible = True
'        dg_datos.Enabled = False
'        VAR_SW = "MOD"
'    '    dtc_desc1.Visible = False
'    '    lbl_aux1.Visible = True
'    '    lbl_aux1.Caption = dtc_desc1.Text
'        dtc_desc4.SetFocus
'    '    BtnVer.Visible = True
''        dtc_codigo9.Enabled = False
'        'Call OptFilGral1_Click
'        'Ado_datos.Recordset.Move marca1 - 1
'        Select Case parametro
'            Case "DVTA"             'INI COMERCIAL
'                dtc_codigo2.Text = "COM-01"   '3
'            Case "COMEX"            'INI COMEX
'                dtc_codigo2.Text = "CMX-01"   '3
'            Case "DNINS"            'INI GRABA INSTALACIONES
'                dtc_codigo2.Text = "COM-03" '4
'            Case "DNAJS"            'AJUSTE
'                dtc_codigo2.Text = "COM-04" '5
'            Case "DNMAN", "DMANB", "DMANS", "DMANC"            'MANTENIMIENTO PREVENTIVO
'                dtc_codigo2.Text = "TEC-02" '10
'            Case "DNREP", "DREPB", "DREPS", "DREPC"         'MANTENIMIENTO CORRECTIVO / REPARACIONES
'                dtc_codigo2.Text = "TEC-03" '10
'            Case "DNEME"           'EMERGENCIAS
'                dtc_codigo2.Text = "TEC-04" '10
'            Case "DNMOD"            'MODERNIZACION
'                dtc_codigo2.Text = "TEC-05" '10
'            Case Else
'                dtc_codigo2.Text = "TEC-01"   '3
'        End Select
'    Else
'      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
'    End If
'  Else
'        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
'  Exit Sub
'
'EditErr:
'  MsgBox Err.Description
'End Sub

Private Sub BtnSalir_Click()
'  If glPersOtro = "O" Then
'    frmmo_pacientes.Dtc_ocupac = rs_datos!ocup_codigo
'    frmmo_pacientes.Dtc_OcupacDes = rs_datos!ocup_descripcion
'  End If
'  glPersOtro = "N"
  Unload Me
End Sub

Private Sub BtnVer_Click()
  On Error GoTo QError
  If rs_datos!estado_codigo = "APR" Then
    Dim ARCH_FOTO As String
    Dim SW0 As String
    Select Case Left(Trim(Ado_datos.Recordset("edif_codigo")), 1)
        Case "1"    'CHQ
            VAR_DPTO = "CHQ"
        Case "2"    'LPZ
            VAR_DPTO = "LPZ"
        Case "3"    'CBB
            VAR_DPTO = "CBB"
        Case "4"    'SCZ
            VAR_DPTO = "SCZ"
        Case "5"    'PTS
            VAR_DPTO = "PTS"
        Case "6"    'ORU
            VAR_DPTO = "ORU"
        Case "7"    'TJA
            VAR_DPTO = "TJA"
        Case "8"    'BEN
            VAR_DPTO = "BEN"
        Case "9"    'PDO
            VAR_DPTO = "PDO"
    End Select
    If Ado_datos.Recordset!archivo_respaldo_cargado = "N" Then
      'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!edif_tipo) & "\" & Trim(Ado_datos.Recordset!negocia_codigo) & "\"
      NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "DED2"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'      Else
         e = NombreCarpeta
'      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
      SW0 = 1
    Else
      'MsgBox ""
      'negocia_codigo, unidad_codigo, negocia_fecha_inicio as fecha1, negocia_descripcion, estado_codigo, fecha_registro, usr_codigo, solicitud_tipo as codigo2, edif_codigo as codigo3, beneficiario_codigo as codigo4, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero As campo1, poa_codigo As codigo10, hora_registro, ges_gestion, archivo_respaldo, archivo_respaldo_cargado
      sino = MsgBox("El archivo ya existe, elija: <SI> para Volver a Cargarlo. <NO> para Visualizarlo. ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!edif_tipo) & "\" & Trim(Ado_datos.Recordset!negocia_codigo) & "\"
          NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "DED2"
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'          Else
            e = NombreCarpeta
'          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
          SW0 = 1
      Else
        SW0 = 0
        'e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!codigo_beneficiario) & "\LICENCIAS\" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
        e = ShellExecute(0, vbNullString, App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\" & Trim(Ado_datos.Recordset("archivo_respaldo")), vbNullString, vbNullString, vbNormalFocus)
      End If
    End If
    '    If SW0 = 1 Then
    '    '    If GlServidor = "SRVPRO" Then
    '    '        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
    '    '    Else
    '            'ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset!edif_tipo) + "\" + Trim(Ado_datos.Recordset!edif_codigo)
    '            ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset!edif_tipo) + "\" + Trim(Ado_datos.Recordset!edif_codigo) + ".JPG"
    '    '    End If
    '        'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + Ado_datos.Recordset!codigo_beneficiario + "\" + Ado_datos.Recordset("codigo_beneficiario") + "-FOTO.JPG"
    '        CodBien = Ado_datos.Recordset!edif_codigo
    '        If Guardar_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo= '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
    '            MsgBox "Se cargo la Imagen Correctamente !!"
    '        Else
    '            MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    '        End If
    '    Else
    '        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
    '        Image2 = Img_Foto
    '    End If
  Else
       MsgBox "No se puede Guardar el documento PDF, debe APROBAR previamente el registro ...", vbExclamation, "Validación de Registro"
  End If
QError:
    ' Manejo de errores
    If Err.Number > 0 Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
    '    db.RollbackTrans
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub dtc_aux1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_aux1.BoundText
    dtc_codigo1.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
    dtc_desc10.BoundText = dtc_codigo10.BoundText
End Sub

Private Sub dtc_codigo11_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

'Private Sub dtc_codigo5_Click(Area As Integer)
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'End Sub

'Private Sub dtc_codigo6_Click(Area As Integer)
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'End Sub

'Private Sub dtc_codigo7_Click(Area As Integer)
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'End Sub

'Private Sub dtc_codigo8_Click(Area As Integer)
'    dtc_desc8.BoundText = dtc_codigo8.BoundText
'End Sub

'Private Sub dtc_codigo9_Click(Area As Integer)
'    dtc_desc9.BoundText = dtc_codigo9.BoundText
'End Sub

'Private Sub dtc_codigo9_LostFocus()
''  If VAR_SW = "ADD" Then
''    Set rs_aux2 = New ADODB.Recordset
''    SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9.Text & "'  "
''    rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
''    If rs_aux2.RecordCount > 0 Then
''        rs_aux2!correl_doc = rs_aux2!correl_doc + 1
''        txt_campo1.Caption = rs_aux2!correl_doc
''        rs_aux2.Update
''    End If
''  End If
'  txt_aux9.Text = dtc_desc9.Text
'End Sub

'Private Sub dtc_desc5_Click(Area As Integer)
'    dtc_codigo5.BoundText = dtc_desc5.BoundText
''    Call pnivel5(dtc_codigo5.BoundText)
''    dtc_desc6.Enabled = True
'End Sub
   
'Private Sub pnivel5(codigo5 As String)
'   'Dim strConsultaF As String
'   'strConsultaF = "select * from gc_proceso_nivel2 where proceso_codigo = '" & codigo5 & "'"
'
'   Set dtc_codigo6.RowSource = Nothing
'   'Set dtc_codigo6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_codigo6.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel2 '" & codigo5 & "' ")
'   dtc_codigo6.ReFill
'   dtc_codigo6.BoundText = Empty
'
'   Set dtc_desc6.RowSource = Nothing
'   'Set dtc_desc6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_desc6.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel2 '" & codigo5 & "' ")
'   dtc_desc6.ReFill
'   dtc_desc6.BoundText = Empty
'End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    'Call pnivel1(dtc_codigo1.BoundText)
    'dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub
   
Private Sub pnivel1(codigo1 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
   
   Set dtc_codigo10.RowSource = Nothing
'   Set dtc_codigo10.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_codigo10.ReFill
   dtc_codigo10.BoundText = Empty
   
   Set dtc_desc10.RowSource = Nothing
   'Set dtc_desc10.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_desc10.ReFill
   dtc_desc10.BoundText = Empty
End Sub
  
Private Sub pnivel11(codigo2 As String)
    Select Case codigo2
        Case "DVTA"             'INI COMERCIAL
            dtc_codigo2.Text = "COM-01"   '3
        Case "COMEX"            'INI COMEX
            dtc_codigo2.Text = "CMX-01"   '3
        Case "DNINS"            'INI GRABA INSTALACIONES
            dtc_codigo2.Text = "COM-03" '4
        Case "DNAJS"            'AJUSTE
            dtc_codigo2.Text = "COM-04" '5
        Case "DNMAN"            'MANTENIMIENTO PREVENTIVO
            dtc_codigo2.Text = "TEC-02" '10
        Case "DNREP"            'MANTENIMIENTO CORRECTIVO / REPARACIONES
            dtc_codigo2.Text = "TEC-03" '10
        Case "DNEME"            'EMERGENCIAS
            dtc_codigo2.Text = "TEC-04" '10
        Case "DNMOD"            'MODERNIZACION
            dtc_codigo2.Text = "TEC-05" '10
        Case Else
            dtc_codigo2.Text = "TEC-01"   '3
    End Select
    
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
'    Dim strConsultaF As String
'   'strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
'   strConsultaF = "Select * from gv_personal_contratado where unidad_codigo = '" & codigo1 & "' order by beneficiario_denominacion"
'
'   Set dtc_codigo11.RowSource = Nothing
'   Set dtc_codigo11.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_codigo11.ReFill
'   dtc_codigo11.BoundText = Empty
'
'   Set dtc_desc11.RowSource = Nothing
'   Set dtc_desc11.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_desc11.ReFill
'   dtc_desc11.BoundText = Empty
End Sub

'Private Sub dtc_desc1_LostFocus()
''    dtc_codigo5.Text = dtc_aux1.Text
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
'    Call pnivel5(dtc_codigo5.BoundText)
'    dtc_desc6.Enabled = True
'End Sub

Private Sub dtc_desc10_Click(Area As Integer)
    dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub

Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub
 
Private Sub dtc_desc3_LostFocus()
    dtc_codigo4.Text = dtc_aux3.Text
    'Txt_descripcion.Text = lbl_titulo + " - Edificio: " + dtc_desc3.Text
    Select Case parametro
        Case "DNMAN", "DMANS", "DMANB", "DMANC"
            Txt_descripcion.Text = "Propuesta de Servicio de MANTENIMIENTO INTEGRAL. Edificio: " + dtc_desc3.Text + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
        Case "DNREP", "DREPS", "DREPB", "DREPC"
            Txt_descripcion.Text = "Servicio de REPARACIONES. Edificio: " + dtc_desc3.Text
        Case "DNMOD", "DMODS", "DMODB", "DMODC"
            Txt_descripcion.Text = "Propuesta de MODERNIZACION de equipos. Edificio: " + dtc_desc3.Text + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
        Case "DNINS", "DINSS", "DINSB", "DINSC"
            Txt_descripcion.Text = "Servicio de INSTALACION de equipos. Edificio: " + dtc_desc3.Text
        Case "DNEME", "DEMES", "DEMEB", "DEMEC"
            Txt_descripcion.Text = "Atención de EMERGENCIAS. Edificio: " + dtc_desc3.Text + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
            Set rs_aux9 = New ADODB.Recordset
            If rs_aux9.State = 1 Then rs_aux9.Close
            rs_aux9.Open "Select * from tv_zona_piloto_edif_resp ", db, adOpenStatic
            If rs_aux9.RecordCount > 0 Then
                dtc_codigo11.Text = rs_aux9!beneficiario_codigo
                dtc_desc11.BoundText = dtc_codigo11.BoundText
                'dtc_desc11.Text = rs_aux9!beneficiario_denominacion
            Else
                dtc_codigo11.Text = "4245046"
                dtc_desc11.BoundText = dtc_codigo11.BoundText
                'dtc_desc11.Text = "ORAQUENI QUITO JAVIER"
            End If
        Case Else
    End Select
    dtc_desc4.BoundText = dtc_codigo4.BoundText
  
    'Call pnivel1(dtc_codigo1.BoundText)
    'dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

'Private Sub dtc_desc6_Click(Area As Integer)
'    dtc_codigo6.BoundText = dtc_desc6.BoundText
''    Call pnivel6(dtc_codigo6.BoundText)
''    dtc_desc7.Enabled = True
'End Sub
  
'Private Sub pnivel6(codigo6 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from gc_proceso_nivel3 where subproceso_codigo = '" & codigo6 & "'"
'
'   Set dtc_codigo7.RowSource = Nothing
'   Set dtc_codigo7.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo7.RowSource = db.Execute("EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
'   dtc_codigo7.ReFill
'   dtc_codigo7.BoundText = Empty
'
'   Set dtc_desc7.RowSource = Nothing
'   Set dtc_desc7.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo7.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
'   dtc_desc7.ReFill
'   dtc_desc7.BoundText = Empty
'End Sub

'Private Sub dtc_desc7_Click(Area As Integer)
'    dtc_codigo7.BoundText = dtc_desc7.BoundText
'End Sub

'Private Sub dtc_desc8_Click(Area As Integer)
'    dtc_codigo8.BoundText = dtc_desc8.BoundText
'    Call pnivel8(dtc_codigo8.BoundText)
'    'dtc_desc9.Enabled = True
'    dtc_codigo9.Enabled = True
'End Sub
   
'Private Sub pnivel8(codigo8 As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_documentos_respaldo where clasif_codigo = '" & codigo8 & "'"
'
'   Set dtc_codigo9.RowSource = Nothing
'   Set dtc_codigo9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo9.ReFill
'   dtc_codigo9.BoundText = Empty
'
'   Set dtc_desc9.RowSource = Nothing
'   Set dtc_desc9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc9.ReFill
'   dtc_desc9.BoundText = Empty
'End Sub

'Private Sub dtc_desc9_Click(Area As Integer)
'    dtc_codigo9.BoundText = dtc_codigo9.BoundText
'End Sub

Private Sub Form_Load()
    buscados = 0
    swnuevo = 0
    VAR_SW = ""
    Set rs_aux8 = New ADODB.Recordset
    If rs_aux8.State = 1 Then rs_aux8.Close
    rs_aux8.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux8.RecordCount > 0 Then
        usuario2 = rs_aux8!beneficiario_codigo
        VAR_DA = rs_aux8!da_codigo
        VAR_DPTOC = IIf(IsNull(rs_aux8!depto_codigo), "2", rs_aux8!depto_codigo)
    Else
        usuario2 = "3361040"
        VAR_DA = "1.3"
        VAR_DPTOC = "2"
    End If
    VAR_UORIGEN = Aux
    If Aux = "DNMAN" Then
        Select Case VAR_DA
            Case "1.8"    'Cochabamba
                Aux = "DMANB"
                'VAR_DPTOC = "3"
            Case "1.7"    'Santa Cruz
                Aux = "DMANS"
                'VAR_DPTOC = "7"
            Case "1.3"    'La Paz - Tecnico
                Aux = "DNMAN"
                'VAR_DPTOC = "2"
            Case "1.9"    ' Chuquisaca
                Aux = "DMANC"
                'VAR_DPTOC = "1"
            Case Else    ' TODO
                Aux = "DNMAN"
                'VAR_DPTOC = "0"
         End Select
         VAR_TIPO = 6
     End If
     If Aux = "DNREP" Then
        Select Case VAR_DA
            Case "1.8"    'Cochabamba
                Aux = "DREPB"
                'VAR_DPTOC = "3"
            Case "1.7"    'Santa Cruz
                Aux = "DREPS"
                'VAR_DPTOC = "7"
            Case "1.3"    'La Paz - Tecnico
                Aux = "DNREP"
                'VAR_DPTOC = "2"
            Case "1.9"    ' Chuquisaca
                Aux = "DREPC"
                'VAR_DPTOC = "1"
            Case "0"    ' TODO
                Aux = "DNREP"
                'VAR_DPTOC = "0"
         End Select
         VAR_TIPO = 7
     End If
     If Aux = "DNINS" Then
        Select Case VAR_DA
            Case "1.8"    'Cochabamba
                Aux = "DINSB"
                'VAR_DPTOC = "3"
            Case "1.7"    'Santa Cruz
                Aux = "DINSS"
                'VAR_DPTOC = "7"
            Case "1.3", "1.2"    'La Paz - Tecnico
                Aux = "DNINS"
                'VAR_DPTOC = "2"
            Case "1.9"    ' Chuquisaca
                Aux = "DINSC"
                'VAR_DPTOC = "1"
            Case Else    ' TODO
                Aux = "DNINS"
                'VAR_DPTOC = "0"
         End Select
         VAR_TIPO = 4
     End If
    If Aux = "DNEME" Then
        Select Case VAR_DA
            Case "1.8"    'Cochabamba
                Aux = "DMANB"
                'VAR_DPTOC = "3"
            Case "1.7"    'Santa Cruz
                Aux = "DMANS"
                'VAR_DPTOC = "7"
            Case "1.3"    'La Paz - Tecnico
                Aux = "DNEME"
                'VAR_DPTOC = "2"
            Case "1.9"    ' Chuquisaca
                Aux = "DMANC"
                'VAR_DPTOC = "1"
            Case "0"    ' TODO
                Aux = "DNEME"
                'VAR_DPTOC = "2"
         End Select
         VAR_TIPO = 8
     End If
    
    parametro = Aux
    db.Execute "UPDATE ao_solicitud SET ao_solicitud.observacion_proy = gc_edificaciones.edif_descripcion from ao_solicitud inner join gc_edificaciones on ao_solicitud.edif_codigo = gc_edificaciones.edif_codigo WHERE (ao_solicitud.unidad_codigo = '" & parametro & "')"
    'db.Execute "UPDATE ao_solicitud SET ao_solicitud.observacion_proy = gc_edificaciones.edif_descripcion from ao_solicitud inner join gc_edificaciones on ao_solicitud.edif_codigo = gc_edificaciones.edif_codigo where ao_solicitud.edif_codigo <> '0' and ao_solicitud.observacion_proy is null"
    'parametro = "estado_codigo" + " = " + "'REG'"
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    'JQA 2014-JUL-14
    'db.Execute (" EXEC gp_actualiza_beneficiario_edif ")
'    lbl_aux1.Visible = False
    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_unidad_ejecutora
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    'gc_tipo_solicitud
    
    'gc_proceso_nivel2
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    If parametro = "DNINS" Or parametro = "DNAJS" Then
        rs_datos2.Open "Select * from gc_proceso_nivel2 WHERE proceso_codigo = 'COM' order by subproceso_descripcion", db, adOpenStatic
    Else
        rs_datos2.Open "Select * from gc_proceso_nivel2 WHERE proceso_codigo = 'TEC' order by subproceso_descripcion", db, adOpenStatic
    End If
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    'gc_edificaciones
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_edificaciones order by edif_descripcion", db, adOpenStatic
    'rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'gc_beneficiario (Personas Nat. y Juridicas / Clientes, Proveedores, etc.)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
'    Set rs_datos5 = New ADODB.Recordset
'    If rs_datos5.State = 1 Then rs_datos5.Close
'    'rs_datos5.Open "Select * from gc_proceso_nivel1 order by proceso_descripcion", db, adOpenStatic
'    rs_datos5.Open "gp_listar_apr_gc_proceso_nivel1", db, adOpenStatic
'    Set Ado_datos5.Recordset = rs_datos5
''    dtc_desc5.BoundText = dtc_codigo5.BoundText

'    Set rs_datos6 = New ADODB.Recordset
'    If rs_datos6.State = 1 Then rs_datos6.Close
'    rs_datos6.Open "Select * from gc_proceso_nivel2 WHERE proceso_codigo = 'TEC' order by subproceso_descripcion", db, adOpenStatic
'    'rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
'    Set Ado_datos6.Recordset = rs_datos6
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'
'    Set rs_datos7 = New ADODB.Recordset
'    If rs_datos7.State = 1 Then rs_datos7.Close
'    'rs_datos7.Open "Select * from gc_proceso_nivel3 order by etapa_descripcion", db, adOpenStatic
'    rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
'    Set Ado_datos7.Recordset = rs_datos7
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'
'    Set rs_datos8 = New ADODB.Recordset
'    If rs_datos8.State = 1 Then rs_datos8.Close
'    'rs_datos8.Open "Select * from gc_documentos_clasificacion order by clasif_codigo", db, adOpenStatic
'    rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
'    Set Ado_datos8.Recordset = rs_datos8
''    dtc_desc8.BoundText = dtc_codigo8.BoundText
    
'    'gc_documentos_respaldo
'    Set rs_datos9 = New ADODB.Recordset
'    If rs_datos9.State = 1 Then rs_datos9.Close
'    'rs_datos9.Open "Select * from gc_documentos_respaldo order by doc_codigo", db, adOpenStatic
'    rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
'    Set Ado_datos9.Recordset = rs_datos9
'    dtc_desc9.BoundText = dtc_codigo9.BoundText
    
    'gc_ContratoTipo
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from gc_ContratoTipo WHERE solicitud_tipo = " & VAR_TIPO & " order by TipoContratoCodigo", db, adOpenStatic
    'rs_datos10.Open "pp_listar_apr_pc_poa_actividad", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText
    
    'gc_beneficiario (Personal CGI)
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub ABRIR_TABLA_DET()
    'BITACORA
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    'rs_det1.Open "select * from ao_solicitud_bitacora where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    rs_det1.Open "select * from ao_solicitud_bitacora where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    If rs_det1.RecordCount > 0 Then
        dg_det1.Visible = True
        Set dg_det1.DataSource = Ado_detalle1.Recordset
    Else
        dg_det1.Visible = False
        'Set Ado_detalle1.Recordset = rsNada
        Set dg_det1.DataSource = rsNada
    End If
    
    'EQUIPOS par_codigo = '43340'
    Set rs_det2 = New ADODB.Recordset
    If rs_det2.State = 1 Then rs_det2.Close
    'rs_det2.Open "select * from av_solicitud_bienes where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & " and (par_codigo = '43340' ) ", db, adOpenKeyset, adLockOptimistic, adCmdText       'and estado_codigo = 'APR'
    rs_det2.Open "select * from av_solicitud_bienes where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & " and (par_codigo = '43340' ) ", db, adOpenKeyset, adLockOptimistic, adCmdText       'and estado_codigo = 'APR'
    Set Ado_detalle2.Recordset = rs_det2
    If rs_det2.RecordCount > 0 Then
        dg_det2.Visible = True
        Set dg_det2.DataSource = Ado_detalle2.Recordset
    Else
        dg_det2.Visible = False
        'Set Ado_detalle2.Recordset = rsNada
        Set dg_det2.DataSource = rsNada
    End If
    
    'INSUMOS y materiales par_codigo = '43340'
'    Set rs_det3 = New Recordset
'    If rs_det3.State = 1 Then rs_det3.Close
'    'rs_det3.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (grupo_codigo = '30000' and (par_codigo <> '39810' and par_codigo <> '39820' and par_codigo <> '34800'))   ", db, adOpenKeyset, adLockOptimistic, adCmdText        'and estado_codigo = 'APR'
'    rs_det3.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & "  and (grupo_codigo = '30000' and (par_codigo <> '39810' and par_codigo <> '39820' and par_codigo <> '34800'))   ", db, adOpenKeyset, adLockOptimistic, adCmdText        'and estado_codigo = 'APR'
'    Set Ado_detalle3.Recordset = rs_det3.DataSource
'    If rs_det3.RecordCount > 0 Then
'        dg_det3.Visible = True
'        Set dg_det3.DataSource = Ado_detalle3.Recordset
'    Else
'        dg_det3.Visible = False
'        Set dg_det3.DataSource = rsNada
'    End If

    'REPUESTOS par_codigo = '39800'
'    Set rs_det5 = New Recordset
'    If rs_det5.State = 1 Then rs_det5.Close
'    'rs_det5.Open "select * from av_solicitud_bienes3 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  AND (almacen_tipo = 'R')  ", db, adOpenKeyset, adLockOptimistic, adCmdText        'and estado_codigo = 'APR'
'    rs_det5.Open "select * from av_solicitud_bienes3 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & "  AND (almacen_tipo = 'R')  ", db, adOpenKeyset, adLockOptimistic, adCmdText        'and estado_codigo = 'APR'
'    Set Ado_detalle5.Recordset = rs_det5.DataSource
'    If rs_det5.RecordCount > 0 Then
'        dg_det5.Visible = True
'        Set dg_det5.DataSource = Ado_detalle5.Recordset
'    Else
'        dg_det5.Visible = False
'        Set dg_det5.DataSource = rsNada
'    End If

    'HERRAMIENTAS par_codigo = '43700' - par_codigo = '34800'
'    Set rs_det6 = New Recordset
'    If rs_det6.State = 1 Then rs_det6.Close
'    'rs_det6.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (par_codigo = '43700' or par_codigo = '34800')  ", db, adOpenKeyset, adLockOptimistic, adCmdText     'and estado_codigo = 'APR'
'    rs_det6.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & "  and (par_codigo = '43700' or par_codigo = '34800')  ", db, adOpenKeyset, adLockOptimistic, adCmdText     'and estado_codigo = 'APR'
'    Set Ado_detalle6.Recordset = rs_det6.DataSource
'    If rs_det6.RecordCount > 0 Then
'        dg_det6.Visible = True
'        Set dg_det6.DataSource = Ado_detalle6.Recordset
'    Else
'        dg_det6.Visible = False
'        Set dg_det6.DataSource = rsNada
'    End If
    
'    Set rs_det4 = New Recordset
'    If rs_det4.State = 1 Then rs_det4.Close
'    'rs_det4.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    rs_det4.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    Set Ado_detalle4.Recordset = rs_det4.DataSource
'    Set dg_det4.DataSource = Ado_detalle4.Recordset
    
    'REPUESTOS par_codigo = '24000'
'    Set rs_det7 = New Recordset
'    If rs_det7.State = 1 Then rs_det7.Close
'    'rs_det7.Open "select * from av_solicitud_bienes7 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & " and par_codigo = '24300'   ", db, adOpenKeyset, adLockOptimistic, adCmdText      '
'    rs_det7.Open "select * from av_solicitud_bienes7 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & " and par_codigo = '24300'   ", db, adOpenKeyset, adLockOptimistic, adCmdText      '
'    Set Ado_detalle7.Recordset = rs_det7.DataSource
'    If rs_det7.RecordCount > 0 Then
'        dg_det7.Visible = True
'        Set dg_det7.DataSource = Ado_detalle7.Recordset
'    Else
'        dg_det7.Visible = False
'        Set dg_det7.DataSource = rsNada
'    End If
End Sub

Private Sub ABRIR_TABLA_AUX2()
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & VAR_UNI & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If Ado_datos.Recordset.RecordCount > 0 Then
     If OptFilGral1.Value = True Then
        TxtContrato.Text = Ado_datos.Recordset!venta_monto_total_bs
        DTPicker1.Value = Ado_datos.Recordset!venta_fecha_inicio
        DTPicker2.Value = Ado_datos.Recordset!venta_fecha_fin
     Else
    
     End If
     VAR_SOL = Ado_datos.Recordset!solicitud_codigo
     VAR_UNI = Ado_datos.Recordset!unidad_codigo
     If buscados = 0 Then
        OptFilGral1.Visible = True
        OptFilGral2.Visible = True
     Else
        OptFilGral1.Visible = False
        OptFilGral2.Visible = False
     End If
'     If OptFilGral1.Enabled = True Then
'        TxtContrato.Text = Ado_datos.Recordset!venta_monto_total_bs
'        DTPicker1.Value = Ado_datos.Recordset!venta_fecha_inicio
'        DTPicker2.Value = Ado_datos.Recordset!venta_fecha_fin
'     Else
'
'     End If
    'Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
    ' <-- Inicio                Identificación del Cliente                Fin -->   'esto es de Caption
    If VAR_SW <> "ADD" Then
        'Select Case rs_datos!solicitud_tipo     'dtc_codigo2.Text
'        If VAR_SOL = 0 Then
        'If VAR_SW <> "" Then
        If Not (Ado_datos.Recordset.EOF) Then   'And Not (Ado_datos.Recordset.BOF)
            VAR_SOL = Ado_datos.Recordset!solicitud_codigo
            VAR_UNI = Ado_datos.Recordset!unidad_codigo
        
            
        End If
        Call ABRIR_TABLA_DET
        'VAR_SOL = Ado_datos.Recordset!solicitud_codigo
        Call ABRIR_TABLA_AUX2
    Else
        'Set rs_det1 = New ADODB.Recordset
        'Set dg_det2.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
    FraDet1.Caption = "BITÁCORA " + dtc_desc2.Text
    'FraDet1.Caption = "BITÁCORA DE " + lbl_titulo
'    txt_aux9.Text = dtc_desc9.Text
    If Not (Ado_datos.Recordset.EOF) Then
        If Ado_datos.Recordset!estado_codigo = "APR" Then
            FrmABMDet2.Visible = False
'            FrmABMDet3.Visible = False
'            BtnAprobar.Visible = False
'            If glusuario = "ADMIN" Or glusuario = "ADMINSTC" Or glusuario = "ADMINCBB" Or glusuario = "ADMINCHQ" Then
'                BtnDesAprobar.Visible = True
'            Else
'                BtnDesAprobar.Visible = False
'            End If
        Else
'            If Ado_datos.Recordset!estado_codigo = "REG" Then
''                BtnAprobar.Visible = True
'                BtnDesAprobar.Visible = False
'            End If
            FrmABMDet2.Visible = True
'            FrmABMDet3.Visible = True
        End If
    End If
  Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det1.DataSource = rsNada
        Set dg_det2.DataSource = rsNada
        'Set dg_det3.DataSource = rsNada
        'Set dg_det5.DataSource = rsNada
        'Set dg_det6.DataSource = rsNada
        'Set dg_det7.DataSource = rsNada
     If buscados = 0 Then
        OptFilGral1.Visible = True
        OptFilGral2.Visible = True
     Else
        OptFilGral1.Visible = False
        OptFilGral2.Visible = False
     End If
  End If
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

'Private Sub BtnAñadir_Click()
'  On Error GoTo AddErr
'    VAR_SW = "ADD"
'    'lblStatus.Caption = "Agregar registro"
'    Fra_datos.Enabled = True
'    fraOpciones.Visible = False
'    FraGrabarCancelar.Visible = True
'    dg_datos.Enabled = False
'    dg_det1.Visible = False
'    dg_det2.Visible = False
'    dg_det3.Visible = False
'    dg_det5.Visible = False
'    dg_det6.Visible = False
'    dg_det7.Visible = False
'    FrmABMDet2.Enabled = False
'    FrmABMDet5.Enabled = False
'    FrmABMDet3.Enabled = False
'    FrmABMDet6.Enabled = False
'    FrmABMDet.Enabled = False
'    FrmABMDet7.Enabled = False
'    'txt_codigo.Enabled = False
''    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
''    rs_datos.AddNew
'    Ado_datos.Recordset.AddNew
'    dtc_desc11.SetFocus
'    'dtc_desc1.BackColor = &H80000005
'    dtc_codigo1.Text = parametro
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
'    dtc_aux1.BoundText = dtc_codigo1.BoundText
'    dtc_desc2.Locked = True
'    Select Case parametro
'        Case "DVTA"             'INI COMERCIAL
'            dtc_codigo2.Text = "COM-01"   '3
'        Case "COMEX"            'INI COMEX
'            dtc_codigo2.Text = "CMX-01"   '3
'        Case "DNINS"            'INI GRABA INSTALACIONES
'            dtc_codigo2.Text = "COM-03" '4
'        Case "DNAJS"            'AJUSTE
'            dtc_codigo2.Text = "COM-04" '5
'        Case "DNMAN", "DMANB", "DMANS", "DMANC"            'MANTENIMIENTO PREVENTIVO
'            dtc_codigo2.Text = "TEC-02" '10
'        Case "DNREP", "DREPB", "DREPS", "DREPC"         'MANTENIMIENTO CORRECTIVO / REPARACIONES
'            dtc_codigo2.Text = "TEC-03" '10
'        Case "DNEME"           'EMERGENCIAS
'            dtc_codigo2.Text = "TEC-04" '10
'        Case "DNMOD"            'MODERNIZACION
'            dtc_codigo2.Text = "TEC-05" '10
'        Case Else
'            dtc_codigo2.Text = "TEC-01"   '3
'    End Select
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
''    dtc_codigo5.Text = "COM"
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
''    dtc_codigo6.Text = "COM-01"
''    dtc_desc6.BoundText = dtc_codigo6.BoundText
''    dtc_codigo7.Text = "COM-01-02"
''    dtc_desc7.BoundText = dtc_codigo7.BoundText
''    BtnVer.Visible = False
''    dtc_codigo9.Enabled = False
'  Exit Sub
'AddErr:
'  MsgBox Err.Description
'End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  rs_datos.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Function ExisteReg(Unidad As String, Codigo As Integer) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_ventas_cabecera WHERE unidad_codigo = '" & Unidad & "' and solicitud_codigo=" & Codigo & " and estado_codigo = 'APR'   "
'    <> 'ANL'
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

'Private Function ExisteReg(Unidad As String) As Boolean
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE edif_codigo = '" & Unidad & "'"
'    rs.Open GlSqlAux, db, adOpenStatic
'    ExisteReg = rs!Cuantos > 0
'End Function

Private Sub OptFilGral1_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DPTOC
        Case "2"
            If glusuario = "ADMIN" Or glusuario = "VBELLIDO" Or glusuario = "APACIOS" Or glusuario = "CSALINAS" Then
                queryinicial = "Select * from av_solicitud_venta where (estado_codigo <> 'ANL'  AND ges_gestion = " & Year(Date) & "  and solicitud_tipo = '10' ) "
            Else
                If parametro = "DNINS" Then
                    queryinicial = "Select * from av_solicitud_venta where (estado_codigo <> 'ANL' AND unidad_codigo = 'DVTA' AND ges_gestion = " & Year(Date) & " ) "
                Else
                    queryinicial = "Select * from av_solicitud_venta where (estado_codigo <> 'ANL' AND ges_gestion = " & Year(Date) & " AND (unidad_codigo = '" & parametro & "' ) ) "
                End If
            End If
        Case "7"
            queryinicial = "Select * from av_solicitud_venta where (estado_codigo <> 'ANL' AND unidad_codigo = '" & parametro & "' AND ges_gestion = " & Year(Date) & " AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR   Left(edif_codigo, 1) = '8' OR   Left(edif_codigo, 1) = '9' OR   Left(edif_codigo, 1) = '1' )) "
        Case "3"
            queryinicial = "Select * from av_solicitud_venta where ((estado_codigo <> 'ANL' AND unidad_codigo = '" & parametro & "' AND ges_gestion = " & Year(Date) & ") AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR  Left(edif_codigo, 1) = '4' )) "
        Case "9"
            'queryinicial = "Select * from ao_solicitud where (estado_codigo <> 'ANL' AND unidad_codigo = '" & parametro & "' AND ges_gestion = " & Year(Date) & " AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR  Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' )) "
            queryinicial = "Select * from av_solicitud_venta where (estado_codigo <> 'ANL' AND unidad_codigo = '" & parametro & "' AND ges_gestion = " & Year(Date) & " AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR  Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' )) "
        Case Else
            'queryinicial = "Select * from ao_solicitud where (estado_codigo <> 'ANL' AND ges_gestion = " & Year(Date) & " AND Left(edif_codigo, 1) = '" & VAR_DPTOC & "' AND (unidad_codigo = '" & parametro & "' OR unidad_codigo = '" & VAR_UORIGEN & "')) "
            queryinicial = "Select * from av_solicitud_venta where (estado_codigo <> 'ANL' AND ges_gestion = " & Year(Date) & " AND Left(edif_codigo, 1) = '" & VAR_DPTOC & "' AND (unidad_codigo = '" & parametro & "' OR unidad_codigo = '" & VAR_UORIGEN & "')) "
    End Select
            'queryinicial = "Select * from ao_solicitud where estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "' "
            'queryinicial = "select * From av_ventas_cabecera WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG' AND unidad_codigo='" & VAR_UORIGEN & "' AND left(edif_codigo,1) = '" & VAR_DPTO & "')) "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "unidad_codigo, solicitud_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
    If rs_datos.RecordCount > 0 Then
    rs_datos.MoveFirst
    End If
End Sub

Private Sub OptFilGral2_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DPTOC
        Case "2"
            If parametro = "DNINS" Then
                queryinicial = "Select * from ao_solicitud WHERE (unidad_codigo = 'DVTA') "
            Else
                queryinicial = "Select * from ao_solicitud WHERE (unidad_codigo = '" & parametro & "') "
            End If
        Case "7"
            queryinicial = "Select * from ao_solicitud where (unidad_codigo = '" & parametro & "' AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR   Left(edif_codigo, 1) = '8' OR   Left(edif_codigo, 1) = '9' OR   Left(edif_codigo, 1) = '1' )) "
        Case "3"
            queryinicial = "Select * from ao_solicitud where (unidad_codigo = '" & parametro & "' AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR  Left(edif_codigo, 1) = '4' )) "
        Case "9"
            queryinicial = "Select * from ao_solicitud where (unidad_codigo = '" & parametro & "' AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR  Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' )) "
        Case Else
            'queryinicial = "Select * from ao_solicitud where (unidad_codigo = '" & parametro & "' AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' )) "
            queryinicial = "Select * from ao_solicitud where (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' AND (unidad_codigo = '" & parametro & "' OR unidad_codigo = '" & VAR_UORIGEN & "'))"
    End Select
    'queryinicial = "Select * from ao_solicitud where unidad_codigo = '" & parametro & "' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "unidad_codigo, solicitud_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
     If rs_datos.RecordCount > 0 Then
    rs_datos.MoveFirst
    End If
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_obs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

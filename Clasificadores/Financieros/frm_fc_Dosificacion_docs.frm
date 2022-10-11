VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_fc_dosificacion_docs 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Clasificadores - Dosificacion de Facturas"
   ClientHeight    =   7020
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "frm_fc_Dosificacion_docs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra_ABM 
      BackColor       =   &H00C0C0C0&
      Height          =   5895
      Left            =   6480
      TabIndex        =   23
      Top             =   720
      Width           =   10815
      Begin VB.TextBox Txt_Tipo 
         Alignment       =   2  'Center
         DataField       =   "dgral_codigo"
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2160
         TabIndex        =   55
         Text            =   "1"
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComCtl2.DTPicker TxtFecha1 
         Height          =   285
         Left            =   240
         TabIndex        =   50
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         Format          =   110624769
         CurrentDate     =   44312
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "frm_fc_Dosificacion_docs.frx":0A02
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5040
         TabIndex        =   49
         Top             =   3480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "frm_fc_Dosificacion_docs.frx":0A1B
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3240
         TabIndex        =   47
         Top             =   480
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "doc_descripcion"
         BoundColumn     =   "doc_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "frm_fc_Dosificacion_docs.frx":0A34
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2160
         TabIndex        =   46
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "doc_codigo"
         BoundColumn     =   "doc_codigo"
         Text            =   ""
      End
      Begin VB.TextBox txt_codigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "correl_ini"
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
         Left            =   1920
         TabIndex        =   30
         Text            =   "1"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Txt_descripcion 
         DataField       =   "dosifica_llave"
         DataSource      =   "Ado_datos"
         Height          =   645
         Left            =   1080
         TabIndex        =   29
         Text            =   "-"
         Top             =   975
         Width           =   9375
      End
      Begin VB.TextBox txt_dosifica 
         DataField       =   "dosifica_autorizacion"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   7440
         TabIndex        =   28
         Text            =   "0"
         Top             =   1815
         Width           =   3015
      End
      Begin VB.TextBox Combo3 
         DataField       =   "correl_fin"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   5400
         TabIndex        =   27
         Text            =   "100000"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Combo4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "correl"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   9120
         TabIndex        =   26
         Text            =   "0"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox TxtLeyenda1 
         DataField       =   "dosifica_leyenda1"
         DataSource      =   "Ado_datos"
         Height          =   525
         Left            =   1680
         TabIndex        =   25
         Text            =   "-"
         Top             =   4440
         Width           =   8775
      End
      Begin VB.TextBox TxtLeyenda2 
         DataField       =   "dosifica_leyenda2"
         DataSource      =   "Ado_datos"
         Height          =   525
         Left            =   1680
         TabIndex        =   24
         Text            =   "-"
         Top             =   5160
         Width           =   8775
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "frm_fc_Dosificacion_docs.frx":0A4D
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3480
         TabIndex        =   48
         Top             =   3480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker TxtFecha2 
         Height          =   285
         Left            =   4320
         TabIndex        =   51
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         Format          =   110624769
         CurrentDate     =   44312
      End
      Begin MSComCtl2.DTPicker TxtFecha3 
         Height          =   285
         Left            =   8640
         TabIndex        =   52
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         Format          =   110624769
         CurrentDate     =   44312
      End
      Begin MSComCtl2.DTPicker TxtFecha 
         Height          =   285
         Left            =   1920
         TabIndex        =   53
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         Format          =   110624769
         CurrentDate     =   44312
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "frm_fc_Dosificacion_docs.frx":0A66
         DataField       =   "dgral_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3000
         TabIndex        =   57
         Top             =   3960
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Factura_tipo_descripcion"
         BoundColumn     =   "Factura_tipo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "frm_fc_Dosificacion_docs.frx":0A7F
         DataField       =   "dgral_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2160
         TabIndex        =   58
         Top             =   3960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "Factura_tipo"
         BoundColumn     =   "Factura_tipo"
         Text            =   ""
      End
      Begin VB.Label LblTipo 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Facturación"
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
         Left            =   3000
         TabIndex        =   56
         Top             =   3960
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Facturación"
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
         TabIndex        =   54
         Top             =   3960
         Width           =   1800
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
         Left            =   8880
         TabIndex        =   45
         Top             =   495
         Width           =   645
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
         Height          =   315
         Left            =   9600
         TabIndex        =   44
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lbl_codigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "No.de Autorización"
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
         Left            =   5640
         TabIndex        =   43
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "LLAVE"
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
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lbl_enlace1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Documento"
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
         Top             =   495
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Encargado Emisión Factura/Recibo"
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
         TabIndex        =   40
         Top             =   3465
         Width           =   3210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Solicitud"
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
         TabIndex        =   39
         Top             =   1800
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Correlativo Inicial"
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
         TabIndex        =   38
         Top             =   2295
         Width           =   1530
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Correlativo Final"
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
         Left            =   3840
         TabIndex        =   37
         Top             =   2295
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Correlativo Actual"
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
         Left            =   7440
         TabIndex        =   36
         Top             =   2295
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio Emisión"
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
         TabIndex        =   35
         Top             =   2745
         Width           =   1845
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin Emisión"
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
         Left            =   4440
         TabIndex        =   34
         Top             =   2745
         Width           =   1650
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Limite Emisión"
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
         Left            =   8535
         TabIndex        =   33
         Top             =   2745
         Width           =   1905
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "LEYENDA 1 (Pie de Página)"
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
         Height          =   600
         Left            =   240
         TabIndex        =   32
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "LEYENDA 2 (Pie de Página)"
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
         TabIndex        =   31
         Top             =   5160
         Width           =   1455
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   12
      Top             =   0
      Width           =   20280
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Picture         =   "frm_fc_Dosificacion_docs.frx":0A98
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10800
         Picture         =   "frm_fc_Dosificacion_docs.frx":0CA2
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "frm_fc_Dosificacion_docs.frx":10E4
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   19
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
         Picture         =   "frm_fc_Dosificacion_docs.frx":18A3
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   18
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
         Picture         =   "frm_fc_Dosificacion_docs.frx":21B8
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   17
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
         Picture         =   "frm_fc_Dosificacion_docs.frx":2904
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   16
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
         Picture         =   "frm_fc_Dosificacion_docs.frx":3137
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   15
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
         Picture         =   "frm_fc_Dosificacion_docs.frx":38EC
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   14
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17400
         Picture         =   "frm_fc_Dosificacion_docs.frx":41B9
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   13
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
         Left            =   12855
         TabIndex        =   22
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
      TabIndex        =   8
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
         Picture         =   "frm_fc_Dosificacion_docs.frx":497B
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   10
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
         Picture         =   "frm_fc_Dosificacion_docs.frx":5151
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   9
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
         TabIndex        =   11
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GERENCIA GENERAL"
      ForeColor       =   &H00FF0000&
      Height          =   5895
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   6255
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   5520
         Width           =   6045
         _ExtentX        =   10663
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
         Caption         =   " <-- Inicio                     Dosificación Facturas                       Fin -->"
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
         Bindings        =   "frm_fc_Dosificacion_docs.frx":5A3D
         Height          =   5235
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   9234
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "dosifica_autorizacion"
            Caption         =   "No.autirizacion"
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
         BeginProperty Column01 
            DataField       =   "dosifica_fecha"
            Caption         =   "Fecha_solicitud"
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
            DataField       =   "dosifica_fecha_limite"
            Caption         =   "Fecha_Limite"
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
            DataField       =   "doc_codigo"
            Caption         =   "Tipo.Doc"
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
            Caption         =   "Estado."
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
            DataField       =   "correl_ini"
            Caption         =   "Correl_ini"
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
            DataField       =   "correl_fin"
            Caption         =   "Correl_Fin"
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
            DataField       =   "correl"
            Caption         =   "Correl_actual"
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
            DataField       =   "dosifica_fecha_ini"
            Caption         =   "Fecha_ini"
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
            DataField       =   "dosifica_fecha_fin"
            Caption         =   "Fecha_Fin"
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
         BeginProperty Column11 
            DataField       =   "fecha_registro"
            Caption         =   "fecha_registro"
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
            DataField       =   "dgral_codigo"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1695.118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   615.118
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
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   900.284
            EndProperty
         EndProperty
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
      TabIndex        =   0
      Top             =   7020
      Width           =   11280
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   5040
      Top             =   6720
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
      Top             =   6600
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
      Top             =   6600
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   5520
      Top             =   6720
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
End
Attribute VB_Name = "frm_fc_dosificacion_docs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If rs_datos!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "APR"
         rs_datos!Fecha_Registro = Date
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
        rs_datos.CancelUpdate
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
        txt_dosifica.Enabled = True
        dtc_desc1.Enabled = True
    End If
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
   If ExisteReg(Ado_datos.Recordset!da_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
   If rs_datos!estado_codigo = "APR" Then
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "ANL"
         rs_datos!Fecha_Registro = Date
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

Private Sub BtnDesAprobar_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_datos!estado_codigo = "APR" Then
      If sino = vbYes Then
         rs_datos!estado_codigo = "REG"
         rs_datos!Fecha_Registro = Date
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
        Set rs_aux1 = New ADODB.Recordset
        SQL_FOR = "select * from Fc_dosificacion_docs where dosifica_autorizacion = '" & txt_dosifica.Text & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            MsgBox " CODIGO DUPLICADO, Vuelva a intentar..."
            txt_dosifica.SetFocus
            Exit Sub
            'var_cod = rs_aux1.RecordCount + 1
        'Else
        '    var_cod = 1
        End If
     End If
     
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_parametros_sistema WHERE (estado_codigo = 'APR')", db, adOpenStatic
'    Set Ado_datos3.Recordset = rs_datos3
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
     
     rs_datos!doc_codigo.Value = dtc_codigo1.Text         '"R-101"
     Select Case dtc_codigo1.Text
        Case "R-101"
            rs_datos!dosifica_autorizacion.Value = txt_dosifica.Text
        Case "R-102"
            rs_datos!dosifica_autorizacion.Value = txt_dosifica.Text
        Case "R-103", "R-642", "R-209"
            rs_datos!dosifica_autorizacion.Value = dtc_codigo2.Text + "-" + Format(txtFecha.Value, "YYYYMMDD")
        Case Else
            rs_datos!dosifica_autorizacion.Value = txt_dosifica.Text
     End Select
'        If dtc_codigo1.Text = "R-101" Or dtc_codigo1.Text = "R-102" Then
'            rs_datos!dosifica_autorizacion.Value = txt_dosifica.Text
'        Else
'            rs_datos!dosifica_autorizacion.Value = dtc_codigo2.Text + Format(TxtFecha.Value, "YYYYMMDD")
'        End If
        rs_datos!dosifica_fecha.Value = txtFecha.Value
        'Veficiar NIT CGI WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
        rs_datos!beneficiario_codigo.Value = rs_datos3!param_codigo
        rs_datos!beneficiario_codigo_resp.Value = dtc_codigo2.Text
        'rs_datos!beneficiario_codigo.Value = dtc_codigo2.Text
        rs_datos!correl_ini.Value = txt_codigo.Text
        rs_datos!correl_fin.Value = Combo3.Text
        rs_datos!CORREL.Value = IIf(Combo4.Text = "", "0", Combo4.Text)
        rs_datos!dgral_codigo = dtc_codigo4.Text                         'Txt_Tipo.Text
        rs_datos!dosifica_fecha_ini.Value = TxtFecha1.Value
        rs_datos!dosifica_fecha_fin.Value = txtFecha2.Value
        rs_datos!dosifica_fecha_limite.Value = txtFecha3.Value
        rs_datos!dosifica_codigo_control.Value = rs_datos.RecordCount + 1
        rs_datos!dosifica_llave.Value = Txt_descripcion.Text
        rs_datos!dosifica_leyenda1.Value = TxtLeyenda1.Text
        rs_datos!dosifica_leyenda2.Value = TxtLeyenda2.Text
        
        rs_datos!estado_codigo.Value = "REG"
        rs_datos!usr_codigo.Value = glusuario
        rs_datos!Fecha_Registro.Value = Date
        
        rs_datos.Update
     Call ABRIR_TABLA
     rs_datos.MoveLast
     mbDataChanged = False
      
      Fra_ABM.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      dg_datos.Enabled = True
      txt_dosifica.Enabled = True
      dtc_desc1.Enabled = True
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
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar: " + lbl_enlace1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  Dim iResult As Integer
  cr01.WindowShowPrintSetupBtn = True
  cr01.WindowShowRefreshBtn = True
  cr01.ReportFileName = App.Path & "\REPORTES\clasificadores\fr_dosificacion_facturas.rpt"
  iResult = cr01.PrintReport
  If iResult <> 0 Then
      MsgBox cr01.LastErrorNumber & " : " & cr01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
  cr01.WindowState = crptMaximized
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset.RecordCount > 0 Then
        If Ado_datos.Recordset!estado_codigo = "REG" Then
            Fra_ABM.Enabled = True
            fraOpciones.Visible = False
            FraGrabarCancelar.Visible = True
            dg_datos.Enabled = False
            VAR_SW = "MOD"
            txt_dosifica.Enabled = False
            dtc_desc1.Enabled = False
        Else
            MsgBox "No se puede modificar un registro aprobado, revise y vuelva a intentar ..."
        End If
        MsgBox "No existen registros para modificar, , revise y vuelva a intentar ..."
    End If
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

Private Sub dtc_codigo1_LostFocus()
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    Select Case dtc_codigo1.Text
        Case "R-101", "R-102"           'FACTURA
            Txt_descripcion.Visible = True
            txt_dosifica.Visible = True
            lbl_descripcion.Visible = True
            lbl_codigo.Visible = True
            rs_datos2.Open "Select * from rv_unidad_vs_responsable where unidad_codigo= 'DCONT' ", db, adOpenStatic
        Case "R-103"                    'RECIBO COBRANZAS
            Txt_descripcion.Text = "NINGUNO"
            Txt_descripcion.Visible = False
            txt_dosifica.Visible = False
            lbl_descripcion.Visible = False
            lbl_codigo.Visible = False
            rs_datos2.Open "Select * from rv_unidad_vs_responsable where unidad_codigo= 'DCOBR' ", db, adOpenStatic
        Case "R-209"                    'RECIBO COMERCIAL
            Txt_descripcion.Text = "NINGUNO"
            Txt_descripcion.Visible = False
            txt_dosifica.Visible = False
            lbl_descripcion.Visible = False
            lbl_codigo.Visible = False
            rs_datos2.Open "Select * from rv_unidad_vs_responsable where unidad_codigo= 'DCOBR' AND beneficiario_codigo = '4792382' ", db, adOpenStatic
    End Select
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from gc_tipo_facturacion where doc_codigo = '" & dtc_codigo1.Text & "' ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText

End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc1_LostFocus()
    
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    Select Case dtc_codigo1.Text
        Case "R-101", "R-102"           'FACTURA
            Txt_descripcion.Visible = True
            txt_dosifica.Visible = True
            lbl_descripcion.Visible = True
            lbl_codigo.Visible = True
            rs_datos2.Open "Select * from rv_unidad_vs_responsable where unidad_codigo= 'DCONT' ", db, adOpenStatic
        Case "R-103"                    'RECIBO COBRANZAS
            Txt_descripcion.Text = "NINGUNO"
            Txt_descripcion.Visible = False
            txt_dosifica.Visible = False
            lbl_descripcion.Visible = False
            lbl_codigo.Visible = False
            rs_datos2.Open "Select * from rv_unidad_vs_responsable where unidad_codigo= 'DCOBR' ", db, adOpenStatic
        Case "R-209"                    'RECIBO COMERCIAL
            Txt_descripcion.Text = "NINGUNO"
            Txt_descripcion.Visible = False
            txt_dosifica.Visible = False
            lbl_descripcion.Visible = False
            lbl_codigo.Visible = False
            rs_datos2.Open "Select * from rv_unidad_vs_responsable where unidad_codigo= 'DCOBR' AND beneficiario_codigo = '4792382' ", db, adOpenStatic
    End Select
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from gc_tipo_facturacion where doc_codigo = '" & dtc_codigo1.Text & "' ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLAS_AUX
    Call ABRIR_TABLA
    txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_ABM.Enabled = False
    dg_datos.Enabled = True
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
End Sub

Private Sub ABRIR_TABLA()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from fc_dosificacion_docs  "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from gc_documentos_respaldo WHERE (clasif_codigo = 'ADM' or clasif_codigo = 'CMX') AND (doc_original = 'FACTURA') order by doc_descripcion ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from rv_unidad_vs_responsable where unidad_codigo= 'DCONT' ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from gc_tipo_facturacion where estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
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
      If Ado_datos.Recordset.RecordCount > 0 Then
        Select Case Ado_datos.Recordset!dgral_codigo
            Case "0"
                lblTipo.Caption = "FACTURACION ON-LINE"
            Case "1"
                lblTipo.Caption = "FACTURACION CGI"
            Case "2"
                lblTipo.Caption = "FACTURACION CGE"
            Case "9"
                lblTipo.Caption = "RECIBOS VENTAS NUEVAS"
            Case Else
                lblTipo.Caption = "RECIBOS COBRANZAS"
        End Select
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
    'txt_codigo.Enabled = False
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
    'GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE dgral_codigo = '" & Unidad & "'"
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM pc_poa_estrategico WHERE estado_codigo = 'APR' and dgral_codigo = '" & Unidad & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub txt_codigo_LostFocus()
    Combo4.Text = Val(txt_codigo.Text) - 1
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub




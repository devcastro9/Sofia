VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_to_solicitud_bitacora 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bitacora de Eventos"
   ClientHeight    =   5985
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_to_solicitud_bitacora.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   10635
      TabIndex        =   38
      Top             =   120
      Width           =   10695
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   240
         Picture         =   "frm_to_solicitud_bitacora.frx":6C032
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "frm_to_solicitud_bitacora.frx":6C23C
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BITACORA DE EVENTOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   330
         Left            =   4965
         TabIndex        =   41
         Top             =   240
         Width           =   3465
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00000000&
      Height          =   4215
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   10695
      Begin VB.ComboBox MM 
         Height          =   315
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox HH 
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Txt_campo5 
         DataField       =   "bitacora_cite"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
         Height          =   285
         Left            =   8640
         TabIndex        =   33
         Text            =   "0"
         Top             =   3600
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "frm_to_solicitud_bitacora.frx":6C446
         DataField       =   "negocia_forma"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   4200
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "negocia_forma"
         BoundColumn     =   "negocia_forma"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "frm_to_solicitud_bitacora.frx":6C45F
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   9000
         TabIndex        =   22
         Top             =   1920
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "frm_to_solicitud_bitacora.frx":6C478
         DataField       =   "beneficiario_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   3600
         TabIndex        =   20
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin VB.TextBox Txt_campo4 
         DataField       =   "negocia_observaciones"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   3600
         Width           =   8085
      End
      Begin VB.TextBox Txt_campo3 
         DataField       =   "negocia_tarea_realizada"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   2880
         Width           =   9980
      End
      Begin VB.TextBox Txt_monto1 
         DataField       =   "negocia_gasto_estimado"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
         Height          =   285
         Left            =   8880
         TabIndex        =   24
         Text            =   "0"
         Top             =   1440
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "frm_to_solicitud_bitacora.frx":6C491
         DataField       =   "negocia_forma"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   1440
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "negocia_forma_descripcion"
         BoundColumn     =   "negocia_forma"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "frm_to_solicitud_bitacora.frx":6C4AA
         DataField       =   "beneficiario_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   2160
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
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
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "frm_to_solicitud_bitacora.frx":6C4C3
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   5280
         TabIndex        =   21
         Top             =   2160
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "negocia_fecha_real"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
         Height          =   300
         Left            =   5280
         TabIndex        =   23
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   42074113
         CurrentDate     =   41678
      End
      Begin MSComCtl2.DTPicker Txt_campo2 
         DataField       =   "negocia_hora_real"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
         Height          =   300
         Left            =   7080
         TabIndex        =   42
         Top             =   1440
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Format          =   42074114
         CurrentDate     =   0.375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro.Solicitud"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Left            =   360
         TabIndex        =   35
         Top             =   450
         Width           =   1140
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cite / Referencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   8760
         TabIndex        =   34
         Top             =   3345
         Width           =   1485
      End
      Begin VB.Label Txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "unidad_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   5760
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Conclusiones u Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   360
         TabIndex        =   12
         Top             =   3340
         Width           =   2790
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tema Tratado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   360
         TabIndex        =   32
         Top             =   2620
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Responsable (Personal de CGI)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   5280
         TabIndex        =   31
         Top             =   1905
         Width           =   2865
      End
      Begin VB.Label lbl_persona1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Persona Contactada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   360
         TabIndex        =   30
         Top             =   1905
         Width           =   1845
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Gasto en Bs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   8880
         TabIndex        =   29
         Top             =   1200
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Horas -  Minutos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7080
         TabIndex        =   28
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha Contacto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   5280
         TabIndex        =   27
         Top             =   1200
         Width           =   1410
      End
      Begin VB.Label Txt_descripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2040
         TabIndex        =   18
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Correl.Bitácora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Left            =   7395
         TabIndex        =   16
         Top             =   450
         Width           =   1335
      End
      Begin VB.Label Txt_Correl 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "bitacora_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
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
         Left            =   7440
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
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
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   8
         Left            =   2040
         TabIndex        =   10
         Top             =   450
         Width           =   2160
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Forma de Negociación / Tipo de Contacto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   360
         TabIndex        =   9
         Top             =   1190
         Width           =   3765
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "REG"
         DataField       =   "estado_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle1"
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
         Left            =   9000
         TabIndex        =   0
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   2
         Left            =   8880
         TabIndex        =   8
         Top             =   450
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
      ScaleWidth      =   10935
      TabIndex        =   1
      Top             =   5985
      Width           =   10935
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   6
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   2040
      Top             =   6600
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
      Top             =   5520
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
      Top             =   5520
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
      Left            =   4680
      Top             =   5520
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
End
Attribute VB_Name = "frm_to_solicitud_bitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos1 As New ADODB.Recordset
Attribute rs_datos1.VB_VarHelpID = -1
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
'BUSCADOR
Dim var_cod As String
Dim VAR_VAL As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        'tw_identificacion_cliente.Ado_detalle1.Recordset.CancelUpdate
        tw_identificacion_cliente.Ado_detalle1.Recordset.Cancel
        Unload Me
    End If
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    Select Case txt_campo1.Caption
        Case "DRRHH"    'SOLO COMPRAS BB y SS
'            If swnuevo = 1 Then
'                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("ges_gestion").Value = glGestion  'Year(Date)
'                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("unidad_codigo").Value = Txt_campo1.Caption
'                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("solicitud_codigo").Value = txt_codigo.Caption
'                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("estado_codigo").Value = "REG"
'                Set rs_aux1 = New ADODB.Recordset
'                If rs_aux1.State = 1 Then rs_aux1.Close
'                rs_aux1.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & Txt_campo1.Caption & "' ", db, adOpenKeyset, adLockOptimistic
'                If rs_aux1.RecordCount > 0 Then
'                    var_cod = rs_aux1!correl_bitacora + 1
'                Else
'                    var_cod = 1
'                End If
'                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("bitacora_codigo").Value = var_cod
'                'Actualiza correaltivo ...
'                db.Execute "Update gc_unidad_ejecutora Set correl_bitacora = " & var_cod & " Where unidad_codigo = '" & Txt_campo1.Caption & "'   "
'             End If
'             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_forma").Value = dtc_codigo1.Text
'             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_fecha_real").Value = DTPfecha1.Value
'             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_hora_real").Value = Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
'             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_gasto_estimado").Value = Txt_monto1.Text
'             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("beneficiario_codigo").Value = dtc_codigo2.Text
'             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("beneficiario_codigo_resp").Value = dtc_codigo3.Text
'             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_tarea_realizada").Value = Txt_campo3.Text
'             If swnuevo = 1 Then
'                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_observaciones").Value = Trim(dtc_desc1.Text) + " - " + Txt_campo4.Text
'             Else
'                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_observaciones").Value = Txt_campo4.Text
'             End If
'             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("bitacora_cite").Value = Txt_campo5.Text
'
'             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("fecha_registro").Value = Date
'             'frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("hora_registro").Value = Date
'             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("usr_codigo").Value = glusuario
'             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset.UpdateBatch adAffectAll
        
        Case "2"    'SOLO VENTA DE BIENES
        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
            

        Case "DNMAN", "DNREP", "DNINS", "DNAJS", "DNEME", "DNMOD"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
             If swnuevo = 1 Then
                tw_identificacion_cliente.Ado_detalle1.Recordset("ges_gestion").Value = glGestion  'Year(Date)
                tw_identificacion_cliente.Ado_detalle1.Recordset("unidad_codigo").Value = txt_campo1.Caption
                tw_identificacion_cliente.Ado_detalle1.Recordset("solicitud_codigo").Value = txt_codigo.Caption
                tw_identificacion_cliente.Ado_detalle1.Recordset("estado_codigo").Value = "REG"
                Set rs_aux1 = New ADODB.Recordset
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & txt_campo1.Caption & "' ", db, adOpenKeyset, adLockOptimistic
                If rs_aux1.RecordCount > 0 Then
                    var_cod = rs_aux1!correl_bitacora + 1
                Else
                    var_cod = 1
                End If
                tw_identificacion_cliente.Ado_detalle1.Recordset("bitacora_codigo").Value = var_cod
                'Actualiza correaltivo ...
                db.Execute "Update gc_unidad_ejecutora Set correl_bitacora = " & var_cod & " Where unidad_codigo = '" & txt_campo1.Caption & "'   "
             End If
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_forma").Value = dtc_codigo1.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_fecha_real").Value = DTPfecha1.Value
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_hora_real").Value = Format(Txt_campo2.Value, "HH:MM")    '   Trim(HH.Text) + ":" + Trim(MM.Text)
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_gasto_estimado").Value = txt_monto1.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("beneficiario_codigo").Value = dtc_codigo2.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("beneficiario_codigo_resp").Value = dtc_codigo3.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_tarea_realizada").Value = Txt_campo3.Text
             If swnuevo = 1 Then
                tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_observaciones").Value = Trim(dtc_desc1.Text) + " - " + Txt_campo4.Text
             Else
                tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_observaciones").Value = Txt_campo4.Text
             End If
             tw_identificacion_cliente.Ado_detalle1.Recordset("bitacora_cite").Value = Txt_campo5.Text
        
             tw_identificacion_cliente.Ado_detalle1.Recordset("fecha_registro").Value = Date
             'tw_identificacion_cliente.Ado_detalle1.Recordset("hora_registro").Value = Date
             tw_identificacion_cliente.Ado_detalle1.Recordset("usr_codigo").Value = glusuario
             tw_identificacion_cliente.Ado_detalle1.Recordset.UpdateBatch adAffectAll
     


        Case "5"    ' SERVICIO MODERNIZACION
    End Select
    
     'db.Execute "Update ao_solicitud Set correl_bitacora = " & tw_identificacion_cliente.Ado_detalle1.Recordset("bitacora_codigo") & " Where unidad_codigo = '" & var_cod & "' and solicitud_codigo = '" & txt_codigo.Caption & "'   "
     Unload Me

'     Call ABRIR_TABLA
'     rs_datos.MoveLast
'     mbDataChanged = False
'
'      Fra_ABM.Enabled = False
'      fraOpciones.Visible = True
'      FraGrabarCancelar.Visible = False
'      dg_datos.Enabled = True
'      txt_codigo.Enabled = True
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar la " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo2.Text = "" Then
    MsgBox "Debe registrar la " + lbl_persona1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo3.Text = "" Then
    MsgBox "Debe registrar la " + lbl_persona1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
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

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub Form_Activate()
    'var_cod = AUX
    'var_cod = "DRRHH"
    var_cod = frm_ao_solicitud_bitacora.txt_campo1.Caption
    'Call ABRIR_TABLA
End Sub

Private Sub Form_Load()
    var_cod = Aux
    'var_cod = "DRRHH"
    'var_cod = frm_ao_solicitud_bitacora.Txt_campo1.Caption
    Call ABRIR_TABLA
    mbDataChanged = False
'    If swnuevo = 2 Then
'        dtc_desc2.BoundText = dtc_codigo2.BoundText
'        dtc_desc3.BoundText = dtc_codigo3.BoundText
'    End If
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLA()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from ac_negociacion_forma ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    Select Case var_cod
'        Case "DVTA"        'INI COMERCIAL
'            rs_datos2.Open "Select * from gc_beneficiario where tipoben_codigo = '1' order by solicitud_tipo", db, adOpenStatic
'        Case "COMEX"        'INI COMEX
'            dtc_codigo2.Text = 3
'        Case "DNINS"                        'INI GRABA INSTALACIONES
'            dtc_codigo2.Text = 4
'        Case "DNAJS"
'            dtc_codigo2.Text = 4
'        Case "DNMAN"
'            dtc_codigo2.Text = 4
        Case "DRRHH"            'RECURSOS HUMANOS - CGI
            rs_datos2.Open "Select * from gc_beneficiario where tipoben_codigo = '1' order by beneficiario_denominacion ", db, adOpenStatic
            'rs_datos2.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
        Case Else
            rs_datos2.Open "Select * from gc_beneficiario where tipoben_codigo <> '1' order by beneficiario_denominacion", db, adOpenStatic
    End Select
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & var_cod & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    'rs_datos3.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
    sino = rs_datos3.RecordCount
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText

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

'Private Sub MM_LostFocus()
'    Txt_campo2.Text = Trim(HH) + ":" + Trim(MM)
'End Sub

Private Sub txt_campo3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_campo4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tw_solicitud_bitacora 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Bitacora de Eventos"
   ClientHeight    =   6150
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   650
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   10665
      TabIndex        =   22
      Top             =   120
      Width           =   10695
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   -30
         Picture         =   "tw_solicitud_bitacora.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   -30
         Width           =   1365
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   1300
         MaskColor       =   &H00000000&
         Picture         =   "tw_solicitud_bitacora.frx":07D6
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cancelar"
         Top             =   -30
         Width           =   1365
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BITACORA DE EVENTOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   5430
         TabIndex        =   25
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      Height          =   5175
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   10695
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "tw_solicitud_bitacora.frx":10C2
         DataField       =   "motivo_codigo"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
         Height          =   315
         Left            =   9120
         TabIndex        =   47
         Top             =   3600
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "motivo_codigo"
         BoundColumn     =   "motivo_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "tw_solicitud_bitacora.frx":10DB
         DataField       =   "motivo_codigo"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
         Height          =   315
         Left            =   360
         TabIndex        =   46
         Top             =   3600
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "motivo_denominacion"
         BoundColumn     =   "motivo_codigo"
         Text            =   "Todos"
      End
      Begin VB.CommandButton BtnGrabar2 
         BackColor       =   &H80000015&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   9000
         Picture         =   "tw_solicitud_bitacora.frx":10F4
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4275
         Visible         =   0   'False
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker Txt_campo2 
         DataField       =   "negocia_hora_real"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
         Height          =   300
         Left            =   7200
         TabIndex        =   2
         Top             =   1440
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Format          =   117964802
         CurrentDate     =   0.375
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente Contactado (Registre una de las 2 opciones...)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   240
         TabIndex        =   41
         Top             =   1800
         Width           =   10215
         Begin VB.TextBox txt_cliente 
            DataField       =   "beneficiario_nombre_ref"
            DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
            Height          =   315
            Left            =   5160
            TabIndex        =   5
            Top             =   560
            Width           =   4935
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "tw_solicitud_bitacora.frx":1B85
            DataField       =   "beneficiario_codigo"
            DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
            Height          =   315
            Left            =   3840
            TabIndex        =   43
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "tw_solicitud_bitacora.frx":1B9E
            DataField       =   "beneficiario_codigo"
            DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   560
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin VB.Label lbl_persona2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "2. Datos Referenciales Cliente (Apellidos, Nombres ...)"
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
            Left            =   5160
            TabIndex        =   44
            Top             =   300
            Width           =   4830
         End
         Begin VB.Label lbl_persona1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "1. Existente en la Base de Datos"
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
            TabIndex        =   42
            Top             =   300
            Width           =   2880
         End
      End
      Begin VB.TextBox Txt_campo5 
         Alignment       =   2  'Center
         DataField       =   "bitacora_cite"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
         Height          =   285
         Left            =   7200
         TabIndex        =   9
         Text            =   "0"
         Top             =   4560
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "tw_solicitud_bitacora.frx":1BB7
         DataField       =   "negocia_forma"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
         Height          =   315
         Left            =   3960
         TabIndex        =   32
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "negocia_forma"
         BoundColumn     =   "negocia_forma"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "tw_solicitud_bitacora.frx":1BD0
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
         Height          =   315
         Left            =   4440
         TabIndex        =   34
         Top             =   3000
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
      Begin VB.TextBox Txt_campo4 
         DataField       =   "negocia_observaciones"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
         Height          =   555
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   4320
         Width           =   6405
      End
      Begin VB.TextBox Txt_campo3 
         DataField       =   "negocia_tarea_realizada"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
         Height          =   315
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   3600
         Width           =   9980
      End
      Begin VB.TextBox Txt_monto1 
         DataField       =   "negocia_gasto_estimado"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
         Height          =   285
         Left            =   8880
         TabIndex        =   3
         Text            =   "0"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Txt_campo2A 
         DataSource      =   "aw_p_ao_solicitud.ado_detalle2"
         Height          =   285
         Left            =   7200
         TabIndex        =   10
         Text            =   "0"
         Top             =   1440
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "tw_solicitud_bitacora.frx":1BE9
         DataField       =   "negocia_forma"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
         Height          =   315
         Left            =   360
         TabIndex        =   0
         Top             =   1440
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "negocia_forma_descripcion"
         BoundColumn     =   "negocia_forma"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "negocia_fecha_real"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
         Height          =   300
         Left            =   5280
         TabIndex        =   1
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   117964801
         CurrentDate     =   41678
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "tw_solicitud_bitacora.frx":1C02
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
         Height          =   315
         Left            =   4410
         TabIndex        =   6
         Top             =   3000
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7320
         TabIndex        =   40
         Top             =   4275
         Width           =   1485
      End
      Begin VB.Label Txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         DataField       =   "unidad_codigo"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5280
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   28
         Top             =   4065
         Width           =   2790
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   39
         Top             =   3345
         Width           =   1305
      End
      Begin VB.Label lbl_persona3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Personal CGI que recepciona la informacion"
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
         Left            =   360
         TabIndex        =   38
         Top             =   3000
         Width           =   3930
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8880
         TabIndex        =   37
         Top             =   1200
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora del Contacto"
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
         Left            =   7080
         TabIndex        =   36
         Top             =   1200
         Width           =   1620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Contacto"
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
         Left            =   5160
         TabIndex        =   35
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Txt_descripcion 
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
         Left            =   2040
         TabIndex        =   33
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Correl. Bitácora"
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
         Left            =   7200
         TabIndex        =   31
         Top             =   450
         Width           =   1380
      End
      Begin VB.Label Txt_Correl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "bitacora_codigo"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
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
         Left            =   7320
         TabIndex        =   30
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
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
         Left            =   360
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblLabels 
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
         Index           =   8
         Left            =   2040
         TabIndex        =   26
         Top             =   450
         Width           =   2160
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   " Tipo de Contacto / Forma de Negociación"
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
         Left            =   360
         TabIndex        =   21
         Top             =   1185
         Width           =   3810
      End
      Begin VB.Label lbl_codigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. Tramite"
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
         Left            =   360
         TabIndex        =   20
         Top             =   450
         Width           =   1155
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REG"
         DataField       =   "estado_codigo"
         DataSource      =   "tw_tecnico_bitacora.ado_detalle1"
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
         Left            =   9000
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         Left            =   8880
         TabIndex        =   19
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
      ScaleWidth      =   10950
      TabIndex        =   12
      Top             =   6150
      Width           =   10950
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   17
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   3120
      Top             =   6000
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
      Left            =   2160
      Top             =   6120
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
      Left            =   6480
      Top             =   6120
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
      Left            =   4320
      Top             =   6120
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
      Left            =   0
      Top             =   6120
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
Attribute VB_Name = "tw_solicitud_bitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos1 As New ADODB.Recordset
Attribute rs_datos1.VB_VarHelpID = -1
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset

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
        tw_tecnico_bitacora.Ado_detalle1.Recordset.CancelUpdate
        Unload Me
    End If
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If swnuevo = 1 Then
        tw_tecnico_bitacora.Ado_detalle1.Recordset("ges_gestion").Value = Year(DTPfecha1.Value)
        tw_tecnico_bitacora.Ado_detalle1.Recordset("unidad_codigo").Value = Txt_campo1.Caption
        tw_tecnico_bitacora.Ado_detalle1.Recordset("solicitud_codigo").Value = txt_codigo.Caption
        tw_tecnico_bitacora.Ado_detalle1.Recordset("estado_codigo").Value = "REG"
     End If
     tw_tecnico_bitacora.Ado_detalle1.Recordset("negocia_forma").Value = dtc_codigo1.Text
     tw_tecnico_bitacora.Ado_detalle1.Recordset("negocia_fecha_real").Value = DTPfecha1.Value 'Format(Time, "hh:mm:ss")
     tw_tecnico_bitacora.Ado_detalle1.Recordset("negocia_hora_real").Value = IIf(Txt_campo2A.Text = "" Or Txt_campo2A.Text = "0", "09:00:00", Txt_campo2A.Text)        'Str(FormatDateTime(txt_campo2.Value, vbLongTime))
     tw_tecnico_bitacora.Ado_detalle1.Recordset("negocia_gasto_estimado").Value = Txt_monto1.Text
     If dtc_codigo2.Text = "" Or dtc_codigo2.Text = "0" Then
        tw_tecnico_bitacora.Ado_detalle1.Recordset("beneficiario_codigo").Value = "0"      'IIf(dtc_codigo2.Text = "", "0", dtc_codigo2.Text)
     Else
        tw_tecnico_bitacora.Ado_detalle1.Recordset("beneficiario_codigo").Value = IIf(dtc_codigo2.Text = "", "0", dtc_codigo2.Text)
     End If
     tw_tecnico_bitacora.Ado_detalle1.Recordset!beneficiario_nombre_ref = IIf(txt_cliente = "", dtc_desc2, txt_cliente)
     tw_tecnico_bitacora.Ado_detalle1.Recordset("beneficiario_codigo_resp").Value = dtc_codigo3.Text
     tw_tecnico_bitacora.Ado_detalle1.Recordset("negocia_tarea_realizada").Value = Txt_campo3.Text
     If Txt_campo4.Text = "" Then
        tw_tecnico_bitacora.Ado_detalle1.Recordset("negocia_observaciones").Value = Trim(dtc_desc1.Text) '+ " - " + txt_campo4.Text
     Else
        tw_tecnico_bitacora.Ado_detalle1.Recordset("negocia_observaciones").Value = Trim(Txt_campo4.Text)
     End If
     tw_tecnico_bitacora.Ado_detalle1.Recordset("bitacora_cite").Value = Txt_campo5.Text
     tw_tecnico_bitacora.Ado_detalle1.Recordset!motivo_codigo.Value = dtc_codigo4.Text
     tw_tecnico_bitacora.Ado_detalle1.Recordset("fecha_registro").Value = Date
     'tw_tecnico_bitacora.Ado_detalle1.Recordset("hora_registro").Value = Date
     tw_tecnico_bitacora.Ado_detalle1.Recordset("usr_codigo").Value = glusuario
     tw_tecnico_bitacora.Ado_detalle1.Recordset.UpdateBatch adAffectAll
     db.Execute "Update ao_solicitud Set correl_bitacora = " & tw_tecnico_bitacora.Ado_detalle1.Recordset("bitacora_codigo") & " Where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = '" & txt_codigo.Caption & "'   "
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
  If dtc_codigo2.Text = "" And txt_cliente.Text = "" Then
    MsgBox "Debe registrar Cliente " + lbl_persona1.Caption + " o " + lbl_persona2.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo3.Text = "" Then
    MsgBox "Debe registrar aL " + lbl_persona3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnGrabar2_Click()
    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "Select * from gc_unidad_ejecutora correl_bitacora where unidad_codigo = '" & Txt_campo1.Caption & "' ", db, adOpenStatic
    If rs_aux1.RecordCount > 0 Then
        'If rs_aux1!UNIDAD_CODIGO = "" Then
        Select Case Txt_campo1.Caption
                Case "DNMAN"
                    'MANTENIMIENTO
                    Txt_campo5 = "TECM-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                Case "DMANS"
                    'MANTENIMIENTO
                    Txt_campo5 = "TECMS-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                Case "DMANB"
                    'MANTENIMIENTO
                    Txt_campo5 = "TECMB-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                Case "DMANC"
                    'MANTENIMIENTO
                    Txt_campo5 = "TECMC-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                Case "DNREP"
                    'REPARACIONES LA PAZ
                    Txt_campo5 = "TECR-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                Case "DREPS"
                    'REPARACIONES STA.CRUZ
                    Txt_campo5 = "TECS-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                Case "DREPB"
                    'REPARACIONES CBBA
                    Txt_campo5 = "TECB-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                Case "DREPC"
                    'REPARACIONES CHUQUISACA
                    Txt_campo5 = "TECC-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                Case "DNEME"
                    'EMERGENCIAS
                    Txt_campo5 = "TECE-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                 Case "DNMOD"
                    'MODERNIZACION
                    Txt_campo5 = "TECD-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
                Case Else
                    Txt_campo5 = "TECM-" + Str(rs_aux1!correl_bitacora + 1) + "/" + Str(Year(Date))
        End Select
        'db.Execute "update ao_solicitud set correl_bitacora =  correl_bitacora + 1 where unidad_codigo = '" & txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & " "
        db.Execute "update gc_unidad_ejecutora set correl_bitacora =  correl_bitacora + 1 where unidad_codigo = '" & Txt_campo1.Caption & "'  "
        
    End If
    'Set Ado_datos1.Recordset = rs_datos1
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

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc1_LostFocus()
    If (dtc_codigo1.Text = "2" Or dtc_codigo1.Text = "5") And (Txt_campo5.Text = "" Or Txt_campo5.Text = "0") Then
        BtnGrabar2.Visible = True
    Else
        BtnGrabar2.Visible = False
    End If
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc2_LostFocus()
    If dtc_codigo2.Text = "" Or dtc_codigo2.Text = "0" Then
        'tw_tecnico_bitacora.Ado_detalle1.Recordset!beneficiario_nombre_ref = IIf(txt_cliente = "", " ", txt_cliente)
        txt_cliente.Text = IIf(txt_cliente = "", " ", txt_cliente)
     Else
        'tw_tecnico_bitacora.Ado_detalle1.Recordset!beneficiario_nombre_ref = dtc_desc2.Text
        txt_cliente.Text = dtc_desc2.Text
     End If
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc4_LostFocus()
    If dtc_codigo4.Text = "NN" Then
        dtc_desc4.Visible = False
        Txt_campo3.Visible = True
    Else
        Txt_campo3.Text = dtc_desc4.Text
    End If
End Sub

Private Sub Form_Activate()
    Call ABRIR_TABLA
End Sub

Private Sub Form_Load()
    DTPfecha1.Value = Date
    dtc_desc4.Visible = True
    Txt_campo3.Visible = False
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
    'rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic   'order by descripcion
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    'rs_datos2.Open "Select * from gc_tipo_solicitud order by solicitud_tipo", db, adOpenStatic
    rs_datos2.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    '
    Select Case Aux
        Case "DNMAN"
            rs_datos3.Open "select * from rv_unidad_vs_responsable where (unidad_codigo = 'DNMAN' OR unidad_codigo = 'DMANS' OR unidad_codigo = 'DMANB' OR unidad_codigo = 'DMANC' OR unidad_codigo = 'DCOBR' OR unidad_codigo = 'DCONT' OR unidad_codigo = 'GADM') ORDER BY beneficiario_denominacion ", db, adOpenStatic
        Case "DNREP"
            rs_datos3.Open "select * from rv_unidad_vs_responsable where (unidad_codigo = 'DNREP' OR unidad_codigo = 'DREPS' OR unidad_codigo = 'DREPB' OR unidad_codigo = 'DREPC' OR unidad_codigo = 'DCOBR' OR unidad_codigo = 'DCONT' OR unidad_codigo = 'GADM') ORDER BY beneficiario_denominacion ", db, adOpenStatic
        Case "DNINS"
            rs_datos3.Open "select * from rv_unidad_vs_responsable where (unidad_codigo = 'DNINS' OR unidad_codigo = 'DINSS' OR unidad_codigo = 'DINSB' OR unidad_codigo = 'DINSC' OR unidad_codigo = 'DCOBR' OR unidad_codigo = 'DCONT' OR unidad_codigo = 'GADM') ORDER BY beneficiario_denominacion ", db, adOpenStatic
        Case "DNMOD"
            rs_datos3.Open "select * from rv_unidad_vs_responsable where (unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC' OR unidad_codigo = 'DCOBR' OR unidad_codigo = 'DCONT' OR unidad_codigo = 'GADM') ORDER BY beneficiario_denominacion ", db, adOpenStatic
        Case "DVTA"
            rs_datos3.Open "select * from rv_unidad_vs_responsable where (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC' OR unidad_codigo = 'DCOBR' OR unidad_codigo = 'DCONT' OR unidad_codigo = 'GADM') ORDER BY beneficiario_denominacion ", db, adOpenStatic
    End Select
    
    'rs_datos3.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & Aux & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    'rs_datos3.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic   'Txt_campo1
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText

    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from  rc_motivo_proceso where motivo_tipo = '10' ", db, adOpenStatic
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

Private Sub Txt_campo2_GotFocus()
    Txt_campo2A.Text = IIf(IsNull(Txt_campo2.Value), "09:00:00", Txt_campo2.Value)
    'Str(FormatDateTime(txt_campo2.Value, vbLongTime))
End Sub

Private Sub txt_campo3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_campo4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

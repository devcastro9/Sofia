VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form tw_bitacora_emergencia 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bitacora de Emergencias"
   ClientHeight    =   8535
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   10965
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
      TabIndex        =   55
      Top             =   120
      Width           =   10695
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "tw_bitacora_emergencia.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   72
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1275
         Picture         =   "tw_bitacora_emergencia.frx":07D6
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   71
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BITACORA DE EMERGENCIAS"
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
         Left            =   4530
         TabIndex        =   56
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      Height          =   7455
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   10695
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente que Realiza el Reclamo (Registre una de las 2 opciones...)"
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
         TabIndex        =   58
         Top             =   1080
         Width           =   10215
         Begin VB.TextBox Text1 
            DataField       =   "beneficiario_nombre_ref"
            DataSource      =   "tw_identificacion_cliente.ado_detalle1"
            Height          =   315
            Left            =   5160
            TabIndex        =   59
            Top             =   560
            Width           =   4935
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "tw_bitacora_emergencia.frx":10C2
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "tw_identificacion_cliente.ado_detalle1"
            Height          =   315
            Left            =   3720
            TabIndex        =   62
            Top             =   240
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
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "tw_bitacora_emergencia.frx":10DB
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "tw_identificacion_cliente.ado_detalle1"
            Height          =   315
            Left            =   120
            TabIndex        =   54
            Top             =   560
            Width           =   4950
            _ExtentX        =   8731
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label25 
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
            TabIndex        =   61
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
            TabIndex        =   60
            Top             =   300
            Width           =   2880
         End
      End
      Begin VB.TextBox Txt_campo10AA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "negocia_hora_trabajo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   285
         Left            =   8880
         TabIndex        =   52
         Text            =   "00:00"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Txt_campo8AA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "negocia_hora_mora"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "00:00"
         Top             =   4800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Txt_campo5 
         DataField       =   "bitacora_cite"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   285
         Left            =   6240
         TabIndex        =   30
         Text            =   "0"
         Top             =   6960
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "tw_bitacora_emergencia.frx":10F4
         DataField       =   "negocia_forma"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   4200
         TabIndex        =   17
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "negocia_forma"
         BoundColumn     =   "negocia_forma"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "tw_bitacora_emergencia.frx":110D
         DataField       =   "beneficiario_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   9240
         TabIndex        =   20
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin VB.TextBox Txt_campo4 
         DataField       =   "negocia_observaciones"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   6480
         Width           =   8565
      End
      Begin VB.TextBox Txt_campo3 
         DataField       =   "negocia_tarea_realizada"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   435
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   5880
         Width           =   8535
      End
      Begin VB.TextBox Txt_monto1 
         DataField       =   "negocia_gasto_estimado"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Text            =   "0"
         Top             =   6960
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "tw_bitacora_emergencia.frx":1126
         DataField       =   "negocia_forma"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   3000
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "negocia_forma_descripcion"
         BoundColumn     =   "negocia_forma"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "tw_bitacora_emergencia.frx":113F
         DataField       =   "beneficiario_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   4200
         TabIndex        =   19
         Top             =   2320
         Width           =   6135
         _ExtentX        =   10821
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
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "negocia_fecha_real"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   300
         Left            =   3120
         TabIndex        =   21
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   85262337
         CurrentDate     =   43101
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "negocia_fecha_real"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   300
         Left            =   8760
         TabIndex        =   34
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   85262337
         CurrentDate     =   43101
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "tw_bitacora_emergencia.frx":1158
         DataField       =   "tipo_falla"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   2640
         TabIndex        =   47
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tipo_falla"
         BoundColumn     =   "tipo_falla"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "tw_bitacora_emergencia.frx":1171
         DataField       =   "tipo_falla"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   360
         TabIndex        =   48
         Top             =   5400
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "tipo_falla_descripcion"
         BoundColumn     =   "tipo_falla"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "tw_bitacora_emergencia.frx":118A
         DataField       =   "falla_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   3960
         TabIndex        =   50
         Top             =   5400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "falla_codigo"
         BoundColumn     =   "falla_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "tw_bitacora_emergencia.frx":11A3
         DataField       =   "falla_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   5400
         TabIndex        =   51
         Top             =   5400
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "falla_descripcion"
         BoundColumn     =   "falla_codigo"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker Txt_campo2 
         DataField       =   "negocia_hora_real"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   300
         Left            =   360
         TabIndex        =   57
         Top             =   4320
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Format          =   85262338
         CurrentDate     =   0.375
      End
      Begin MSComCtl2.DTPicker Txt_campo6 
         DataField       =   "negocia_hora_envio"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   300
         Left            =   2160
         TabIndex        =   63
         Top             =   4320
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Format          =   85262338
         CurrentDate     =   0.375
      End
      Begin MSComCtl2.DTPicker Txt_campo7 
         DataField       =   "negocia_hora_llegada"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   300
         Left            =   3840
         TabIndex        =   64
         Top             =   4320
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Format          =   85262338
         CurrentDate     =   0.375
      End
      Begin MSComCtl2.DTPicker Txt_campo8 
         DataField       =   "negocia_hora_mora"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   300
         Left            =   5520
         TabIndex        =   65
         Top             =   4320
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   12632256
         CalendarTitleBackColor=   12632256
         Format          =   85262338
         CurrentDate     =   0.375
      End
      Begin MSComCtl2.DTPicker Txt_campo9 
         DataField       =   "negocia_hora_salida"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   300
         Left            =   7200
         TabIndex        =   66
         Top             =   4320
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Format          =   85262338
         CurrentDate     =   0.375
      End
      Begin MSComCtl2.DTPicker Txt_campo10 
         DataField       =   "negocia_hora_trabajo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   300
         Left            =   8880
         TabIndex        =   67
         Top             =   4320
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
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
         CalendarTitleBackColor=   16777215
         Format          =   85262338
         CurrentDate     =   0.375
      End
      Begin MSDataListLib.DataCombo dtc_codigo6 
         Bindings        =   "tw_bitacora_emergencia.frx":11BC
         DataField       =   "prioridad_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   9360
         TabIndex        =   68
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "prioridad_codigo"
         BoundColumn     =   "prioridad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc6 
         Bindings        =   "tw_bitacora_emergencia.frx":11D5
         DataField       =   "prioridad_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
         Height          =   315
         Left            =   5760
         TabIndex        =   69
         Top             =   3000
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "prioridad_descripcion"
         BoundColumn     =   "prioridad_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Prioridad de Atención"
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
         Left            =   5760
         TabIndex        =   70
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Falla        Descripción de la Falla"
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
         Left            =   3960
         TabIndex        =   49
         Top             =   5145
         Width           =   3585
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Falla"
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
         TabIndex        =   46
         Top             =   5145
         Width           =   1200
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Horas:Minutos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   7200
         TabIndex        =   45
         Top             =   4635
         Width           =   1410
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Horas:Minutos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8880
         TabIndex        =   44
         Top             =   4635
         Width           =   1410
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Horas:Minutos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   5595
         TabIndex        =   43
         Top             =   4635
         Width           =   1290
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Horas:Minutos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3840
         TabIndex        =   42
         Top             =   4635
         Width           =   1410
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Horas:Minutos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2160
         TabIndex        =   41
         Top             =   4635
         Width           =   1410
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Horas:Minutos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   360
         TabIndex        =   40
         Top             =   4635
         Width           =   1410
      End
      Begin VB.Label lbl_campo9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Salida"
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
         TabIndex        =   39
         Top             =   4080
         Width           =   1080
      End
      Begin VB.Label lbl_campo10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo.Trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   8880
         TabIndex        =   38
         Top             =   4080
         Width           =   1470
      End
      Begin VB.Label lbl_campo8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo.Mora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   5520
         TabIndex        =   37
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lbl_campo7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Llegada al Sitio"
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
         TabIndex        =   36
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label lbl_campo6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Envio al Tec."
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
         Left            =   2160
         TabIndex        =   35
         Top             =   4080
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Atención de la Llamada"
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
         TabIndex        =   33
         Top             =   3525
         Width           =   3015
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cod.Trámite"
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
         TabIndex        =   32
         Top             =   330
         Width           =   1110
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
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
         Left            =   4680
         TabIndex        =   31
         Top             =   6960
         Width           =   1485
      End
      Begin VB.Label Txt_campo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "unidad_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
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
         Left            =   6360
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         Left            =   360
         TabIndex        =   12
         Top             =   6465
         Width           =   1380
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Anormalidad Encontrada"
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
         Left            =   360
         TabIndex        =   29
         Top             =   5865
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Personal de CGI que Recibió el Reclamo"
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
         Top             =   2340
         Width           =   3690
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
         Left            =   360
         TabIndex        =   27
         Top             =   6960
         Width           =   1380
      End
      Begin VB.Label lbl_campo2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Recepción"
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
         TabIndex        =   26
         Top             =   4080
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Contacto del Cliente"
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
         TabIndex        =   25
         Top             =   3520
         Width           =   2685
      End
      Begin VB.Label Txt_descripcion 
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1920
         TabIndex        =   18
         Top             =   600
         Width           =   5655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7875
         TabIndex        =   16
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Txt_Correl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "bitacora_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
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
         Left            =   7920
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
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
         Left            =   360
         TabIndex        =   13
         Top             =   600
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
         Left            =   1920
         TabIndex        =   10
         Top             =   330
         Width           =   2160
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Contacto"
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
         TabIndex        =   9
         Top             =   2745
         Width           =   1545
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REG"
         DataField       =   "estado_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle1"
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
         Left            =   9480
         TabIndex        =   0
         Top             =   600
         Width           =   855
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
         Left            =   9600
         TabIndex        =   8
         Top             =   330
         Width           =   645
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
      ScaleWidth      =   10965
      TabIndex        =   1
      Top             =   8535
      Width           =   10965
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   8400
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
      Left            =   2280
      Top             =   8400
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
      Left            =   4440
      Top             =   8400
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
      Left            =   6600
      Top             =   8400
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
      Left            =   8760
      Top             =   8400
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
      Left            =   120
      Top             =   8760
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
End
Attribute VB_Name = "tw_bitacora_emergencia"
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
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
'BUSCADOR
Dim var_cod As String
Dim VAR_VAL As String

Dim VAR_DIA As Integer

'Dim VAR_DDIF As TIME TimeSpan
'Dim VAR_Fecha1 As DateTime
'Dim VAR_Fecha2 As DateTime
 
Dim mvBookMark As Variant
Dim mbDataChanged As Boolean


Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        tw_identificacion_cliente.Ado_detalle1.Recordset.CancelUpdate
        Unload Me
    End If
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    Select Case Txt_campo1.Caption
        Case "DRRHH"    'SOLO COMPRAS BB y SS
            If swnuevo = 1 Then
                
                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("ges_gestion").Value = glGestion  'Year(Date)
                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("unidad_codigo").Value = Txt_campo1.Caption
                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("solicitud_codigo").Value = txt_codigo.Caption
                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("estado_codigo").Value = "REG"
                Set rs_aux1 = New ADODB.Recordset
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & Txt_campo1.Caption & "' ", db, adOpenKeyset, adLockOptimistic
                If rs_aux1.RecordCount > 0 Then
                    var_cod = rs_aux1!correl_bitacora + 1
                Else
                    var_cod = 1
                End If
                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("bitacora_codigo").Value = var_cod
                'Actualiza correaltivo ...
                db.Execute "Update gc_unidad_ejecutora Set correl_bitacora = " & var_cod & " Where unidad_codigo = '" & Txt_campo1.Caption & "'   "
             End If
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_forma").Value = dtc_codigo1.Text
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_fecha_real").Value = DTPfecha1.Value
'             If HH.Text = "" Then
'                HH.Text = "00"
'                MM.Text = "00"
'             End If
'             If HH.Text = "" Then
'                HH.Text = "00"
'                MM.Text = "00"
'             End If
'             If HH.Text = "" Then
'                HH.Text = "00"
'                MM.Text = "00"
'             End If
'             If HH.Text = "" Then
'                HH.Text = "00"
'                MM.Text = "00"
'             End If
'             If HH.Text = "" Then
'                HH.Text = "00"
'                MM.Text = "00"
'             End If
'             If HH.Text = "" Then
'                HH.Text = "00"
'                MM.Text = "00"
'             End If
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_hora_real").Value = Txt_campo2.Value  ' Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_hora_envio").Value = Txt_campo6.Value   ' Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_hora_llegada").Value = Txt_campo7.Value   ' Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_hora_mora").Value = Txt_campo8.Value  '  Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_hora_salida").Value = Txt_campo9.Value  '  Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_hora_trabajo").Value = Txt_campo10.Value   ' Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_gasto_estimado").Value = Txt_monto1.Text
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("beneficiario_codigo").Value = dtc_codigo2.Text
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("beneficiario_codigo_resp").Value = dtc_codigo3.Text
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_tarea_realizada").Value = Txt_campo3.Text
             If swnuevo = 1 Then
                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_observaciones").Value = Trim(dtc_desc1.Text) + " - " + Txt_campo4.Text
             Else
                frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("negocia_observaciones").Value = Txt_campo4.Text
             End If
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("bitacora_cite").Value = Txt_campo5.Text
        
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("fecha_registro").Value = Date
             'frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("hora_registro").Value = Date
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset("usr_codigo").Value = glusuario
             frm_ao_solicitud_rrhh.Ado_detalle1.Recordset.UpdateBatch adAffectAll
        
        Case "2"    'SOLO VENTA DE BIENES
        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
            

        Case "DNMAN", "DNREP", "DNINS", "DNAJS", "DNEME", "DNMOD"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
             If swnuevo = 1 Then
                tw_identificacion_cliente.Ado_detalle1.Recordset("ges_gestion").Value = glGestion  'Year(Date)
                tw_identificacion_cliente.Ado_detalle1.Recordset("unidad_codigo").Value = Txt_campo1.Caption
                tw_identificacion_cliente.Ado_detalle1.Recordset("solicitud_codigo").Value = txt_codigo.Caption
                tw_identificacion_cliente.Ado_detalle1.Recordset("estado_codigo").Value = "REG"
                Set rs_aux1 = New ADODB.Recordset
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & Txt_campo1.Caption & "' ", db, adOpenKeyset, adLockOptimistic
                If rs_aux1.RecordCount > 0 Then
                    var_cod = rs_aux1!correl_bitacora + 1
                Else
                    var_cod = 1
                End If
                tw_identificacion_cliente.Ado_detalle1.Recordset("bitacora_codigo").Value = var_cod
                'Actualiza correaltivo ...
                db.Execute "Update gc_unidad_ejecutora Set correl_bitacora = " & var_cod & " Where unidad_codigo = '" & Txt_campo1.Caption & "'   "
             End If
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_forma").Value = dtc_codigo1.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_fecha_real").Value = DTPfecha1.Value
             'frm_to_id_emergencia.Ado_detalle1.Recordset("negocia_hora_real").Value = Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
             If Txt_campo2.Value = "00:00" Then
                Txt_campo2.Value = "08:00"
                'MM.Text = "00"
             End If

             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_hora_real").Value = Str(Txt_campo2.Value)  ' Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_hora_envio").Value = Str(Txt_campo6.Value)   ' Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_hora_llegada").Value = Str(Txt_campo7.Value)   ' Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_hora_mora").Value = Str(Txt_campo8.Value)  '  Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_hora_salida").Value = Str(Txt_campo9.Value)  '  Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_hora_trabajo").Value = Str(Txt_campo10.Value)   ' Trim(HH.Text) + ":" + Trim(MM.Text)   ' Txt_campo2.Text
             
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_gasto_estimado").Value = Txt_monto1.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("beneficiario_codigo").Value = dtc_codigo2.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("beneficiario_codigo_resp").Value = dtc_codigo3.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_tarea_realizada").Value = Txt_campo3.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("tipo_falla").Value = dtc_codigo4.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("falla_codigo").Value = dtc_codigo5.Text
             tw_identificacion_cliente.Ado_detalle1.Recordset("prioridad_codigo").Value = IIf(dtc_codigo6.Text = "", 6, dtc_codigo6.Text)
             
             If swnuevo = 1 Then
                tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_observaciones").Value = Trim(dtc_desc1.Text) + " - " + Txt_campo4.Text
             Else
                tw_identificacion_cliente.Ado_detalle1.Recordset("negocia_observaciones").Value = Txt_campo4.Text
             End If
             tw_identificacion_cliente.Ado_detalle1.Recordset("bitacora_cite").Value = Txt_campo5.Text
        
             tw_identificacion_cliente.Ado_detalle1.Recordset("fecha_registro").Value = Date
             'frm_to_id_emergencia.Ado_detalle1.Recordset("hora_registro").Value = Date
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

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
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

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
    Call pnivel4(dtc_codigo4.BoundText)
    dtc_desc5.Enabled = True
End Sub

Private Sub pnivel4(codigo4 As String)
   Dim strConsultaF As String
   strConsultaF = "select * from tc_fallas where tipo_falla = '" & codigo4 & "'"

   Set dtc_codigo5.RowSource = Nothing
   Set dtc_codigo5.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo7.RowSource = db.Execute("EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
   dtc_codigo5.ReFill
   dtc_codigo5.BoundText = Empty

   Set dtc_desc5.RowSource = Nothing
   Set dtc_desc5.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo7.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
   dtc_desc5.ReFill
   dtc_desc5.BoundText = Empty
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
End Sub

Private Sub DTPfecha1_LostFocus()
    Txt_campo2.Visible = True
    Txt_campo6.Visible = True
    DTPicker1.Visible = True
End Sub

Private Sub DTPicker1_LostFocus()
    Txt_campo7.Visible = True
    Txt_campo8.Visible = True
    Txt_campo9.Visible = True
    Txt_campo10.Visible = True
End Sub

Private Sub DTPicker4_LostFocus()
    'DTPicker5.Value = Format(DateDiff("n", Format(DTPicker3.Value, "Short Time"), Format(DTPicker4.Value, "Short Time")), "Short Time")
    DTPicker5.Value = Format(TimeValue(Format(DTPicker4.Value, "hh:mm")) - TimeValue(Format(DTPicker3.Value, "hh:mm")), "hh:mm")
    'horaentrada = Format("9:00:00", "hh:mm:ss")
    'horasalida = Format("16:30:20", "hh:mm:ss")

    'MsgBox "Tardaste: " & Format(TimeValue(horasalida) - TimeValue(horaentrada), "hh:mm:ss") & " horas"
    MsgBox "Tardaste: " & Format(TimeValue(Format(DTPicker4.Value, "hh:mm")) - TimeValue(Format(DTPicker3.Value, "hh:mm")), "hh:mm") & " horas"
End Sub

Private Sub DTPicker6_LostFocus()
    'Me.Print Format(DateDiff("y", Fecha_Inicial, Fecha_Final), Formato) & " dias"
    VAR_DIA = DateDiff("y", DTPfecha1.Value, DTPicker1.Value)
    If VAR_DIA = "0" Then
        DTPicker7.Value = Format(TimeValue(Format(DTPicker6.Value, "hh:mm")) - TimeValue(Format(DTPicker4.Value, "hh:mm")), "hh:mm")
        MsgBox "Tardaste: " & Format(TimeValue(Format(DTPicker6.Value, "hh:mm")) - TimeValue(Format(DTPicker4.Value, "hh:mm")), "hh:mm") & " horas"
    Else
        DTPicker7.Value = Format(TimeValue(Format(DTPicker6.Value, "hh:mm")) - TimeValue(Format(DTPicker4.Value, "hh:mm")), "hh:mm")
'        DTPicker7.Value = Format(TimeValue(Format(DTPicker7.Value, "hh:mm")) + TimeValue(Format("24:00:00", "hh:mm")), "hh:mm")

'        'VAR_Fecha1 = "2015-09-28 10:51:49.817"
'        'FORMAT(DTPfecha1.VALUE, "DD/MM/AAAA")
'        'Format(DTPicker4.Value, "hh:mm")
'        'VAR_Fecha2 = "2015-09-28 11:02:19.457"
'        VAR_Fecha1 = Format(DTPfecha1.Value, "DD/MM/AAAA") & " " & Format(DTPicker4.Value, "hh:mm")
'        VAR_Fecha1 = Format(DTPicker1.Value, "DD/MM/AAAA") & " " & Format(DTPicker6.Value, "hh:mm")
'
'
'        VAR_DDIF = VAR_Fecha2 - VAR_Fecha1
'         MsgBox "Tardaste: " & VAR_DDIF
    End If
End Sub

Private Sub Form_Activate()
    'var_cod = AUX
    'var_cod = "DRRHH"
    var_cod = tw_bitacora_emergencia.Txt_campo1.Caption
    Call ABRIR_TABLA
End Sub

Private Sub Form_Load()
    var_cod = Aux
    'var_cod = "DRRHH"
    'var_cod = tw_bitacora_emergencia.Txt_campo1.Caption
    Call ABRIR_TABLA
    mbDataChanged = False
'    If swnuevo = 2 Then
'        dtc_desc2.BoundText = dtc_codigo2.BoundText
'        dtc_desc3.BoundText = dtc_codigo3.BoundText
'    End If
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
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'tc_fallas_tipo
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "select * from tc_fallas_tipo ORDER BY tipo_falla ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText

    'tc_fallas
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "select * from tc_fallas ORDER BY tipo_falla ", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText

    ' tc_prioridad_atencion
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "select * from tc_prioridad_atencion  ", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    
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

Private Sub MM_LostFocus()
'    Txt_campo2.Value = Trim(HH) + ":" + Trim(MM)
End Sub

Private Sub Txt_campo2_LostFocus()
    DTPicker1.Visible = True
    DTPicker1.Value = DTPfecha1.Value
End Sub

Private Sub txt_campo3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_campo4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_campo6_LostFocus()
    If Txt_campo2.Value > Txt_campo6.Value Then
        MsgBox "La Fecha de: " + lbl_campo2.Caption + " Debe ser MENOR, a la Fecha de: " + lbl_campo6.Caption, vbExclamation, "Validación de datos"
        Txt_campo6.SetFocus
    End If
End Sub

Private Sub Txt_campo7_LostFocus()
'    'Txt_campo8.Text = Txt_campo7.Text - Txt_campo6.Text
'    'Format("17:08", "Short Time")
'    'Me.Print Format(DateDiff("n", Fecha_Inicial, Fecha_Final), Formato) & " minutos"
    
    'Txt_campo8.Text = DateDiff("n", Format(Txt_campo6.Text, "Short Time"), Format(Txt_campo7.Text, "Short Time"))

    Txt_campo8.Value = Format(TimeValue(Format(Txt_campo7.Value, "hh:mm")) - TimeValue(Format(Txt_campo6.Value, "hh:mm")), "hh:mm")
End Sub

Private Sub Txt_campo9_Change()
    VAR_DIA = DateDiff("y", DTPfecha1.Value, DTPicker1.Value)
    If VAR_DIA = "0" Then
        Txt_campo10.Value = Format(TimeValue(Format(Txt_campo9.Value, "hh:mm")) - TimeValue(Format(Txt_campo7.Value, "hh:mm")), "hh:mm")
        MsgBox "Tardaste: " & Format(TimeValue(Format(Txt_campo9.Value, "hh:mm")) - TimeValue(Format(Txt_campo7.Value, "hh:mm")), "hh:mm") & " horas"
    Else
        Txt_campo10.Value = Format(TimeValue(Format(Txt_campo9.Value, "hh:mm")) - TimeValue(Format(Txt_campo7.Value, "hh:mm")), "hh:mm")
'        DTPicker7.Value = Format(TimeValue(Format(DTPicker7.Value, "hh:mm")) + TimeValue(Format("24:00:00", "hh:mm")), "hh:mm")

'        'VAR_Fecha1 = "2015-09-28 10:51:49.817"
'        'FORMAT(DTPfecha1.VALUE, "DD/MM/AAAA")
'        'Format(DTPicker4.Value, "hh:mm")
'        'VAR_Fecha2 = "2015-09-28 11:02:19.457"
'        VAR_Fecha1 = Format(DTPfecha1.Value, "DD/MM/AAAA") & " " & Format(DTPicker4.Value, "hh:mm")
'        VAR_Fecha1 = Format(DTPicker1.Value, "DD/MM/AAAA") & " " & Format(DTPicker6.Value, "hh:mm")
'
'
'        VAR_DDIF = VAR_Fecha2 - VAR_Fecha1
'         MsgBox "Tardaste: " & VAR_DDIF
    End If
End Sub

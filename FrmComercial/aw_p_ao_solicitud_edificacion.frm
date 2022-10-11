VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form aw_p_ao_solicitud_edificacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Cotización - Detalle de la Edificación"
   ClientHeight    =   6315
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "aw_p_ao_solicitud_edificacion.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   10635
      TabIndex        =   20
      Top             =   120
      Width           =   10695
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   1200
         Picture         =   "aw_p_ao_solicitud_edificacion.frx":6C032
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   2520
         MaskColor       =   &H00000000&
         Picture         =   "aw_p_ao_solicitud_edificacion.frx":6C808
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETALLE DE LA EDIFICACION"
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
         Left            =   5250
         TabIndex        =   21
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      Height          =   4935
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   10695
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
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
         Height          =   1455
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   10455
         Begin VB.TextBox Text3 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   9840
            TabIndex        =   50
            Top             =   975
            Width           =   365
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   6135
            TabIndex        =   49
            Top             =   975
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "aw_p_ao_solicitud_edificacion.frx":6D0F4
            DataField       =   "edif_codigo"
            DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
            Height          =   315
            Left            =   960
            TabIndex        =   51
            Top             =   960
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "edif_descripcion"
            BoundColumn     =   "edif_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_aux1 
            Bindings        =   "aw_p_ao_solicitud_edificacion.frx":6D10E
            DataField       =   "edif_codigo"
            DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
            Height          =   315
            Left            =   7080
            TabIndex        =   52
            Top             =   960
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   16777215
            ListField       =   "edif_tipo_descripcion"
            BoundColumn     =   "edif_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux3 
            Bindings        =   "aw_p_ao_solicitud_edificacion.frx":6D129
            DataField       =   "edif_codigo"
            DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
            Height          =   315
            Left            =   8280
            TabIndex        =   55
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483629
            ListField       =   "munic_sigla"
            BoundColumn     =   "edif_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux2 
            Bindings        =   "aw_p_ao_solicitud_edificacion.frx":6D144
            DataField       =   "edif_codigo"
            DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
            Height          =   315
            Left            =   7080
            TabIndex        =   56
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483629
            ListField       =   "edif_tipo"
            BoundColumn     =   "edif_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "aw_p_ao_solicitud_edificacion.frx":6D15F
            DataField       =   "edif_codigo"
            DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
            Height          =   315
            Left            =   3600
            TabIndex        =   57
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "edif_codigo"
            BoundColumn     =   "edif_codigo"
            Text            =   ""
         End
         Begin VB.Label Txt_campo19 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "codigo"
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
            Left            =   960
            TabIndex        =   60
            Top             =   1320
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.Label Txt_campo20 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "0"
            DataField       =   "codigo1"
            DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
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
            Left            =   7800
            TabIndex        =   59
            Top             =   1320
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Txt_campo18 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "0"
            DataField       =   "codigo1"
            DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
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
            Left            =   5160
            TabIndex        =   58
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Txt_campo1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "0"
            DataField       =   "codigo1"
            DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
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
            Left            =   4680
            TabIndex        =   54
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label txt_gestion 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "0"
            DataField       =   "ges_gestion"
            DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
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
            Left            =   9120
            TabIndex        =   53
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   48
            Top             =   960
            Width           =   660
         End
         Begin VB.Label lbl_persona1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
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
            Left            =   6600
            TabIndex        =   47
            Top             =   960
            Width           =   420
         End
         Begin VB.Label Txt_estado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REG"
            DataField       =   "estado_codigo"
            DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
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
            Left            =   9120
            TabIndex        =   46
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label txt_codigo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "codigo2"
            DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
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
            Left            =   1080
            TabIndex        =   45
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Txt_descripcion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "codigo"
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
            Left            =   3360
            TabIndex        =   44
            Top             =   360
            Width           =   4695
         End
         Begin VB.Label lblLabels 
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
            Height          =   480
            Index           =   2
            Left            =   8280
            TabIndex        =   43
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl_codigo 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Número Tramite"
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
            TabIndex        =   42
            Top             =   240
            Width           =   870
         End
         Begin VB.Label lblLabels 
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
            Height          =   480
            Index           =   8
            Left            =   2400
            TabIndex        =   41
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.CommandButton BtnVer2 
         BackColor       =   &H80000015&
         Caption         =   "Cargar       Imagen"
         Height          =   360
         Left            =   8160
         Picture         =   "aw_p_ao_solicitud_edificacion.frx":6D17A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4320
         Width           =   2205
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H80000015&
         Caption         =   "Cargar       Imagen"
         Height          =   360
         Left            =   5640
         Picture         =   "aw_p_ao_solicitud_edificacion.frx":6D504
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4320
         Width           =   2205
      End
      Begin VB.PictureBox Img_Foto 
         Height          =   2055
         Left            =   5640
         ScaleHeight     =   1995
         ScaleWidth      =   2115
         TabIndex        =   29
         Top             =   2160
         Width           =   2175
         Begin VB.Image Image1 
            DataField       =   "foto"
            DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
            Height          =   1995
            Left            =   0
            Picture         =   "aw_p_ao_solicitud_edificacion.frx":6D88E
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2115
         End
      End
      Begin VB.PictureBox Img_Foto2 
         Height          =   2055
         Left            =   8160
         ScaleHeight     =   1995
         ScaleWidth      =   2115
         TabIndex        =   28
         Top             =   2160
         Width           =   2175
         Begin VB.Image Image2 
            DataField       =   "foto_bien"
            DataSource      =   "Ado_datos"
            Height          =   1995
            Left            =   0
            Picture         =   "aw_p_ao_solicitud_edificacion.frx":70538
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2115
         End
      End
      Begin VB.TextBox Txt_campo11 
         DataField       =   "edif_num_habit_dorm_4"
         DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Text            =   "0"
         Top             =   4215
         Width           =   1455
      End
      Begin VB.TextBox Txt_campo10 
         DataField       =   "edif_num_habit_dorm_3"
         DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
         Height          =   285
         Left            =   3960
         TabIndex        =   8
         Text            =   "0"
         Top             =   3495
         Width           =   1455
      End
      Begin VB.TextBox Txt_campo8 
         DataField       =   "edif_num_habit_ocupadas"
         DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Text            =   "0"
         Top             =   3495
         Width           =   1455
      End
      Begin VB.TextBox Txt_campo9 
         DataField       =   "edif_num_habit_dorm_2"
         DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Text            =   "0"
         Top             =   3495
         Width           =   1455
      End
      Begin VB.TextBox Txt_campo7 
         DataField       =   "edif_num_habit_libres"
         DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
         Height          =   285
         Left            =   3960
         TabIndex        =   5
         Text            =   "0"
         Top             =   2775
         Width           =   1455
      End
      Begin VB.TextBox Txt_campo5 
         DataField       =   "edif_num_salas_may_200m"
         DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Text            =   "0"
         Top             =   2775
         Width           =   1455
      End
      Begin VB.TextBox Txt_campo6 
         DataField       =   "edif_num_salas_men_200m"
         DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Text            =   "0"
         Top             =   2775
         Width           =   1455
      End
      Begin VB.TextBox Txt_campo4 
         DataField       =   "edif_num_pisos"
         DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
         Height          =   285
         Left            =   3960
         TabIndex        =   2
         Text            =   "0"
         Top             =   2055
         Width           =   1455
      End
      Begin VB.TextBox Txt_campo3 
         DataField       =   "edif_area_util_m2"
         DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Text            =   "0"
         Top             =   2055
         Width           =   1455
      End
      Begin VB.TextBox Txt_campo2 
         DataField       =   "edif_area_total_m2"
         DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
         Height          =   285
         Left            =   360
         TabIndex        =   0
         Text            =   "0"
         Top             =   2055
         Width           =   1455
      End
      Begin VB.Label Txt_campo13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "edif_capacidad_min_trafico"
         DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
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
         Left            =   3120
         TabIndex        =   39
         Top             =   4215
         Width           =   1455
      End
      Begin VB.Label Txt_campo12 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "0"
         DataField       =   "edif_indicador_min_trafico"
         DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
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
         Left            =   3960
         TabIndex        =   38
         Top             =   4440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Capacidad de Tráfico Mínima"
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
         TabIndex        =   37
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Indice Min.Tráfico"
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
         Left            =   2040
         TabIndex        =   36
         Top             =   4440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lbl_campo11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Dpto.>= 4 Dorm."
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
         TabIndex        =   35
         Top             =   3960
         Width           =   1425
      End
      Begin VB.Label lbl_campo7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Habit. Libres"
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
         TabIndex        =   34
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label lbl_campo6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Salas < 200 mt2"
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
         TabIndex        =   33
         Top             =   2520
         Width           =   1665
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Pisos"
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
         TabIndex        =   32
         Top             =   1800
         Width           =   1065
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Area Util mt2"
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
         TabIndex        =   31
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Plano Corte Transversal"
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
         Left            =   8205
         TabIndex        =   30
         Top             =   1880
         Width           =   2190
      End
      Begin VB.Label lbl_campo5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Salas >200 mt2"
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
         TabIndex        =   22
         Top             =   2520
         Width           =   1620
      End
      Begin VB.Label lbl_campo2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Area Total mt2"
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
         Top             =   1800
         Width           =   1305
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Plano Planta Tipo"
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
         Left            =   5955
         TabIndex        =   26
         Top             =   1880
         Width           =   1620
      End
      Begin VB.Label lbl_campo10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Dpto.de 3 Dorm."
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
         TabIndex        =   25
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lbl_campo9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Dpto.de 2 Dorm."
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
         TabIndex        =   24
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lbl_campo8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "NºHabit.Ocupadas"
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
         TabIndex        =   23
         Top             =   3240
         Width           =   1695
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
      TabIndex        =   14
      Top             =   6315
      Width           =   10935
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos01 
      Height          =   330
      Left            =   120
      Top             =   5880
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
End
Attribute VB_Name = "aw_p_ao_solicitud_edificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos01 As New ADODB.Recordset
Attribute rs_datos01.VB_VarHelpID = -1
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
'BUSCADOR
Dim var_cod As String
Dim VAR_VAL As String
Dim var_ctm, var_itm As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        mw_solicitud.Ado_detalle1.Recordset.CancelUpdate
        Unload Me
    End If
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
     If swnuevo = 1 Then
        mw_solicitud.Ado_detalle1.Recordset("ges_gestion").Value = txt_gestion 'Year(Date)
        mw_solicitud.Ado_detalle1.Recordset("unidad_codigo").Value = Txt_campo1.Caption
        mw_solicitud.Ado_detalle1.Recordset("solicitud_codigo").Value = txt_codigo.Caption
        mw_solicitud.Ado_detalle1.Recordset("estado_codigo").Value = "REG"
        mw_solicitud.Ado_detalle1.Recordset("archivo_foto_cargado").Value = "N"
        mw_solicitud.Ado_detalle1.Recordset("archivo_plano_cargado").Value = "N"
        mw_solicitud.Ado_detalle1.Recordset("edif_codigo").Value = dtc_codigo1.Text
     End If
'     Set rs_aux1 = New ADODB.Recordset
'     SQL_FOR = "select * from ao_solicitud_edificacion where unidad_codigo = '" & mw_solicitud.Ado_datos.Recordset("unidad_codigo") & "' and solicitud_codigo = " & mw_solicitud.Ado_datos.Recordset("solicitud_codigo") & " and edif_codigo = '" & dtc_codigo1.Text & "'  "
'     rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'     If rs_aux1.RecordCount > 0 Then
'        MsgBox "El código ya existe, consulte con el administrador del Sistema..."
'        var_cod = 0
'        Exit Sub
'     Else
'        mw_solicitud.Ado_detalle1.Recordset("edif_codigo").Value = dtc_codigo1.Text
'     End If
     mw_solicitud.Ado_detalle1.Recordset("edif_area_total_m2").Value = txt_campo2.Text
     mw_solicitud.Ado_detalle1.Recordset("edif_area_util_m2").Value = txt_campo3.Text
     mw_solicitud.Ado_detalle1.Recordset("edif_num_pisos").Value = txt_campo4.Text
     mw_solicitud.Ado_detalle1.Recordset("edif_num_salas_may_200m").Value = txt_campo5.Text
     mw_solicitud.Ado_detalle1.Recordset("edif_num_salas_men_200m").Value = txt_campo6.Text
     mw_solicitud.Ado_detalle1.Recordset("edif_num_habit_libres").Value = txt_campo7.Text
     mw_solicitud.Ado_detalle1.Recordset("edif_num_habit_ocupadas").Value = txt_campo8.Text
     mw_solicitud.Ado_detalle1.Recordset("edif_num_habit_dorm_2").Value = txt_campo9.Text
     mw_solicitud.Ado_detalle1.Recordset("edif_num_habit_dorm_3").Value = txt_campo10.Text
     mw_solicitud.Ado_detalle1.Recordset("edif_num_habit_dorm_4").Value = Txt_campo11.Text
     Select Case dtc_aux2.Text
        Case "DPTO"
            var_itm = Round(Val(txt_campo8.Text) * 2 + Val(txt_campo9.Text) * 4 + Val(txt_campo10.Text) * 5 + Val(Txt_campo11.Text) * 6 + Val(txt_campo7), 2)
            var_ctm = Round(var_itm * 0.1, 2)
        Case "OFIG"
            var_itm = Round((Val(txt_campo2.Text) - Val(txt_campo3.Text) - Val(txt_campo5.Text)) / 7 + Val(txt_campo5.Text) * 0.85 / 7, 2)
            var_ctm = Round(var_itm * 0.12, 2)
        Case "OFIU"
            var_itm = Round((Val(txt_campo2.Text) - Val(txt_campo3.Text) - Val(txt_campo5.Text)) / 7 + Val(txt_campo5.Text) * 0.85 / 7, 2)
            var_ctm = Round(var_itm * 0.15, 2)
        Case "COMR"
            var_itm = Round((Val(txt_campo2.Text) - Val(txt_campo3.Text) - Val(txt_campo5.Text)) / 4 + Val(txt_campo5.Text) * 0.85 / 4, 2)
            var_ctm = Round(var_itm * 0.1, 2)
        Case "EDUC"
            '=+A29/2+B29/7+C29*0.85
            var_itm = Round((Val(txt_campo2.Text) / 2 + Val(txt_campo3.Text) / 7 + Val(txt_campo5.Text) * 0.85), 2)
            var_ctm = Round(var_itm * 0.2, 2)
        Case "HOTL"
            'var_itm = Round(Val(Txt_campo8.Text) * 0.2, 2)
            var_itm = Round(Val(txt_campo8.Text) * 2, 2)
            var_ctm = Round(var_itm * 0.1, 2)
        Case "REST"
            var_itm = Round(Val(txt_campo3.Text) / 1.5, 2)
            var_ctm = Round(var_itm * 0.06, 2)
        Case "HOSP"
            var_itm = Round(Val(txt_campo8.Text) * 2.5, 2)
            var_ctm = Round(var_itm * 0.08, 2)
        Case "GARJ"
            var_itm = Round(Val(txt_campo8.Text) * 1.4, 2)
            var_ctm = Round(var_itm * 0.1, 2)
     End Select
     Txt_campo12.Caption = var_itm
     Txt_campo13.Caption = var_ctm
     mw_solicitud.Ado_detalle1.Recordset("edif_indicador_min_trafico").Value = var_itm
     mw_solicitud.Ado_detalle1.Recordset("edif_capacidad_min_trafico").Value = var_ctm
     
     mw_solicitud.Ado_detalle1.Recordset("edif_dimension_frente1").Value = "0"     'Txt_campo14.Text
     mw_solicitud.Ado_detalle1.Recordset("edif_dimension_fondo1").Value = "0"     'Txt_campo15.Text
     mw_solicitud.Ado_detalle1.Recordset("edif_dimension_frente2").Value = "0"     'Txt_campo16.Text
     mw_solicitud.Ado_detalle1.Recordset("edif_dimension_fondo2").Value = "0"     'Txt_campo17.Text
     
     mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
     mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
     mw_solicitud.Ado_detalle1.Recordset("fecha_registro").Value = Date
     'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
     mw_solicitud.Ado_detalle1.Recordset("usr_codigo").Value = glusuario
     mw_solicitud.Ado_detalle1.Recordset.UpdateBatch adAffectAll
     
     var_cod = mw_solicitud.Ado_detalle1.Recordset.RecordCount
     db.Execute "Update ao_solicitud Set correl_edificacion = " & var_cod & " Where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "  "
     
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
    MsgBox "Debe registrar el " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If txt_campo4.Text = "" Or txt_campo4.Text = "0" Then
    MsgBox "Debe registrar el " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  Select Case dtc_aux2.Text
    Case "DPTO"
        If (txt_campo8.Text = "") Or (txt_campo9.Text = "") Or (txt_campo10.Text = "") Or (Txt_campo11.Text = "") Then
          MsgBox "Verifique los datos de : " + lbl_campo8.Caption + " o " + lbl_campo9.Caption + " o " + lbl_campo10.Caption + " o " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
        If (txt_campo8.Text = "0") And (txt_campo9.Text = "0") And (txt_campo10.Text = "0") And (Txt_campo11.Text = "0") Then
          MsgBox "Debe registrar uno de los datos de : " + lbl_campo8.Caption + " o " + lbl_campo9.Caption + " o " + lbl_campo10.Caption + " o " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
''        If Txt_campo7.Text = "" Or Txt_campo7.Text = "0" Then
''          MsgBox "Debe registrar : " + lbl_campo7.Caption, vbCritical + vbExclamation, "Validación de datos"
''          VAR_VAL = "ERR"
''          Exit Sub
''        End If
'        If Txt_campo8.Text = "" Or Txt_campo8.Text = "0" Then
'          MsgBox "Debe registrar : " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo9.Text = "" Or Txt_campo9.Text = "0" Then
'          MsgBox "Debe registrar : " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo10.Text = "" Or Txt_campo10.Text = "0" Then
'          MsgBox "Debe registrar : " + lbl_campo10.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo11.Text = "" Or Txt_campo11.Text = "0" Then
'          MsgBox "Debe registrar : " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
    Case "OFIG"
        If txt_campo2.Text = "" Or txt_campo2.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
        If txt_campo3.Text = "" Or txt_campo3.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
        If txt_campo5.Text = "" Or txt_campo5.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
    Case "OFIU"
        If txt_campo2.Text = "" Or txt_campo2.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
        If txt_campo3.Text = "" Or txt_campo3.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
        If txt_campo5.Text = "" Or txt_campo5.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
    Case "COMR"
        If txt_campo2.Text = "" Or txt_campo2.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
        If txt_campo3.Text = "" Or txt_campo3.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
        If txt_campo5.Text = "" Or txt_campo5.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
    Case "EDUC"
        If txt_campo2.Text = "" Or txt_campo2.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
        If txt_campo3.Text = "" Or txt_campo3.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
        If txt_campo5.Text = "" Or txt_campo5.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
    Case "HOTL"
        If txt_campo8.Text = "" Or txt_campo8.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
    Case "REST"
        If txt_campo3.Text = "" Or txt_campo3.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
    Case "HOSP"
        If txt_campo8.Text = "" Or txt_campo8.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
    Case "GARJ"
        If txt_campo8.Text = "" Or txt_campo8.Text = "0" Then
          MsgBox "Debe registrar : " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
        End If
  End Select
'     Txt_campo12.Caption = var_itm
'     Txt_campo13.Caption = var_ctm
  
'        If Txt_campo2.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo3.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo4.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo5.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo6.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo7.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo7.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo8.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo9.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo10.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo10.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo11.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
End Sub

Private Sub BtnVer_Click()
'  On Error GoTo QError
'  If mw_solicitud.Ado_detalle1.Recordset("estado_codigo") = "REG" Then
'    Dim ARCH_FOTO As String
'    Dim SW0 As String
'    If mw_solicitud.Ado_detalle1.Recordset!archivo_foto_cargado = "N" Then
'      NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(dtc_aux3.Text) & "\" & Trim(dtc_codigo1.Text) & "\"
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "FED2"
''      If GlServidor = "SRVPRO" Then
''         e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
''      Else
'         e = NombreCarpeta
''      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
'      SW0 = 1
'    Else
'      'MsgBox ""
'      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'          NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(dtc_aux3.Text) & "\" & Trim(dtc_codigo1.Text) & "\"
'          Frmexporta.DirDestino.Path = NombreCarpeta
'          GlArch = "FED2"
''          If GlServidor = "SRVPRO" Then
''            e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
''          Else
'            e = NombreCarpeta
''          End If
'          Frmexporta.DirDestino2.Path = e
'          Frmexporta.Show vbModal
'          SW0 = 1
'      Else
'        SW0 = 0
'      End If
'    End If
'    If SW0 = 1 Then
'    '    If GlServidor = "SRVPRO" Then
'    '        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    '    Else
'            ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(dtc_aux3.Text) + "\" + Trim(dtc_codigo1.Text) + "\" + Trim(dtc_codigo1.Text) + "-A.JPG"
'    '    End If
'        'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + Ado_datos.Recordset!codigo_beneficiario + "\" + Ado_datos.Recordset("codigo_beneficiario") + "-FOTO.JPG"
'        CodBien = mw_solicitud.Ado_detalle1.Recordset!edif_codigo
'        'If Guardar_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
'        If Guardar_Imagen(db, "Select Foto From ao_solicitud_edificacion Where unidad_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("unidad_codigo") & "' and solicitud_codigo = " & mw_solicitud.Ado_detalle1.Recordset("solicitud_codigo") & " and edif_codigo = '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
'            MsgBox "Se cargo la Imagen Correctamente !!"
'        Else
'            MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
'        End If
'    Else
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("edif_codigo") & "' ", "Foto")
'        Image2 = Img_Foto
'    End If
'  Else
'    MsgBox "Debe Aprobar el registro, para crear la carpeta correspondiente..."
'  End If
'QError:
'    ' Manejo de errores
'    If Err.Number > 0 Then
'        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
'    '    db.RollbackTrans
'        Screen.MousePointer = vbDefault
'    End If
End Sub

Private Sub dtc_aux1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_aux1.BoundText
    dtc_desc1.BoundText = dtc_aux1.BoundText
    dtc_aux2.BoundText = dtc_aux1.BoundText
    dtc_aux3.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_aux2.BoundText
    dtc_desc1.BoundText = dtc_aux2.BoundText
    dtc_aux1.BoundText = dtc_aux2.BoundText
    dtc_aux3.BoundText = dtc_aux2.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_aux3.BoundText
    dtc_desc1.BoundText = dtc_aux3.BoundText
    dtc_aux2.BoundText = dtc_aux3.BoundText
    dtc_aux1.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    dtc_aux2.BoundText = dtc_codigo1.BoundText
    dtc_aux3.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    dtc_aux2.BoundText = dtc_desc1.BoundText
    dtc_aux3.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc1_LostFocus()
    Select Case dtc_aux2.Text
        Case "DPTO", "RECI"
            lbl_campo8.Caption = "Depto.de 1 Dorm."
            lbl_campo7.Caption = "NºHabit.Servicio"
        Case "OFIG"
            lbl_campo3.Caption = "Área Pasillos"
        Case "OFIU"
            lbl_campo3.Caption = "Área Pasillos"
        Case "COMR"
            lbl_campo3.Caption = "Área Pasillos"
        Case "EDUC"
            lbl_campo2.Caption = "Área Aulas"
            lbl_campo3.Caption = "Área Admin."
        Case "HOTL"
            lbl_campo8.Caption = "NºDormitorios"
        Case "REST"
            lbl_campo3.Caption = "Área Comedor"
        Case "HOSP", "CLIN"
            lbl_campo8.Caption = "Nº de Camas"
        Case "HOSS"
            lbl_campo8.Caption = "Nº de Camas"
        Case "GARJ"
            lbl_campo8.Caption = "Nºde Parqueos"
        Case "MIXT"
            lbl_campo8.Caption = "Depto.de 1 Dorm."
            lbl_campo7.Caption = "NºHabit.Servicio"
            lbl_campo3.Caption = "Área Pasillos"
     End Select
End Sub

Private Sub Form_Activate()
    Call ABRIR_TABLA
'    GlEdificio
'    aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText = GlEdificio
'    'aw_p_ao_solicitud_edificacion.dtc_codigo1.Text = mw_solicitud.dtc_codigo3.Text
'            aw_p_ao_solicitud_edificacion.dtc_desc1.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
'            aw_p_ao_solicitud_edificacion.dtc_aux1.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
'            aw_p_ao_solicitud_edificacion.dtc_aux2.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
'            aw_p_ao_solicitud_edificacion.dtc_aux3.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
    
End Sub

Private Sub Form_Load()
    'Call ABRIR_TABLA
    mbDataChanged = False
'    If swnuevo = 2 Then
'        dtc_desc2.BoundText = dtc_codigo2.BoundText
'        dtc_desc3.BoundText = dtc_codigo3.BoundText
'    End If
    If mw_solicitud.Ado_detalle1.Recordset("archivo_foto_cargado") = "S" Then
        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_solicitud_edificacion Where unidad_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("unidad_codigo") & "' and solicitud_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("solicitud_codigo") & "' and edif_codigo = '" & GlEdificio & "' ", "Foto")
        'Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_solicitud_edificacion Where unidad_codigo = '" & parametro & "' and solicitud_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("solicitud_codigo") & "' and edif_codigo = '" & GlEdificio & "' ", "Foto")
        Image1 = Img_Foto
    End If
    If mw_solicitud.Ado_detalle1.Recordset("archivo_plano_cargado") = "S" Then
        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_solicitud_edificacion Where unidad_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("unidad_codigo") & "' and solicitud_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("solicitud_codigo") & "' edif_codigo = '" & GlEdificio & "' ", "Foto1")
        Image2 = Img_Foto
    End If
'    mw_solicitud.Ado_detalle1.Recordset("ges_gestion").Value = Year(Date)
'        mw_solicitud.Ado_detalle1.Recordset("unidad_codigo").Value = txt_campo1.Caption
'        mw_solicitud.Ado_detalle1.Recordset("solicitud_codigo").Value = txt_codigo.Caption
'        mw_solicitud.Ado_detalle1.Recordset("estado_codigo").Value = "REG"
'        mw_solicitud.Ado_detalle1.Recordset("archivo_foto_cargado").Value = "N"
'        mw_solicitud.Ado_detalle1.Recordset("archivo_plano_cargado").Value = "N"
'        mw_solicitud.Ado_detalle1.Recordset("edif_codigo").Value = dtc_codigo1.Text
End Sub

Private Sub ABRIR_TABLA()
    Set rs_datos01 = New ADODB.Recordset
    If rs_datos01.State = 1 Then rs_datos01.Close
    'rs_datos01.Open "Select * from gv_edificaciones_tipo where beneficiario_codigo = '" & mw_solicitud.dtc_codigo4 & "' ", db, adOpenStatic
    rs_datos01.Open "Select * from gv_edificaciones_tipo where edif_codigo = '" & GlEdificio & "' ", db, adOpenStatic
    'rs_datos01.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic   'order by descripcion
    Set Ado_datos01.Recordset = rs_datos01
    dtc_codigo1.Text = GlEdificio
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    dtc_aux2.BoundText = dtc_codigo1.BoundText
    dtc_aux3.BoundText = dtc_codigo1.BoundText
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

Private Sub Txt_campo2_Click()
    Call dtc_desc1_LostFocus
End Sub

'Private Sub txt_campo3_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub Txt_campo4_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub Txt_campo4_Click()
    Call dtc_desc1_LostFocus
End Sub

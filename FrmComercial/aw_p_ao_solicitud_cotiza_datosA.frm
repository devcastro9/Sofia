VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form aw_p_ao_solicitud_cotiza_datosA 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cotización Venta - Datos Complementarios Cotiza (Asia)"
   ClientHeight    =   5970
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11025
   ControlBox      =   0   'False
   Icon            =   "aw_p_ao_solicitud_cotiza_datosA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Fra_datos 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2835
      ScaleWidth      =   10380
      TabIndex        =   33
      Top             =   2280
      Width           =   10440
      Begin VB.TextBox Text8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   310
         Left            =   9120
         TabIndex        =   34
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox Txt_campo7 
         DataField       =   "bien_cotiza_num_accesos"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   3120
         TabIndex        =   8
         Text            =   "0"
         Top             =   2325
         Width           =   1365
      End
      Begin VB.TextBox Txt_campo2 
         DataField       =   "cotiza_energia"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Text            =   "0"
         Top             =   1515
         Width           =   1365
      End
      Begin VB.TextBox Txt_campo3 
         DataField       =   "cotiza_luz"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   285
         Left            =   3120
         TabIndex        =   7
         Text            =   "0"
         Top             =   1920
         Width           =   1365
      End
      Begin VB.TextBox Txt_campo10 
         DataField       =   "dimension_fosa_frente"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   285
         Left            =   8280
         TabIndex        =   10
         Text            =   "0"
         Top             =   1920
         Width           =   1365
      End
      Begin VB.TextBox Txt_campo9 
         DataField       =   "dimension_fosa_fondo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   285
         Left            =   8280
         TabIndex        =   9
         Text            =   "0"
         Top             =   1515
         Width           =   1365
      End
      Begin VB.TextBox Txt_campo8 
         DataField       =   "dimension_fosa_m"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   8280
         TabIndex        =   11
         Text            =   "0"
         Top             =   2325
         Width           =   1365
      End
      Begin VB.TextBox Txt_campo4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "modelo_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   320
         Width           =   2055
      End
      Begin VB.TextBox Txt_campo5 
         DataField       =   "cotiza_nro_montador"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   8280
         TabIndex        =   3
         Text            =   "0"
         Top             =   320
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo dtc_desc7 
         Bindings        =   "aw_p_ao_solicitud_cotiza_datosA.frx":0A02
         DataField       =   "pais_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   4560
         TabIndex        =   2
         Top             =   315
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "pais_descripcion"
         BoundColumn     =   "pais_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "aw_p_ao_solicitud_cotiza_datosA.frx":0A1B
         DataField       =   "tipo_eqp"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   600
         TabIndex        =   4
         Top             =   1005
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "tipo_eqp_descripcion"
         BoundColumn     =   "tipo_eqp"
         Text            =   "36NO"
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
      Begin MSDataListLib.DataCombo dtc_codigo21 
         Bindings        =   "aw_p_ao_solicitud_cotiza_datosA.frx":0A35
         DataField       =   "bien_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   600
         TabIndex        =   0
         Top             =   315
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "bien_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   "36NO"
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
      Begin MSDataListLib.DataCombo dtc_desc21 
         Bindings        =   "aw_p_ao_solicitud_cotiza_datosA.frx":0A50
         DataField       =   "bien_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   0
         TabIndex        =   35
         Top             =   465
         Visible         =   0   'False
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "bien_descripcion"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo7 
         Bindings        =   "aw_p_ao_solicitud_cotiza_datosA.frx":0A6A
         DataField       =   "pais_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   7200
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "pais_codigo"
         BoundColumn     =   "pais_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc24 
         Bindings        =   "aw_p_ao_solicitud_cotiza_datosA.frx":0A83
         DataField       =   "bien_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   0
         TabIndex        =   37
         Top             =   225
         Visible         =   0   'False
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "modelo_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_desc61 
         Bindings        =   "aw_p_ao_solicitud_cotiza_datosA.frx":0A9D
         DataField       =   "cuadro_ctrl_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   5280
         TabIndex        =   5
         Top             =   1005
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cuadro_ctrl_descripcion"
         BoundColumn     =   "cuadro_ctrl_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo61 
         Bindings        =   "aw_p_ao_solicitud_cotiza_datosA.frx":0AB7
         DataField       =   "cuadro_ctrl_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   8640
         TabIndex        =   38
         Top             =   720
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cuadro_ctrl_codigo"
         BoundColumn     =   "cuadro_ctrl_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "aw_p_ao_solicitud_cotiza_datosA.frx":0AD2
         DataField       =   "tipo_eqp"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   3480
         TabIndex        =   39
         Top             =   720
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "tipo_eqp"
         BoundColumn     =   "tipo_eqp"
         Text            =   "0"
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cód.del Equipo          Modelo del Equipo          País de Origen del Equipo                            Nro.Montadores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Left            =   600
         TabIndex        =   32
         Top             =   45
         Width           =   9075
      End
      Begin VB.Label lbl_campo7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Accesos                                                                  Espacio Libre Bajo Dintel (mm)"
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
         Left            =   600
         TabIndex        =   44
         Top             =   2340
         Width           =   7575
      End
      Begin VB.Label lbl_campo6 
         BackColor       =   &H00000000&
         Caption         =   "Cuarto de Control"
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
         Height          =   360
         Left            =   5280
         TabIndex        =   43
         Top             =   705
         Width           =   1905
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Iluminación / Luz (V)                                                                        Dimención Fosa Frente (mm)"
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
         Left            =   600
         TabIndex        =   42
         Top             =   1935
         Width           =   7575
      End
      Begin VB.Label lbl_campo2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fuerza Motriz / Energía (V)                                                           Dimención Fosa Fondo (mm)"
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
         Left            =   600
         TabIndex        =   41
         Top             =   1545
         Width           =   7590
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Tipo de Equipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   600
         TabIndex        =   40
         Top             =   705
         Width           =   1755
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "aw_p_ao_solicitud_cotiza_datosA.frx":0AEB
      ScaleHeight     =   915
      ScaleWidth      =   10635
      TabIndex        =   24
      Top             =   120
      Width           =   10695
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   720
         Picture         =   "aw_p_ao_solicitud_cotiza_datosA.frx":6CB1D
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   2160
         MaskColor       =   &H00000000&
         Picture         =   "aw_p_ao_solicitud_cotiza_datosA.frx":6CD27
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATOS COMPLEMENTARIOS COTIZACION"
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
         Left            =   3870
         TabIndex        =   25
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame Fra_datos99 
      BackColor       =   &H00000000&
      Height          =   4215
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   10695
      Begin VB.Label txt_conti 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         DataField       =   "pais_continente"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.ado_datos"
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
         Left            =   5760
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         DataField       =   "unidad_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.ado_datos"
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
         Left            =   3720
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
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
         TabIndex        =   31
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nº Cotización"
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
         Left            =   7080
         TabIndex        =   30
         Top             =   330
         Width           =   1200
      End
      Begin VB.Label Txt_Correl 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "cotiza_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
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
         Left            =   6960
         TabIndex        =   29
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
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
         TabIndex        =   28
         Top             =   600
         Width           =   1335
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
         TabIndex        =   26
         Top             =   330
         Width           =   2160
      End
      Begin VB.Label lbl_codigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nº de Trámite "
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
         TabIndex        =   23
         Top             =   330
         Width           =   1290
      End
      Begin VB.Label Txt_campo2A 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "36NO-"
         DataField       =   "edif_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
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
         Left            =   8760
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Código Edificio"
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
         TabIndex        =   22
         Top             =   330
         Width           =   1365
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
      ScaleWidth      =   11025
      TabIndex        =   15
      Top             =   5970
      Width           =   11025
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   20
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   2400
      Top             =   5400
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
   Begin MSAdodcLib.Adodc Ado_datos21 
      Height          =   330
      Left            =   2400
      Top             =   5520
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
   Begin MSAdodcLib.Adodc Ado_datos7 
      Height          =   330
      Left            =   4680
      Top             =   5520
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   6960
      Top             =   5520
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
   Begin MSAdodcLib.Adodc Ado_datos61 
      Height          =   330
      Left            =   9240
      Top             =   5520
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
End
Attribute VB_Name = "aw_p_ao_solicitud_cotiza_datosA"
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
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
'BUSCADOR

'OTROS
Dim var_cod As String
Dim VAR_VAL As String

Dim VAR_1A, VAR_2A As Double
Dim VAR_3B, VAR_4B, VAR_5B, VAR_6B, VAR_7B As Double
Dim VAR_8C, VAR_9C, VAR_10C, VAR_11C, VAR_12C As Double
Dim VAR_13D, VAR_14D As Double
Dim totbs2, totdl2, totbs3, totdl3 As Double
Dim VAR_SUBD, VAR_SUBB As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset.CancelUpdate
        Unload Me
    End If
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & txt_conti & "'   ", db, adOpenKeyset, adLockOptimistic
    If rs_datos10.RecordCount > 0 Then
      If Txt_Correl.Caption = 1 Then
        sino = MsgBox("SI (Graba todos los Registros) - NO (Graba SOLO el Registro Activo) ... ", vbYesNo + vbQuestion, "Atención")
        If sino = vbYes Then
           'TODOS LOS REGISTROS
           aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset.MoveFirst
           While Not aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset.EOF
             '-
             Set aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset = rs_datos10
             aw_p_ao_solicitud_cotiza_venta.txt_codigo1.Caption = aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_codigo
             MsgBox Str(aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_codigo)
             'aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!pais_continente = "AMERICA"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!pais_codigo = dtc_codigo7.Text
    '         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!bien_codigo = IIf(IsNull(dtc_codigo21.Text) Or dtc_codigo21.Text = "", "NA1", dtc_codigo21.Text)
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_fecha = Date     'DTPfecha1.Value
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!modelo_codigo = txt_campo4.Text     '      'MODELO PROVISIONAL AUTOMATICO JQA 02-2015
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!modelo_codigo_h = "S/M"  'Txt_campo5.Text    'dtc_codigo41.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!modelo_codigo_x = "S/M"   'Txt_campo6.Text    'dtc_codigo51.Text
        
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_energia = txt_campo2.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_luz = txt_campo3.Text
    
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!bien_cotiza_num_accesos = txt_campo7.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!dimension_fosa_m = txt_campo8.Text        '
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!dimension_fosa_fondo = txt_campo9.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!dimension_fosa_frente = txt_campo10.Text  '
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cuadro_ctrl_codigo = IIf((dtc_codigo61.Text = ""), 1, dtc_codigo61.Text)
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_nro_montador = IIf((txt_campo5.Text = ""), "2", txt_campo5.Text)
             'aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!Foto = Date
             'aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!ARCHIVO_Foto = var_cod + ".JPG"
             'aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!archivo_foto_cargado = "N"
             'hora_registro
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!proceso_codigo = "COM"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!subproceso_codigo = "COM-01"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!etapa_codigo = "COM-01-04"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!clasif_codigo = "COM"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!doc_codigo = "R-222"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!doc_numero = "0"  'txt_campo1.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!clasif_codigo2 = "COM"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!doc_codigo2 = "R-224"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!doc_numero2 = "0"  'txt_campo1.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!poa_codigo = "3.1.1"
             'WWWWWWWWWWWWWW JQA 02-2015
                If VAR_COD2 < 10 Then
                   aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-00000" + Trim(txt_codigo)
                End If
                If VAR_COD2 > 9 And VAR_COD2 < 100 Then
                   aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-0000" + Trim(txt_codigo)
                End If
                If VAR_COD2 > 99 And VAR_COD2 < 1000 Then
                   aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-000" + Trim(txt_codigo)
                End If
                If VAR_COD2 > 999 And VAR_COD2 < 10000 Then
                   aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-00" + Trim(txt_codigo)
                End If
                If VAR_COD2 > 9999 And VAR_COD2 < 100000 Then
                   aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-0" + Trim(txt_codigo)
                End If
                If VAR_COD2 > 99999 Then
                   aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-" + Trim(txt_codigo)
                End If
             'aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = var_cod2     'txt_codigo1.Text
             'WWWWWWWWWWWWWW JQA 02-2015
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!tipo_eqp = dtc_codigo2.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!fecha_registro = Date     'no cambia
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset.Update 'adAffectAll    'Batch
             'costos
           aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset.MoveNext
           Wend
        Else
             'SOLO EL REGISTRO ACTIVO
             MsgBox Str(aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_codigo)
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!pais_codigo = dtc_codigo7.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!bien_codigo = IIf(IsNull(dtc_codigo21.Text) Or dtc_codigo21.Text = "", "NA1", dtc_codigo21.Text)
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_fecha = Date     'DTPfecha1.Value
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!modelo_codigo = txt_campo4.Text     '      'MODELO PROVISIONAL AUTOMATICO JQA 02-2015
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!modelo_codigo_h = "S/M"  'Txt_campo5.Text    'dtc_codigo41.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!modelo_codigo_x = "S/M"   'Txt_campo6.Text    'dtc_codigo51.Text
        
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_energia = txt_campo2.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_luz = txt_campo3.Text
    
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!bien_cotiza_num_accesos = txt_campo7.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!dimension_fosa_m = txt_campo8.Text        '
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!dimension_fosa_fondo = txt_campo9.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!dimension_fosa_frente = txt_campo10.Text  '
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cuadro_ctrl_codigo = IIf((dtc_codigo61.Text = ""), 1, dtc_codigo61.Text)
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_nro_montador = IIf((txt_campo5.Text = ""), "2", txt_campo5.Text)
             'aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!Foto = Date
             'aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!ARCHIVO_Foto = var_cod + ".JPG"
             'aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!archivo_foto_cargado = "N"
             'hora_registro
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!proceso_codigo = "COM"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!subproceso_codigo = "COM-01"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!etapa_codigo = "COM-01-04"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!clasif_codigo = "COM"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!doc_codigo = "R-222"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!doc_numero = "0"  'txt_campo1.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!clasif_codigo2 = "COM"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!doc_codigo2 = "R-224"
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!doc_numero2 = "0"  'txt_campo1.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!poa_codigo = "3.1.1"
             'WWWWWWWWWWWWWW JQA 02-2015
                If VAR_COD2 < 10 Then
                   aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-00000" + Trim(txt_codigo)
                End If
                If VAR_COD2 > 9 And VAR_COD2 < 100 Then
                   aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-0000" + Trim(txt_codigo)
                End If
                If VAR_COD2 > 99 And VAR_COD2 < 1000 Then
                   aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-000" + Trim(txt_codigo)
                End If
                If VAR_COD2 > 999 And VAR_COD2 < 10000 Then
                   aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-00" + Trim(txt_codigo)
                End If
                If VAR_COD2 > 9999 And VAR_COD2 < 100000 Then
                   aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-0" + Trim(txt_codigo)
                End If
                If VAR_COD2 > 99999 Then
                   aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-" + Trim(txt_codigo)
                End If
             'aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = var_cod2     'txt_codigo1.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!tipo_eqp = dtc_codigo2.Text
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!fecha_registro = Date     'no cambia
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
             aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset.Update    'Batch 'adAffectAll
        End If
      Else
        'SOLO EL REGISTRO ACTIVO
        MsgBox Str(aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_codigo)
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!pais_codigo = dtc_codigo7.Text
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!bien_codigo = IIf(IsNull(dtc_codigo21.Text) Or dtc_codigo21.Text = "", "NA1", dtc_codigo21.Text)
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_fecha = Date     'DTPfecha1.Value
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!modelo_codigo = txt_campo4.Text     '      'MODELO PROVISIONAL AUTOMATICO JQA 02-2015
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!modelo_codigo_h = "S/M"  'Txt_campo5.Text    'dtc_codigo41.Text
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!modelo_codigo_x = "S/M"   'Txt_campo6.Text    'dtc_codigo51.Text
    
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_energia = txt_campo2.Text
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_luz = txt_campo3.Text

         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!bien_cotiza_num_accesos = txt_campo7.Text
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!dimension_fosa_m = txt_campo8.Text        '
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!dimension_fosa_fondo = txt_campo9.Text
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!dimension_fosa_frente = txt_campo10.Text  '
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cuadro_ctrl_codigo = IIf((dtc_codigo61.Text = ""), 1, dtc_codigo61.Text)
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!cotiza_nro_montador = IIf((txt_campo5.Text = ""), "2", txt_campo5.Text)
         'aw_p_ao_solicitud_cotiza_venta.Ado_datosA.RECORDSET!Foto = Date
         'aw_p_ao_solicitud_cotiza_venta.Ado_datosA.RECORDSET!ARCHIVO_Foto = var_cod + ".JPG"
         'aw_p_ao_solicitud_cotiza_venta.Ado_datosA.RECORDSET!archivo_foto_cargado = "N"
         'hora_registro
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!proceso_codigo = "COM"
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!subproceso_codigo = "COM-01"
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!etapa_codigo = "COM-01-04"
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!clasif_codigo = "COM"
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!doc_codigo = "R-222"
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!doc_numero = "0"  'txt_campo1.Text
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!clasif_codigo2 = "COM"
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!doc_codigo2 = "R-224"
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!doc_numero2 = "0"  'txt_campo1.Text
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!poa_codigo = "3.1.1"
         'WWWWWWWWWWWWWW JQA 02-2015
            If VAR_COD2 < 10 Then
               aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-00000" + Trim(txt_codigo)
            End If
            If VAR_COD2 > 9 And VAR_COD2 < 100 Then
               aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-0000" + Trim(txt_codigo)
            End If
            If VAR_COD2 > 99 And VAR_COD2 < 1000 Then
               aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-000" + Trim(txt_codigo)
            End If
            If VAR_COD2 > 999 And VAR_COD2 < 10000 Then
               aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-00" + Trim(txt_codigo)
            End If
            If VAR_COD2 > 9999 And VAR_COD2 < 100000 Then
               aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-0" + Trim(txt_codigo)
            End If
            If VAR_COD2 > 99999 Then
               aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!unidad_codigo_ant = parametro + "-" + Trim(txt_codigo)
            End If
         'aw_p_ao_solicitud_cotiza_venta.Ado_datosA.RECORDSET!unidad_codigo_ant = var_cod2     'txt_codigo1.Text
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!tipo_eqp = dtc_codigo2.Text
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!fecha_registro = Date     'no cambia
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
         aw_p_ao_solicitud_cotiza_venta.Ado_datosA.Recordset.Update    'Batch 'adAffectAll
      End If
    End If
  End If
'  dtc_desc1.Visible = True
'  lbl_aux1.Visible = False
  Unload Me
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
    If (txt_campo4 = "") Then
    MsgBox "Debe registrar el Modelo del Equipo ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo7.Text = "") Then
    MsgBox "Debe registrar el Pais Origen del Equipo ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (txt_campo5.Text = "") Then
    MsgBox "Debe registrar cantidad de Montadores (Instaladores / Ajustadores) ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo2.Text = "") Then
    MsgBox "Debe registrar el Tipo de Equipo ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (txt_campo2.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If txt_campo3.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_campo7.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If txt_campo9.Text = "" Then
    MsgBox "Debe registrar: Dimención Fosa Fondo (mm) ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If txt_campo10.Text = "" Then
    MsgBox "Debe registrar: Dimención Fosa Frente (mm) ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If txt_campo7.Text = "" Then
    MsgBox "Debe registrar:Número de Accesos ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If txt_campo8.Text = "" Then
    MsgBox "Debe registrar: Espacio Libre Bajo Dintel (mm) ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo61 = "") Then
    MsgBox "Debe registrar ... " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If

'  If (dtc_codigo11 = "") Then
'    MsgBox "Debe registrar Parámetros de Cálculo, Consulte con el Administrador ... ", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (txt_fob_me1 = "") Or (txt_fob_me1 = "0") Then
'    MsgBox "Debe registrar el Precio FOB del Equipo ... ", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo21_Click(Area As Integer)
    dtc_desc21.BoundText = dtc_codigo21.BoundText
    dtc_desc24.BoundText = dtc_codigo21.BoundText
End Sub

Private Sub dtc_codigo61_Click(Area As Integer)
    dtc_desc61.BoundText = dtc_codigo61.BoundText
End Sub

Private Sub dtc_codigo7_Click(Area As Integer)
    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc21_Click(Area As Integer)
    dtc_codigo21.BoundText = dtc_desc21.BoundText
    dtc_desc24.BoundText = dtc_desc21.BoundText
End Sub

Private Sub dtc_desc24_Click(Area As Integer)
    dtc_desc21.BoundText = dtc_desc24.BoundText
    dtc_codigo21.BoundText = dtc_desc24.BoundText
End Sub

Private Sub dtc_desc61_Click(Area As Integer)
    dtc_codigo61.BoundText = dtc_desc61.BoundText
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc7.BoundText
End Sub

Private Sub Form_Activate()
    Call ABRIR_TABLA
    mbDataChanged = False
End Sub

Private Sub Form_Load()
    'Call ABRIR_TABLA
    mbDataChanged = False
'    If swnuevo = 2 Then
'        dtc_desc2.BoundText = dtc_codigo2.BoundText
'        dtc_desc3.BoundText = dtc_codigo3.BoundText
'    End If
End Sub

Private Sub ABRIR_TABLA()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from ac_costos_comercializacion ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
    'Bien (Equipo)
    Set rs_datos21 = New ADODB.Recordset
    If rs_datos21.State = 1 Then rs_datos21.Close
    rs_datos21.Open "Select * from ac_bienes where edif_codigo = '" & GlEdificio & "' OR modelo_codigo= 'NA' ", db, adOpenStatic
    Set Ado_datos21.Recordset = rs_datos21
    dtc_desc21.BoundText = dtc_codigo21.BoundText
    'gc_pais
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    If aw_p_ao_solicitud_cotiza_venta.SSTab1.Tab = 0 Then
        rs_datos7.Open "Select * from gc_pais where pais_continente = 'AMERICA' order by pais_descripcion", db, adOpenStatic
    End If
    If aw_p_ao_solicitud_cotiza_venta.SSTab1.Tab = 1 Then
        rs_datos7.Open "Select * from gc_pais where pais_continente = 'ASIA' order by pais_descripcion", db, adOpenStatic
    End If
    If aw_p_ao_solicitud_cotiza_venta.SSTab1.Tab = 2 Then
        rs_datos7.Open "Select * from gc_pais where pais_continente = 'EUROPA' order by pais_descripcion", db, adOpenStatic
    End If
'    rs_datos7.Open "Select * from gc_pais where pais_continente = '" & txt_conti & "' order by pais_descripcion", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
    'Tipo de Equipo
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from ac_bienes_equipo_tipos ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    'Cuarto de Control
    Set rs_datos61 = New ADODB.Recordset
    If rs_datos61.State = 1 Then rs_datos61.Close
    rs_datos61.Open "Select * from ac_bienes_equipo_cuadro_ctrl ", db, adOpenStatic
    Set Ado_datos61.Recordset = rs_datos61
    dtc_desc61.BoundText = dtc_codigo61.BoundText
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

'Private Sub Txt_campo4_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub


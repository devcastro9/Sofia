VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form aw_p_ao_solicitud_cotiza_det_eur 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cotizaci�n Venta - Detalle Costos"
   ClientHeight    =   5865
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10935
   ControlBox      =   0   'False
   Icon            =   "aw_p_ao_solicitud_cotiza_det_eur.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "aw_p_ao_solicitud_cotiza_det_eur.frx":0A02
      ScaleHeight     =   915
      ScaleWidth      =   10635
      TabIndex        =   11
      Top             =   120
      Width           =   10695
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Height          =   665
         Left            =   240
         Picture         =   "aw_p_ao_solicitud_cotiza_det_eur.frx":6CA34
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1350
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Height          =   665
         Left            =   1545
         MaskColor       =   &H00000000&
         Picture         =   "aw_p_ao_solicitud_cotiza_det_eur.frx":6D20A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   1475
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETALLE DE COSTOS - EUROPA"
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
         Left            =   4665
         TabIndex        =   14
         Top             =   240
         Width           =   5025
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      Height          =   4215
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   10695
      Begin VB.TextBox Txt_monto3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "costo_monto2"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "0"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Txt_monto4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "costo_monto_usd2"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5040
         TabIndex        =   36
         Text            =   "0"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Txt_monto5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "costo_monto3"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "0"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Txt_monto6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "costo_monto_usd3"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7560
         TabIndex        =   34
         Text            =   "0"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   10455
         Begin VB.TextBox Txt_campo3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "costo_porcentaje"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   720
            TabIndex        =   42
            Text            =   "0"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox Txt_monto2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            DataField       =   "costo_monto2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3360
            TabIndex        =   33
            Text            =   "0"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox Txt_monto1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "costo_monto"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "0"
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Porcentaje"
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
            Left            =   960
            TabIndex        =   41
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Monto Costo USD"
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
            Left            =   8280
            TabIndex        =   39
            Top             =   360
            Width           =   1605
         End
         Begin VB.Label txt_monto01 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "costo_monto_usd"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
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
            Left            =   8280
            TabIndex        =   38
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lbl_campo4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Monto Costo Bs."
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
            TabIndex        =   31
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Monto Costo Eur"
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
            Left            =   3360
            TabIndex        =   30
            Top             =   360
            Width           =   1470
         End
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "aw_p_ao_solicitud_cotiza_det_eur.frx":6DAF6
         DataField       =   "codigo_costo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
         Height          =   315
         Left            =   3960
         TabIndex        =   22
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo_costo"
         BoundColumn     =   "codigo_costo"
         Text            =   ""
      End
      Begin VB.TextBox Txt_campo4 
         BackColor       =   &H00C0C0C0&
         DataField       =   "costo_observaciones"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   3720
         Width           =   9980
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "aw_p_ao_solicitud_cotiza_det_eur.frx":6DB10
         DataField       =   "codigo_costo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
         Height          =   315
         Left            =   2160
         TabIndex        =   19
         Top             =   1320
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "costo_descripcion"
         BoundColumn     =   "codigo_costo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux1 
         Bindings        =   "aw_p_ao_solicitud_cotiza_det_eur.frx":6DB29
         DataField       =   "codigo_costo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
         Height          =   315
         Left            =   5160
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "costo_porcentaje"
         BoundColumn     =   "codigo_costo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux2 
         Bindings        =   "aw_p_ao_solicitud_cotiza_det_eur.frx":6DB43
         DataField       =   "codigo_costo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
         Height          =   315
         Left            =   6720
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "costo_monto"
         BoundColumn     =   "codigo_costo"
         Text            =   ""
      End
      Begin VB.Label Txt_campo8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
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
         Height          =   300
         Left            =   9840
         TabIndex        =   45
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Txt_campo7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
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
         Height          =   300
         Left            =   9000
         TabIndex        =   44
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Txt_campo6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
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
         Height          =   300
         Left            =   8160
         TabIndex        =   43
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Txt_campo5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         DataField       =   "pais_continente"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
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
         Left            =   5400
         TabIndex        =   40
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label txt_monto03 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         DataField       =   "cotiza_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.ado_detalle1"
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
         Left            =   7560
         TabIndex        =   28
         Top             =   3600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label txt_monto02 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         DataField       =   "cotiza_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.ado_detalle1"
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
         Left            =   5040
         TabIndex        =   27
         Top             =   3600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         DataField       =   "unidad_codigo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
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
         Left            =   3840
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle u Observaciones"
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
         TabIndex        =   17
         Top             =   3345
         Width           =   2220
      End
      Begin VB.Label Txt_descripcion 
         Alignment       =   2  'Center
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
         TabIndex        =   23
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N� Cotizaci�n"
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
         TabIndex        =   21
         Top             =   330
         Width           =   1200
      End
      Begin VB.Label Txt_Correl 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "cotiza_codigo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
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
         Left            =   6960
         TabIndex        =   20
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
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
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   15
         Top             =   330
         Width           =   1680
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Items para Costos"
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
         TabIndex        =   10
         Top             =   1330
         Width           =   1620
      End
      Begin VB.Label lbl_codigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N� Negociacion"
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
         Top             =   330
         Width           =   1425
      End
      Begin VB.Label Txt_campo2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36NO-"
         DataField       =   "edif_codigo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.ado_detalle1"
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
         Left            =   8760
         TabIndex        =   0
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C�digo Edificio"
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
         TabIndex        =   8
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
      ScaleWidth      =   10935
      TabIndex        =   1
      Top             =   5865
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
Attribute VB_Name = "aw_p_ao_solicitud_cotiza_det_eur"
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
On Error GoTo AddErr
   sino = MsgBox("Est� Seguro de CANCELAR la operaci�n ? ", vbYesNo + vbQuestion, "Atenci�n")
   If sino = vbYes Then
       ' frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset.CancelUpdate
        Unload Me
    End If
      Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If Txt_campo7.Caption = "" Then
        Txt_campo7.Caption = "2"
    End If
    If swnuevo = 1 Then
        frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("ges_gestion").Value = Year(Date)
        frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("unidad_codigo").Value = txt_campo1.Caption
        frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("solicitud_codigo").Value = txt_codigo.Caption
        frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("edif_codigo").Value = Txt_campo2.Caption
        frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("cotiza_codigo").Value = IIf(Txt_Correl.Caption = "0", "1", Txt_Correl.Caption)
        frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("estado_codigo").Value = "REG"
     End If
     frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("codigo_costo").Value = dtc_codigo1.Text
     frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("costo_porcentaje").Value = CDbl(Txt_campo3.Text)
     
     frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("costo_monto2").Value = Round(CDbl(txt_monto2.Text), Val(Txt_campo7.Caption))
     frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("costo_monto").Value = Round(CDbl(Txt_monto1.Text), Val(Txt_campo7.Caption))
     frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("costo_monto_usd").Value = Round(CDbl(txt_monto01), Val(Txt_campo7.Caption))
     frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("pais_continente").Value = "EUROPA" 'IIf(Txt_campo5.Caption = "", "EUROPA", Txt_campo5.Caption)
'     frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("costo_monto_usd2").Value = Round(CDbl(Txt_monto4.Text), 2)
'     aw_p_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("costo_monto3").Value = Round(CDbl(IIf(Txt_monto5.Text = "", "0", Txt_monto5.Text)), 2)
'     aw_p_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("costo_monto_usd3").Value = Round(CDbl(Txt_monto6.Text), 2)
     If swnuevo = 1 Then
        frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("costo_observaciones").Value = Trim(dtc_desc1.Text) + " - " + Trim(Txt_campo4.Text)
     Else
        frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("costo_observaciones").Value = Trim(dtc_desc1.Text)
     End If
     
     frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("fecha_registro").Value = Date
     'aw_p_ao_negociacion_cabecera.Ado_detalle1.Recordset("hora_registro").Value = Date
     frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("usr_codigo").Value = glusuario
     frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset.UpdateBatch adAffectAll
     
     Call AcumulaMonto(frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("solicitud_codigo"), frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("unidad_codigo"), frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("edif_codigo"), frm_ao_solicitud_cotiza_venta.Ado_detalle1.Recordset("cotiza_codigo"))
     'rsexiste.Open "select count(*) as numero from co_comprobante_m where cod_trans='" & Trim(codigo) & "' and org_codigo='999' and tipo_comp='ANC'", db, adOpenKeyset, adLockReadOnly
     'db.Execute "Update ao_solicitud_cotiza_venta Set cotiza_precio_total_bs = " & aw_p_ao_negociacion_cabecera.Ado_detalle1.Recordset("bitacora_codigo") & " Where unidad_codigo = '" & Txt_campo1.Caption & "' and negocia_codigo = '" & txt_codigo.Caption & "'   "
     Unload Me
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub AcumulaMonto(ges, uni, cod1, cod2)
'GRABA TOTALES E IMPUESTOS
    Set rs_aux4 = New ADODB.Recordset
    If rs_aux4.State = 1 Then rs_aux4.Close
    'rs_aux4.Open "select sum(costo_monto) as totbs, sum (costo_monto_usd) as totdl from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND cotiza_codigo = " & rs_datos!cotiza_codigo & "   ", db, adOpenKeyset, adLockOptimistic
    rs_aux4.Open "select sum(costo_monto) as totbs, sum (costo_monto_usd) as totdl from ao_solicitud_costos where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2, db, adOpenKeyset, adLockOptimistic
    If rs_aux4.RecordCount > 0 Then
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & rs_aux4!totdl & " + ao_solicitud_cotiza_venta.cotiza_fob_seg_dol   where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = ao_solicitud_cotiza_venta.cotiza_precio_total_dol * " & CDbl(Txt_campo8) & " where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
    End If
    'WWWWWWWWWWWWWWWWWW
'        Set rs_aux4 = New ADODB.Recordset
'        If rs_aux4.State = 1 Then rs_aux4.Close
'        rs_aux4.Open "select sum(costo_monto) as totbs, sum(costo_monto_usd) as totdl, sum(costo_monto2) as toteu  from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA'   ", db, adOpenKeyset, adLockOptimistic
'        If rs_aux4.RecordCount > 0 Then
'            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_me = round(" & rs_aux4!toteu & " + cotiza_precio_base_me  - cotiza_precio_flete_me, " & CDbl(cmd_dec) & ") where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = round(cotiza_precio_total_me * " & CDbl(txt_tdc_me.Text) & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = round(cotiza_precio_total_bs / " & GlTipoCambioOficial & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'        Else
'            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_me = round(" & rs_aux4!toteu & " + cotiza_precio_base_me  - cotiza_precio_flete_me, " & CDbl(cmd_dec) & ") where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = round(cotiza_precio_total_me * " & CDbl(txt_tdc_me.Text) & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = round(cotiza_precio_total_bs / " & GlTipoCambioOficial & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'        End If
'        'Importaion Cliente
'        'db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_me = " & rs_aux4!toteu & " - " & VAR_NAC_ME & " - " & VAR_ALM_ME & " - " & VAR_AGE_ME & " - " & VAR_FLE_ME & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_me = " & rs_aux4!toteu & " - " & VAR_NAC_ME & " - " & VAR_ALM_ME & " - " & VAR_AGE_ME & " - " & VAR_FLE_ME & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_bs = round(cotiza_gasto_local_me * " & CDbl(txt_tdc_me.Text) & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_dol = round(cotiza_gasto_local_bs / " & GlTipoCambioOficial & ", " & CDbl(cmd_dec) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_bs = " & CDbl(txt_local_IT_bs.Text) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_dol = cotiza_gasto_local_dol * " & CDbl(txt_local_IT_bs.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_bs = " & CDbl(txt_local_IVA_bs.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_dol = cotiza_gasto_local_dol * " & CDbl(txt_local_IVA_bs.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "   "
'
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = cotiza_precio_total_dol + cotiza_saldo_local_IT_dol + cotiza_saldo_local_IVA_dol where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = cotiza_precio_total_dol_cli * " & GlTipoCambioOficial & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "
'
'        VAR_DOLCLI = rs_aux4!totdl + frm_ao_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_cif_dol - frm_ao_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_fob_dol - frm_ao_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_seg_dol
'        VAR_BSCLI = rs_aux4!totbs + frm_ao_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_cif_bs - frm_ao_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_fob_bs - frm_ao_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_seg_bs
'        'no suma Ado_datos1A.Recordset!cotiza_precio_total_dol JQA-2015
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_totusd_menos_seguro = " & VAR_DOLCLI & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & " "
'        'VAR_SUBD = frm_ao_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_total_dol - frm_ao_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_seg_dol    'sin SEGURO
'        VAR_SUBD = frm_ao_solicitud_cotiza_venta.Ado_datosE.Recordset!cotiza_precio_total_dol         'mas SEGURO
'        VAR_SUBB = VAR_SUBD * GlTipoCambioOficial
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_bs = " & CDbl(txt_cge_IT_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_dol = (" & VAR_SUBD & " * " & CDbl(txt_cge_IT_bs) & ") where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_bs = " & CDbl(txt_cge_IVA_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = (" & VAR_SUBD & " * " & CDbl(txt_cge_IVA_bs) & ") -((cotiza_precio_base_dol * 0.1498) )-((" & CDbl(VAR_AGE) & " * 0.13))  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'        'db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = (" & VAR_SUBD & " * " & CDbl(txt_cge_IVA_bs) & ") -((cotiza_precio_cif_dol * 0.1498) * " & CDbl(dtc_desc16) & ")-((" & CDbl(VAR_AGE) & " * 0.13)* " & CDbl(dtc_desc16) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_bs = " & CDbl(IIf(txt_tac_billing_bs = "", "0", txt_tac_billing_bs)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_dol = (" & VAR_SUBD & " * " & CDbl(IIf(txt_tac_billing_bs = "", "1", txt_tac_billing_bs)) & ") where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & VAR_SUBD & "  + (ao_solicitud_cotiza_venta.cotiza_saldo_cge_IT_dol) + (ao_solicitud_cotiza_venta.cotiza_saldo_cge_IVA_dol) + (ao_solicitud_cotiza_venta.cotiza_saldo_tac_billing_dol) where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
'        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = ao_solicitud_cotiza_venta.cotiza_precio_total_dol_cge * " & GlTipoCambioOficial & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'EUROPA' AND cotiza_codigo = " & CDbl(Txt_Correl) & "  "
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2, db, adOpenKeyset, adLockOptimistic
    If rs_aux1.RecordCount > 0 Then
        'VAR_DOLCLI = rs_aux4!totdl + rs_aux1!cotiza_precio_cif_dol - rs_aux1!cotiza_precio_fob_dol - rs_aux1!cotiza_precio_seg_dol
        'VAR_BSCLI = rs_aux4!totbs + rs_aux1!cotiza_precio_total_bs_x - rs_aux1!cotiza_precio_fob_bs - rs_aux1!cotiza_precio_fob_bs_x
        
        VAR_DOLCLI = rs_aux1!cotiza_precio_total_dol - rs_aux1!cotiza_fob_seg_dol
        VAR_BSCLI = VAR_DOLCLI * CDbl(Txt_campo8)
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_totusd_menos_seguro = " & VAR_DOLCLI & " where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from ao_solicitud_costos where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & " and codigo_costo = '3' ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            VAR_NAC = rs_aux2!costo_monto_usd
        End If
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from ao_solicitud_costos where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & " and codigo_costo = '5' ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            VAR_ALM = rs_aux2!costo_monto_usd
        End If
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from ao_solicitud_costos where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & " and codigo_costo = '6'  ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            VAR_AGE = rs_aux2!costo_monto_usd
        End If
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from ao_solicitud_costos where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & " and codigo_costo = '8'  ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            VAR_FLE = IIf(IsNull(rs_aux2!costo_monto_usd), "0", rs_aux2!costo_monto_usd)
        End If
        'VAR_SUBD = VAR_DOLCLI - VAR_NAC - VAR_ALM - VAR_AGE - VAR_FLE       'rs_aux1!cotiza_precio_total_dol +
        'VAR_SUBB = VAR_SUBD * GlTipoCambioOficial
        VAR_SUBD = rs_aux4!totdl - VAR_NAC - VAR_ALM - VAR_AGE - VAR_FLE       'rs_aux1!cotiza_precio_total_dol +
        VAR_SUBB = VAR_SUBD * CDbl(Txt_campo8)
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = cotiza_precio_total_dol  + (" & VAR_SUBD & " * 0.0309) + (" & VAR_SUBD & " * 0.1491) where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = cotiza_precio_total_dol_cli * " & CDbl(Txt_campo8) & " where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = cotiza_precio_total_dol  + ((cotiza_precio_total_dol - cotiza_precio_seg_dol) * 0.0416) + ((cotiza_precio_total_dol - cotiza_precio_seg_dol) * 0.16) - ((cotiza_precio_cif_dol * 0.1498) * " & Val(frm_ao_solicitud_cotiza_venta.dtc_desc15) & " - ((" & VAR_AGE & " * 0.13)* " & Val(frm_ao_solicitud_cotiza_venta.dtc_desc15) & " ) ) + ((cotiza_precio_total_dol - cotiza_precio_seg_dol) * 0.0350) where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = cotiza_precio_total_dol_cge * " & CDbl(Txt_campo8) & " where unidad_codigo = '" & uni & "' and solicitud_codigo = " & ges & "  and edif_codigo = '" & cod1 & "' and cotiza_codigo = " & cod2 & "   "
    End If
End Sub

Private Sub valida_campos()
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar:  " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar:  " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub dtc_aux1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_aux1.BoundText
    dtc_codigo1.BoundText = dtc_aux1.BoundText
    dtc_aux2.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_aux2.BoundText
    dtc_aux1.BoundText = dtc_aux2.BoundText
    dtc_desc1.BoundText = dtc_aux2.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    dtc_aux2.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    dtc_aux2.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc1_LostFocus()
    txt_monto2.Text = dtc_aux2.Text
    Txt_campo3.Text = dtc_aux1.Text
    'WWWWWWWWWWWWWWWWWWW  JQA-2015 REVISAR CALCULOS
    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "select sum(costo_monto) as totbs, sum(costo_monto_usd) as totdl, sum(costo_monto2) as totbs2, sum(costo_monto_usd2) as totdl2, sum(costo_monto3) as totbs3, sum(costo_monto_usd3) as totdl3 from ao_solicitud_costos where ges_gestion = '" & Year(Date) & "' and unidad_codigo = '" & txt_campo1 & "' and solicitud_codigo = '" & txt_codigo & "' and edif_codigo = '" & Txt_campo2 & "' and cotiza_codigo = " & Txt_Correl & "  ", db, adOpenKeyset, adLockOptimistic
    
    Select Case dtc_codigo1.Text
        Case 1
            'SEGURO DE TRANSPORTE   0.0078
            Txt_monto1.Text = CDbl(txt_monto01) * CDbl(Txt_campo3)
            txt_monto3.Text = CDbl(txt_monto02) * CDbl(Txt_campo3)
            Txt_monto5.Text = CDbl(txt_monto03) * CDbl(Txt_campo3)
            
        Case 2
            'FLETE FRONTERA
            Txt_monto1.Text = dtc_aux2.Text
            txt_monto3.Text = dtc_aux2.Text
            Txt_monto5.Text = dtc_aux2.Text
            
        Case 3
            'NACIONALIZACION 0.1498
            'sum(costo_monto) as totbs, sum (costo_monto_usd) as totdl, sum(costo_monto2) as totbs2, sum (costo_monto_usd2) as totdl2, sum(costo_monto3) as totbs3, sum (costo_monto_usd3) as totdl3
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
                        
        Case 4
            'GAC 0.05
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 5
            'ALMACENAJE 0.007
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 6
            'COMISION AGENCIA       0.015
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 7
            'SPREAD GLOBAL  0.08
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 8
            'TOTAL FLETES
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 9
            'INSTALACION Y PINTURA
            'corregrirrrrrrrrrrrrrrrrrrrrrrrrrrrr
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 10
            'COSTOS DE OPERACION
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
'            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
'            txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 11
            'COSTOS DE INSTALACION INTERIOR
            'corregrirrrrrrrrrrrrrrrrrrrrrrrrrrrr
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 12
            'COSTOS DE AJUSTE INTERIOR
            'corregrirrrrrrrrrrrrrrrrrrrrrrrrrrrr
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 13
            'IMPREVISTOS COMISIONES
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 14
            'UTILIDAD 0.15
            Txt_monto1.Text = rs_aux1!totbs * CDbl(Txt_campo3)
            txt_monto3.Text = rs_aux1!totbs2 * CDbl(Txt_campo3)
            Txt_monto5.Text = rs_aux1!totbs3 * CDbl(Txt_campo3)
        Case 15
            'OTROS
    End Select
        If CDbl(Txt_monto1.Text) > 0 Then
            txt_monto2.Text = Round(CDbl(Txt_monto1.Text) / CDbl(Txt_campo8), 2)
        Else
            txt_monto2.Text = "0"
        End If
        
        If CDbl(txt_monto3.Text) > 0 Then
            Txt_monto4.Text = Round(CDbl(txt_monto3.Text) / CDbl(Txt_campo8), 2)
        Else
            Txt_monto4.Text = "0"
        End If
        
        If CDbl(Txt_monto5.Text) > 0 Then
            Txt_monto6.Text = Round(CDbl(Txt_monto5.Text) / CDbl(Txt_campo8), 2)
        Else
            Txt_monto6.Text = "0"
        End If
    If rs_aux1.State = 1 Then rs_aux1.Close
End Sub

Private Sub Form_Activate()
    Call ABRIR_TABLA
    mbDataChanged = False
End Sub

Private Sub Form_Load()
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
    rs_datos1.Open "Select * from ac_costos_comercializacion ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
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

Private Sub Txt_campo4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

'Private Sub Txt_monto1_Change()
'    If CDbl(txt_monto1.Text) > 0 Then
'        txt_monto2.Text = CDbl(txt_monto1.Text) / CDbl(Txt_campo8)
'    Else
'        txt_monto2.Text = "0"
'    End If
'End Sub

Private Sub txt_monto2_Change()
    If txt_monto2.Text = "" Then
        txt_monto2.Text = "0"
    End If
    If Txt_campo6.Caption = "" Then
        Txt_campo6.Caption = CDbl(GlTipoCambioEuro)
    End If
    If CDbl(txt_monto2.Text) > 0 Then
        Txt_monto1.Text = Round(CDbl(txt_monto2.Text) * CDbl(Txt_campo6), 2)        'CDbl(GlTipoCambioEuro), 2)
        txt_monto01 = Round(CDbl(txt_monto2.Text) * CDbl(IIf(Txt_campo8 = "0", "1", Txt_campo8)), 2)
    Else
        Txt_monto1.Text = "0"
        txt_monto01 = "0"
    End If
End Sub

Private Sub Txt_monto3_Change()
'    If CDbl(txt_monto3.Text) > 0 Then
'        Txt_monto4.Text = CDbl(txt_monto3.Text) / CDbl(GlTipoCambioOficial)
'    Else
'        Txt_monto4.Text = "0"
'    End If
End Sub

Private Sub Txt_monto4_Change()
    If CDbl(Txt_monto4.Text) > 0 Then
        txt_monto3.Text = CDbl(Txt_monto4.Text) * CDbl(Txt_campo8)
    Else
        txt_monto3.Text = "0"
    End If
End Sub

Private Sub Txt_monto5_Change()
    If CDbl(Txt_monto5.Text) > 0 Then
        Txt_monto6.Text = CDbl(Txt_monto5.Text) / CDbl(Txt_campo8)
    Else
        Txt_monto6.Text = "0"
    End If
End Sub

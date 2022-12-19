VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tw_tareas_crono_instalacion 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Instalaciones - Tareas Cronograma Instalacion"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   16770
   Icon            =   "tw_tareas_crono_instalacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   16770
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fechas para el Cronograma por Equipo"
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
      Height          =   2640
      Left            =   7320
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
      Width           =   8460
      Begin VB.PictureBox FraGrabarCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   8280
         TabIndex        =   52
         Top             =   1800
         Visible         =   0   'False
         Width           =   8280
         Begin VB.PictureBox BtnCancelar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4395
            Picture         =   "tw_tareas_crono_instalacion.frx":0A02
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   54
            Top             =   0
            Width           =   1455
         End
         Begin VB.PictureBox BtnGrabar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2640
            Picture         =   "tw_tareas_crono_instalacion.frx":12EE
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   53
            Top             =   0
            Width           =   1280
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
            Left            =   375
            TabIndex        =   55
            Top             =   180
            Visible         =   0   'False
            Width           =   1005
         End
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "tw_tareas_crono_instalacion.frx":1AC4
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   2100
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
         Bindings        =   "tw_tareas_crono_instalacion.frx":1ADD
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5040
         TabIndex        =   18
         Top             =   1800
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
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "fecha_ini_inst"
         DataSource      =   "Ado_detalle2"
         Height          =   315
         Left            =   2400
         TabIndex        =   60
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   118423553
         CurrentDate     =   44885
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "fecha_fin_inst"
         DataSource      =   "Ado_detalle2"
         Height          =   315
         Left            =   6480
         TabIndex        =   61
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   118423553
         CurrentDate     =   44885
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         DataField       =   "fecha_ini_ajuste"
         DataSource      =   "Ado_detalle2"
         Height          =   315
         Left            =   2400
         TabIndex        =   62
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   118423553
         CurrentDate     =   44885
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         DataField       =   "fecha_fin_ajuste"
         DataSource      =   "Ado_detalle2"
         Height          =   315
         Left            =   6480
         TabIndex        =   63
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   118423553
         CurrentDate     =   44885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio Instalacion"
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
         Top             =   480
         Width           =   2085
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin Instalacion"
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
         Left            =   4560
         TabIndex        =   58
         Top             =   480
         Width           =   1890
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio Ajuste . . ."
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
         Top             =   1200
         Width           =   1950
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin Ajuste . . ."
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
         Left            =   4560
         TabIndex        =   56
         Top             =   1200
         Width           =   1755
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable Supervisor Nal."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1860
         Width           =   2055
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Edificio en proceso de Instalación y Ajuste."
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
      Height          =   4680
      Left            =   8040
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   9300
      Begin VB.PictureBox fra_opciones2 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   30
         ScaleHeight     =   660
         ScaleWidth      =   9225
         TabIndex        =   35
         Top             =   3960
         Width           =   9225
         Begin VB.PictureBox BtnGrabarDet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3120
            Picture         =   "tw_tareas_crono_instalacion.frx":1AF6
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   37
            Top             =   0
            Width           =   1280
         End
         Begin VB.PictureBox BtnCancelarDet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4800
            Picture         =   "tw_tareas_crono_instalacion.frx":22CC
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   36
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
         Height          =   525
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3240
         Width           =   7365
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   2160
         TabIndex        =   32
         Top             =   855
         Width           =   270
      End
      Begin VB.ComboBox cmd_campo2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "zorden_cambio"
         DataSource      =   "Ado_detalle1"
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "tw_tareas_crono_instalacion.frx":2BB8
         Left            =   5520
         List            =   "tw_tareas_crono_instalacion.frx":2C8E
         TabIndex        =   6
         Text            =   "0"
         Top             =   3360
         Visible         =   0   'False
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
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "tw_tareas_crono_instalacion.frx":2DA1
         Top             =   3360
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "fmes_plan"
         DataSource      =   "Ado_detalle1"
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
         Height          =   360
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dtc_desc6 
         Bindings        =   "tw_tareas_crono_instalacion.frx":2DA3
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   3360
         TabIndex        =   2
         Top             =   1320
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo6 
         Bindings        =   "tw_tareas_crono_instalacion.frx":2DBC
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   9120
         TabIndex        =   23
         Top             =   1320
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
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "tw_tareas_crono_instalacion.frx":2DD5
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   840
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "tw_tareas_crono_instalacion.frx":2DEE
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   840
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
         Bindings        =   "tw_tareas_crono_instalacion.frx":2E07
         DataField       =   "beneficiario_codigo_rep"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   3360
         TabIndex        =   3
         Top             =   1800
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "tw_tareas_crono_instalacion.frx":2E20
         DataField       =   "beneficiario_codigo_sup"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   3360
         TabIndex        =   4
         Top             =   2280
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo7 
         Bindings        =   "tw_tareas_crono_instalacion.frx":2E39
         DataField       =   "beneficiario_codigo_rep"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   9120
         TabIndex        =   29
         Top             =   1800
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
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "tw_tareas_crono_instalacion.frx":2E52
         DataField       =   "beneficiario_codigo_sup"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   9120
         TabIndex        =   30
         Top             =   2280
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
      Begin MSDataListLib.DataCombo dtc_desc9 
         Bindings        =   "tw_tareas_crono_instalacion.frx":2E6B
         DataField       =   "beneficiario_codigo_cobr"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   3360
         TabIndex        =   39
         Top             =   2760
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo9 
         Bindings        =   "tw_tareas_crono_instalacion.frx":2E84
         DataField       =   "beneficiario_codigo_cobr"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   9120
         TabIndex        =   40
         Top             =   2760
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "#Crono (Número de Cronograma de Instalación y Ajuste)"
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
         Top             =   360
         Width           =   5025
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Ejecutivo de Ventas"
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
         Top             =   2760
         Width           =   1785
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
         Left            =   2400
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   6645
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
         TabIndex        =   33
         Top             =   3360
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
         TabIndex        =   31
         Top             =   885
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
         Left            =   4200
         TabIndex        =   28
         Top             =   3375
         Visible         =   0   'False
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
         Left            =   1680
         TabIndex        =   27
         Top             =   3375
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Label lbl_campo7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Técnico Responsable Ajuste"
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
         TabIndex        =   26
         Top             =   1815
         Width           =   2610
      End
      Begin VB.Label lbl_campo8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Supervisor Instalación y Ajuste"
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
         TabIndex        =   25
         Top             =   2280
         Width           =   2730
      End
      Begin VB.Label lbl_campo6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Tecnico Responsable Instalacion"
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
         TabIndex        =   24
         Top             =   1320
         Width           =   3015
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
      Height          =   7335
      Left            =   6240
      TabIndex        =   17
      Top             =   0
      Width           =   12885
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000018&
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
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   7440
         TabIndex        =   44
         Top             =   6915
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000018&
         Caption         =   "Pendentes (En proceso)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   3120
         TabIndex        =   43
         Top             =   6915
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.PictureBox fra_opciones_det 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   150
         ScaleHeight     =   660
         ScaleWidth      =   12585
         TabIndex        =   20
         Top             =   240
         Width           =   12585
         Begin VB.PictureBox BtnAprobar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4200
            Picture         =   "tw_tareas_crono_instalacion.frx":2E9D
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   67
            ToolTipText     =   "Aprueba el Registro Elegido"
            Top             =   0
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.PictureBox BtnEliminar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2880
            Picture         =   "tw_tareas_crono_instalacion.frx":36D0
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   66
            ToolTipText     =   "Anula Zona elegida"
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.PictureBox BtnModificar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1440
            Picture         =   "tw_tareas_crono_instalacion.frx":3E1C
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   65
            ToolTipText     =   "Modifica datos del Grupo elegido"
            Top             =   0
            Width           =   1430
         End
         Begin VB.PictureBox BtnAñadir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            Picture         =   "tw_tareas_crono_instalacion.frx":4731
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   64
            ToolTipText     =   "Crea una Nueva Zona Piloto"
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
            Left            =   6600
            Picture         =   "tw_tareas_crono_instalacion.frx":4EF0
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   46
            ToolTipText     =   "Busca Registros "
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox BtnModificar2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   8040
            Picture         =   "tw_tareas_crono_instalacion.frx":56A5
            ScaleHeight     =   615
            ScaleWidth      =   1545
            TabIndex        =   38
            Top             =   0
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.PictureBox BtnAddDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   9480
            Picture         =   "tw_tareas_crono_instalacion.frx":664E
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   9
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
            Left            =   5265
            Picture         =   "tw_tareas_crono_instalacion.frx":6E0D
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   10
            Top             =   0
            Width           =   1430
         End
         Begin VB.PictureBox BtnAnlDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   11040
            Picture         =   "tw_tareas_crono_instalacion.frx":7722
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   11
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin MSAdodcLib.Adodc Ado_detalle1 
         Height          =   330
         Left            =   120
         Top             =   6840
         Width           =   12600
         _ExtentX        =   22225
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
         Appearance      =   0
         BackColor       =   -2147483624
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
      Begin TrueOleDBGrid60.TDBGrid dg_det1 
         Bindings        =   "tw_tareas_crono_instalacion.frx":7E6E
         Height          =   5775
         Left            =   120
         OleObjectBlob   =   "tw_tareas_crono_instalacion.frx":7E89
         TabIndex        =   45
         Top             =   960
         Width           =   12615
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
      Height          =   7320
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6180
      Begin VB.PictureBox fra_opciones 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   120
         ScaleHeight     =   1020
         ScaleWidth      =   6000
         TabIndex        =   47
         Top             =   240
         Width           =   6000
         Begin VB.PictureBox BtnSalir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4560
            Picture         =   "tw_tareas_crono_instalacion.frx":DCBD
            ScaleHeight     =   615
            ScaleWidth      =   1245
            TabIndex        =   50
            ToolTipText     =   "Cierra la Ventana Activa"
            Top             =   480
            Width           =   1245
         End
         Begin VB.PictureBox BtnImprimir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "tw_tareas_crono_instalacion.frx":E47F
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   49
            ToolTipText     =   "Imprimir Todas las Zonas Piloto"
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.PictureBox BtnImprimir1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1440
            Picture         =   "tw_tareas_crono_instalacion.frx":ED4C
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   48
            ToolTipText     =   "Edificios en Cronograma vs. Contratos de Mantenimiento"
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.Label lbl_titulo 
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
            Left            =   120
            TabIndex        =   51
            Top             =   60
            Width           =   1815
         End
      End
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
         Left            =   1560
         TabIndex        =   15
         Top             =   6885
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pendentes"
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
         TabIndex        =   14
         Top             =   6885
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   6795
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
         Appearance      =   0
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
         Bindings        =   "tw_tareas_crono_instalacion.frx":F619
         Height          =   5370
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   9472
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
         Caption         =   "TIPOS DE EQUIPOS"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "tipo_eqp"
            Caption         =   "Tipo"
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
            DataField       =   "tipo_eqp_descripcion"
            Caption         =   "Tipo.Descripcion"
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
         BeginProperty Column03 
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
               Object.Visible         =   -1  'True
               ColumnWidth     =   4229.858
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column03 
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
   Begin MSAdodcLib.Adodc Ado_detalle2 
      Height          =   330
      Left            =   8040
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
Attribute VB_Name = "tw_tareas_crono_instalacion"
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
Dim rs_datos10 As New ADODB.Recordset

Dim rsNada As New ADODB.Recordset

Dim rs_det1 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset

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
Dim VAR_SW, VAR_SQL As String

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
Dim VAR_EDIF As String
Dim VAR_DA, VAR_UORIGEN, VAR_DPTO As String
                
Dim VAR_CONT As Integer

Dim mvBookMark, marca1 As Variant
Dim mbDataChanged As Boolean

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
     '<-- Inicio                Identificación del Cliente                Fin -->
     If VAR_SW <> "MOD" Then
        If Ado_datos.Recordset.RecordCount > 0 Then
'            lbl_texto0 = Ado_datos.Recordset!tipo_eqp
            BtnModificar2_Click
            dg_det1.Visible = True
            Call ABRIR_TABLA_DET
        Else
            dg_det1.Visible = False
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
    dg_det1.Visible = True
    fra_opciones.Enabled = False
    FraNavega.Enabled = False
    dg_det1.Enabled = False
'    Fra_datos.Visible = False
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
'  If Ado_datos.Recordset!estado_codigo = "REG" Then
''    Call OptFilGral1_Click
'  Else
    Call OptFilGral2_Click
    Ado_datos.Recordset.Move marca1 - 1
'  End If
'  'Call ABRIR_TABLA_DET
End Sub

Private Sub ABRIR_DET()
    'gc_edificaciones
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    If VAR_UORIGEN = "DNINS" Then
        rs_datos5.Open "Select * from gc_edificaciones where estado_codigo = 'APR' and tomadoInst = 'N' order by edif_descripcion", db, adOpenStatic
    Else
        rs_datos5.Open "Select * from gc_edificaciones where estado_codigo = 'APR' AND depto_codigo = '" & Ado_datos.Recordset!depto_codigo & "' and tomadoInst = 'N' order by edif_descripcion", db, adOpenStatic
    End If
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub BtnAnlDetalle_Click()
   If Ado_detalle1.Recordset("estado_activo") = "REG" Then
      sino = MsgBox("Está Seguro de Anular este registro ? (Este ya no será considerado en la presente Zona) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        cmd_campo2.Text = Ado_detalle1.Recordset!IdTareaInst
        
        db.Execute "update gc_edificaciones set tomadoInst= 'N' where edif_codigo = '" & dtc_codigo5.Text & "' "
        'If cmd_campo2.Text <> "0" Then
            db.Execute "update tc_tareas_crono_instalacion set zorden_cambio = IdTareaInst - 1 where IdTareaInst >= " & cmd_campo2.Text & "  and IdTareaInst = " & Ado_datos.Recordset!tipo_eqp & ""
            db.Execute "update tc_tareas_crono_instalacion set IdTareaInst = zorden_cambio  where zorden_cambio > '0'  and IdTareaInst = " & Ado_datos.Recordset!tipo_eqp & " "
            db.Execute "delete tc_tareas_crono_instalacion where correlativo = " & Text1.Text & " "
            db.Execute "update tc_tareas_crono_instalacion set zorden_cambio = '0'  where zorden_cambio > '0'"
        'End If
        'Call ABRIR_TABLA_DET
      End If
   Else
      MsgBox "No se puede ANULAR, el registro ya fue APROBADO o ya fue ANULADO anteriormente ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAñadir_Click()
  On Error GoTo EditErr
  If Ado_datos.Recordset!estado_codigo <> "ANL" Then
    swnuevo = 1
    dg_det1.Visible = True
    fra_opciones.Enabled = False
    FraNavega.Enabled = False
    dg_det1.Enabled = False
'    Fra_datos.Visible = False
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
    Call ABRIR_TABLA_DET
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
  Set rs_aux2 = New ADODB.Recordset
  If rs_aux2.State = 1 Then rs_aux2.Close
  rs_aux2.Open "select * from tv_zona_piloto_edif where IdTareaInst = '" & Ado_datos.Recordset!tipo_eqp & "' order by IdTareaInst ", db, adOpenKeyset, adLockOptimistic, adCmdText
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

Private Sub BtnCancelarDet_Click()
    swnuevo = 0
    fra_opciones.Enabled = True
    FraNavega.Enabled = True
    dg_det1.Enabled = True
'    Fra_datos.Visible = True
    FraDet2.Visible = False
    FraDet1.Visible = True
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Ado_detalle1.Recordset.CancelUpdate
    End If
    fra_opciones_det.Visible = True
    
    dtc_aux5.Visible = False
    dtc_desc5.Visible = True
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
   If ExisteReg(Ado_datos.Recordset!tipo_eqp) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
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
    If swnuevo = 1 Then
        Set rs_aux1 = New ADODB.Recordset
        If rs_aux1.State = 1 Then rs_aux1.Close
        SQL_FOR = "Select ISNULL(max(IdTareaInst),0) as Orden from tc_tareas_crono_instalacion where IdTareaInst = '" & Ado_datos.Recordset!tipo_eqp & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            Txt_campo1.Text = IIf(IsNull(rs_aux1!Orden), 1, rs_aux1!Orden + 1)
        Else
            Txt_campo1.Text = 1
        End If
        'update gc_edificaciones set tomadoInst = 'S' where edif_codigo = @edif_codigo
        'db.Execute "SELECT Txt_campo1.Text  = ISNULL(MAX(IdTareaInst),0)+1 FROM tc_tareas_crono_instalacion where IdTareaInst = '" & Ado_datos.Recordset!tipo_eqp & "' "
        db.Execute "insert into  tc_tareas_crono_instalacion(IdTareaInst, edif_codigo, IdTareaInst, zona_codigo, beneficiario_codigo, beneficiario_codigo_rep, beneficiario_codigo_cobr, zorden_cambio, observaciones, estado_codigo, fecha_registro, usr_codigo) " & _
        "values (" & Ado_datos.Recordset!IdTareaInst & ", '" & dtc_codigo5.Text & "', '" & Txt_campo1.Text & "', '0', '" & dtc_codigo6.Text & "', '" & dtc_codigo7.Text & "', '" & dtc_codigo8.Text & "', 0, '" & txt_obs.Text & "', 'REG', GETDATE(), 'ADMIN')"
        
        db.Execute "update gc_edificaciones set tomadoInst= 'S' where edif_codigo = '" & dtc_codigo5.Text & "' "
    End If
    If swnuevo = 2 Then
        db.Execute "update tc_tareas_crono_instalacion set edif_codigo= '" & dtc_codigo5.Text & "', IdTareaInst='" & Txt_campo1.Text & "', beneficiario_codigo= '" & dtc_codigo6.Text & "', beneficiario_codigo_rep= '" & dtc_codigo7.Text & "', beneficiario_codigo_sup= '" & dtc_codigo8.Text & "', beneficiario_codigo_cobr= '" & dtc_codigo8.Text & "', zorden_cambio= " & cmd_campo2.Text & ", observaciones = '" & txt_obs.Text & "', fecha_registro='" & Date & "', fecha_ini_max='" & DTPicker1.Value & "', fecha_fin_max='" & DTPicker2.Value & "'   where correlativo=" & Text1.Text & " "
        db.Execute "update tc_tareas_crono_instalacion set fecha_ini_ajuste= '" & DTPicker3.Value & "', fecha_fin_ajuste ='" & DTPicker4.Value & "'   where correlativo=" & Text1.Text & " "
        If cmd_campo2.Text <> "0" Then
            db.Execute "update tc_tareas_crono_instalacion set zorden_cambio = IdTareaInst + 1 where IdTareaInst >= " & cmd_campo2.Text & " and IdTareaInst < " & Txt_campo1.Text & " and " & Txt_campo1.Text & " > " & cmd_campo2.Text & " and IdTareaInst = " & Ado_datos.Recordset!IdTareaInst & ""
            db.Execute "update tc_tareas_crono_instalacion set zorden_cambio = IdTareaInst - 1 where IdTareaInst <= " & cmd_campo2.Text & " and IdTareaInst > " & Txt_campo1.Text & " and " & Txt_campo1.Text & " < " & cmd_campo2.Text & " and IdTareaInst = " & Ado_datos.Recordset!IdTareaInst & ""
            db.Execute "update tc_tareas_crono_instalacion set IdTareaInst = zorden_cambio  where zorden_cambio > '0'  and IdTareaInst = " & Ado_datos.Recordset!IdTareaInst & ""
            db.Execute "update tc_tareas_crono_instalacion set zorden_cambio = '0'  where zorden_cambio > '0'"
        End If
     End If
'     db.Execute "Update to_cronograma_diario Set beneficiario_codigo_resp = " & dtc_codigo4.Text & " Where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'   "
     'Call OptFilGral2_Click
     'rs_datos.MoveFirst
'     mbDataChanged = False

'    Call ABRIR_TABLA_DET
    swnuevo = 0
    fra_opciones.Enabled = True
    FraNavega.Enabled = True
    dg_det1.Enabled = True
'    Fra_datos.Visible = True
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
    CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!IdTareaInst
    
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
    CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!IdTareaInst
    iResult = CR02.PrintReport
    If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR02.WindowState = crptMaximized

End Sub

Private Sub BtnModDetalle_Click()
  If Ado_detalle1.Recordset.RecordCount = 0 Then
    MsgBox "No existen registros para Modificar, Verifique y vuelva a intentar!! ", vbExclamation
    Exit Sub
  End If
  marca1 = Ado_datos.Recordset.Bookmark
  If Ado_detalle1.Recordset.RecordCount > 0 And Ado_detalle1.Recordset!estado_codigo = "REG" Then
    swnuevo = 2
    fra_opciones.Enabled = False
    FraNavega.Enabled = False
    dg_det1.Enabled = False
'    Fra_datos.Visible = False
    FraDet2.Visible = True
    
    fra_opciones_det.Visible = False

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
    If Ado_detalle1.Recordset!estado_activo = "REG" Then
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    Else
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    End If
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnModificar_Click()
  If Ado_detalle2.Recordset.RecordCount = 0 Then
    MsgBox "No existen registros para Modificar, Verifique y vuelva a intentar!! ", vbExclamation
    Exit Sub
  End If
On Error GoTo EditErr

    FraNavega.Enabled = False
    FraDet1.Enabled = False
    
    Fra_datos.Visible = True
    
''  lblStatus.Caption = "Modificar registro"
'    If Ado_datos.Recordset!estado_codigo = "REG" Then
'        Fra_datos.Enabled = True
'        fra_opciones.Visible = False
'        'fra_opciones_det.Visible = False
'        FraGrabarCancelar.Visible = True
'        FraDet1.Visible = False
'        dg_datos.Enabled = False
'        VAR_SW = "MOD"
'    '    BtnVer.Visible = True
'    Else
'      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
'    End If
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnModificar2_Click()
    Set rs_aux4 = New ADODB.Recordset
    If rs_aux4.State = 1 Then rs_aux4.Close
    rs_aux4.Open "select * from tc_tareas_crono_instalacion where IdTareaInst = '" & Ado_datos.Recordset!tipo_eqp & "' order by IdTareaInst ", db, adOpenKeyset, adLockOptimistic, adCmdText
    If rs_aux4.RecordCount > 0 Then
        VAR_CONT = 0
        rs_aux4.MoveFirst
        While Not rs_aux4.EOF
            VAR_CONT = VAR_CONT + 1
            rs_aux4!zorden_cambio = VAR_CONT
            rs_aux4.Update
            rs_aux4.MoveNext
        Wend
        db.Execute "UPDATE tc_tareas_crono_instalacion SET IdTareaInst = zorden_cambio WHERE IdTareaInst = '" & Ado_datos.Recordset!tipo_eqp & "' "
        db.Execute "UPDATE tc_tareas_crono_instalacion SET zorden_cambio ='0' WHERE IdTareaInst = '" & Ado_datos.Recordset!tipo_eqp & "' "
'        Call ABRIR_TABLA_DET
        'MsgBox "Se recodificó la columna ORDEN, satisfactoriamente ...", vbInformation, "Información"
    Else
        MsgBox "No Existen Registros para Ordenar ...", vbExclamation, "Información"
    End If
    
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

'Private Sub dtc_codigo2_Click(Area As Integer)
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
'End Sub

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

'Private Sub dtc_desc2_Click(Area As Integer)
'    dtc_codigo2.BoundText = dtc_desc2.BoundText
'    Call pnivel2(dtc_codigo2.BoundText)
'    dtc_desc3.Enabled = True
'End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

'Private Sub pnivel2(codigo2 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from gc_municipio where prov_codigo = '" & codigo2 & "'"
'   Set dtc_codigo3.RowSource = Nothing
'   Set dtc_codigo3.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_codigo3.ReFill
'   dtc_codigo3.BoundText = Empty
'
'   Set dtc_desc3.RowSource = Nothing
'   Set dtc_desc3.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_desc3.ReFill
'   dtc_desc3.BoundText = Empty
'End Sub

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
    VAR_UORIGEN = Aux
    'HABILITAR CUANDO SE AUTORICE UTILIZAR EN LAS REGIONALES
'    If Aux = "DNINS" Then
'        Select Case VAR_DA
'            Case "1.8"    'Cochabamba
'                Aux = "DINSB"
'                VAR_DPTO = "3"
'            Case "1.7"    'Santa Cruz
'                Aux = "DINSS"
'                VAR_DPTO = "7"
'            Case "1.3", "1.2"    'La Paz - Tecnico
'                Aux = "DNINS"
'                VAR_DPTO = "2"
'            Case "1.9"    ' Chuquisaca
'                Aux = "DINSC"
'                VAR_DPTO = "1"
'            Case Else    ' TODO
'                Aux = "DNINS"
'                VAR_DPTO = "2"
'         End Select
'    End If
    parametro = Aux
    'Actualiza Edificios tomadoInsts en Organizacion de Zonas
    'db.Execute "update gc_edificaciones set tomadoInst = 'N' "
    'db.Execute "update gc_edificaciones set gc_edificaciones.tomadoInst = 'S' from gc_edificaciones inner join tc_tareas_crono_instalacion on gc_edificaciones.edif_codigo = tc_tareas_crono_instalacion.edif_codigo"
    Call ABRIR_TABLAS_AUX
    Call OptFilGral2_Click
    
'    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    

        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_departamento
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from gc_departamento order by depto_codigo ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText

    'gc_provincia
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from gc_provincia order by prov_descripcion", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
'    dtc_desc2.BoundText = dtc_codigo2.BoundText

    'gc_municipio
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_municipio where region_codigo = 'SI' order by munic_descripcion", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
'    dtc_desc3.BoundText = dtc_codigo3.BoundText

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
    
    'Beneficiario Funcionario CGI (Supervisor)
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    
    'Beneficiario Funcionario CGI (Cobrador)
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos3.Close
    'rs_datos9.Open "rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_datos9.Open "rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_desc9.BoundText = dtc_codigo9.BoundText
    
End Sub

'Private Sub dtc_codigo1_Click(Area As Integer)
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
'End Sub

'Private Sub dtc_codigo3_Click(Area As Integer)
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
'End Sub

'Private Sub dtc_desc1_Click(Area As Integer)
'    dtc_codigo1.BoundText = dtc_desc1.BoundText
'    Call pnivel1(dtc_codigo1.BoundText)
'    dtc_desc2.Enabled = True
'End Sub

'Private Sub pnivel1(codigo1 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from gc_provincia where depto_codigo = '" & codigo1 & "'"
'
'   Set dtc_codigo2.RowSource = Nothing
'   Set dtc_codigo2.RowSource = db.Execute(strConsultaF, , adCmdText)
''   Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_codigo2.ReFill
'   dtc_codigo2.BoundText = Empty
'
'   Set dtc_desc2.RowSource = Nothing
'   Set dtc_desc2.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_desc2.ReFill
'   dtc_desc2.BoundText = Empty
'End Sub

'Private Sub dtc_desc3_Click(Area As Integer)
'    dtc_codigo3.BoundText = dtc_desc3.BoundText
'    'Call pnivel5(dtc_codigo3.BoundText)
'    'dtc_desc9.Enabled = True
'End Sub
   
'Private Sub pnivel5(codigo7 As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_zonas where munic_codigo = '" & codigo7 & "' order by zona_denominacion"
'   Set dtc_codigo9.RowSource = Nothing
'   Set dtc_codigo9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_codigo9.ReFill
'   dtc_codigo9.BoundText = Empty
'
'   Set dtc_desc9.RowSource = Nothing
'   Set dtc_desc9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
'   dtc_desc9.ReFill
'   dtc_desc9.BoundText = Empty
'End Sub

Private Sub OptFilGral1_Click()
    '===== Proceso para filtrado general de datos (todos los registros)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    If VAR_UORIGEN = "DNINS" Then
        queryinicial = "Select * from tc_zonas_piloto_inst WHERE IdTareaInst <> '0' "
    Else
'        Select Case VAR_DPTO
'           Case "1"    ' Chuquisaca
'               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "' OR depto_codigo = '5') "
'           Case "2"    'La Paz - Tecnico
'               If glusuario = "ADMIN" Or glusuario = "OCOLODRO" Or glusuario = "JSAAVEDRA" Or glusuario = "CSALINAS" Or glusuario = "JAVIER" Then
'                    queryinicial = "Select * from tc_zonas_piloto  "
'               Else
'                    queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
'               End If
'           Case "3"    'Cochabamba
'               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
'           Case "7"    'Santa Cruz
'               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "' OR depto_codigo = '1' OR depto_codigo = '8') "
'           Case "4"    'Oruro - Tecnico
'               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
'           Case "5"    ' Potosi
'               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
'           Case "6"    ' Tarija
'               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
'           Case "8"    ' Beni
'               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
'           Case "9"    ' Pando
'               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
'           Case Else    ' TODO
'               queryinicial = "select * From tc_zonas_piloto  "     'tv_cronograma_edificaciones
'        End Select
    End If
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    '===== Proceso para filtrado general de datos (todos los registros)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
        queryinicial = "Select * from ac_bienes_equipo_tipos WHERE tipo_eqp <> 'X' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If glPersNew = "P" Then
  End If
  glPersNew = "N"
   
'   If (rstbeneficiario.State = adStateClosed) Then rstbeneficiario.Close
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub

Private Function ExisteReg(Codigo As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'estado_codigo = 'APR' and
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM to_cronograma_diario_final_INST WHERE edif_codigo = '" & Codigo & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub ABRIR_TABLA_DET()
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    rs_det1.Open "select * from tc_tareas_crono_instalacion where tipo_eqp = '" & Ado_datos.Recordset!tipo_eqp & "' order by IdTareaInst ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    Set dg_det1.DataSource = Ado_detalle1.Recordset
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        dg_det1.Visible = True
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
    End If
End Sub


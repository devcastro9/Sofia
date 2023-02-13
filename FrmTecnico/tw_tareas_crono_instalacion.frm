VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
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
   ClientWidth     =   18375
   Icon            =   "tw_tareas_crono_instalacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   18375
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
      Left            =   6600
      TabIndex        =   8
      Top             =   7440
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
         TabIndex        =   30
         Top             =   1800
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   33
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
         TabIndex        =   10
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
         TabIndex        =   38
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   111476737
         CurrentDate     =   44885
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "fecha_fin_inst"
         DataSource      =   "Ado_detalle2"
         Height          =   315
         Left            =   6480
         TabIndex        =   39
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   111476737
         CurrentDate     =   44885
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         DataField       =   "fecha_ini_ajuste"
         DataSource      =   "Ado_detalle2"
         Height          =   315
         Left            =   2400
         TabIndex        =   40
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   111476737
         CurrentDate     =   44885
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         DataField       =   "fecha_fin_ajuste"
         DataSource      =   "Ado_detalle2"
         Height          =   315
         Left            =   6480
         TabIndex        =   41
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   111476737
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   11
         Top             =   1860
         Width           =   2055
      End
   End
   Begin VB.Frame fraAgregarTarea 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Agragar tarea"
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
      Height          =   3240
      Left            =   7440
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   10620
      Begin VB.CheckBox chkHabilitar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cambiar parametro base DEFINITIVAMENTE"
         Height          =   195
         Left            =   240
         TabIndex        =   57
         Top             =   2160
         Width           =   4575
      End
      Begin VB.TextBox txtPeriodosModif 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   56
         Text            =   "0"
         Top             =   1680
         Width           =   495
      End
      Begin VB.ComboBox cmbUnidadMedida 
         Height          =   315
         ItemData        =   "tw_tareas_crono_instalacion.frx":1AF6
         Left            =   8280
         List            =   "tw_tareas_crono_instalacion.frx":1B00
         TabIndex        =   54
         Text            =   "DNINS"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtUnidadMedida 
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "Medio Dia"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtNroPeriodos 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   50
         Text            =   "0"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   375
         Left            =   240
         TabIndex        =   48
         Top             =   1200
         Width           =   10095
      End
      Begin VB.TextBox txtIdTarea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox fra_opciones2 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   30
         ScaleHeight     =   660
         ScaleWidth      =   10545
         TabIndex        =   16
         Top             =   2520
         Width           =   10545
         Begin VB.PictureBox btnGrabarTarea 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3600
            Picture         =   "tw_tareas_crono_instalacion.frx":1B12
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   18
            Top             =   0
            Width           =   1280
         End
         Begin VB.PictureBox btnCancelarTarea 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5280
            Picture         =   "tw_tareas_crono_instalacion.frx":22E8
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   17
            Top             =   0
            Width           =   1400
         End
      End
      Begin VB.TextBox txtTipoEquipo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "A"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblNroPeriodosModif 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Periodos Modif:"
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
         Left            =   3720
         TabIndex        =   55
         Top             =   1680
         Width           =   1830
      End
      Begin VB.Label lblUnidadCodigo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
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
         TabIndex        =   53
         Top             =   1680
         Width           =   705
      End
      Begin VB.Label lblUnidadMedida 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Medida:"
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
         Left            =   4080
         TabIndex        =   51
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label lblNroPeriodos 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. de Periodos:"
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
         TabIndex        =   49
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblIdTarea 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo de tarea:"
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
         Left            =   8040
         TabIndex        =   46
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label lblTipoEquipo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo del Equipo:"
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
         TabIndex        =   20
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion de la tarea:"
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
         TabIndex        =   15
         Top             =   885
         Width           =   2370
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO DE TAREAS POR TIPO DE EQUIPO"
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
      TabIndex        =   9
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   12
         Top             =   240
         Width           =   12585
         Begin VB.PictureBox BtnAprobar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4200
            Picture         =   "tw_tareas_crono_instalacion.frx":2BD4
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   45
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
            Picture         =   "tw_tareas_crono_instalacion.frx":3407
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   44
            ToolTipText     =   "Anula Zona elegida"
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox BtnModificar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1440
            Picture         =   "tw_tareas_crono_instalacion.frx":3B53
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   43
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
            Picture         =   "tw_tareas_crono_instalacion.frx":4468
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   42
            ToolTipText     =   "Crea una Nueva Zona Piloto"
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox BtnBuscar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6600
            Picture         =   "tw_tareas_crono_instalacion.frx":4C27
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   24
            ToolTipText     =   "Busca Registros "
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.PictureBox BtnModificar2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   8040
            Picture         =   "tw_tareas_crono_instalacion.frx":53DC
            ScaleHeight     =   615
            ScaleWidth      =   1545
            TabIndex        =   19
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
            Picture         =   "tw_tareas_crono_instalacion.frx":6385
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   1
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
            Picture         =   "tw_tareas_crono_instalacion.frx":6B44
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   1430
         End
         Begin VB.PictureBox BtnAnlDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   11040
            Picture         =   "tw_tareas_crono_instalacion.frx":7459
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   3
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
         Bindings        =   "tw_tareas_crono_instalacion.frx":7BA5
         Height          =   5775
         Left            =   120
         OleObjectBlob   =   "tw_tareas_crono_instalacion.frx":7BC0
         TabIndex        =   23
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
      TabIndex        =   4
      Top             =   0
      Width           =   6180
      Begin VB.PictureBox fra_opciones 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   120
         ScaleHeight     =   1020
         ScaleWidth      =   6000
         TabIndex        =   25
         Top             =   240
         Width           =   6000
         Begin VB.PictureBox BtnSalir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4560
            Picture         =   "tw_tareas_crono_instalacion.frx":E004
            ScaleHeight     =   615
            ScaleWidth      =   1245
            TabIndex        =   28
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
            Picture         =   "tw_tareas_crono_instalacion.frx":E7C6
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   27
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
            Picture         =   "tw_tareas_crono_instalacion.frx":F093
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   26
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
            TabIndex        =   29
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         Bindings        =   "tw_tareas_crono_instalacion.frx":F960
         Height          =   5370
         Left            =   120
         TabIndex        =   5
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
               ColumnWidth     =   585,071
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   -1  'True
               ColumnWidth     =   4229,858
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   599,811
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
Dim nuevo As Boolean
Dim nroPeriodosAux As Integer

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
    'If Not Ado_datos.BOF And Not Ado_datos.EOF Then
        If VAR_SW <> "MOD" Then
            If Ado_datos.Recordset.RecordCount > 0 Then
                dg_det1.Visible = True
                Call ABRIR_TABLA_DET
            Else
                dg_det1.Visible = False
            End If
        Else
            Set dg_det1.DataSource = rsNada
        End If
    'End If
End Sub

Private Sub BtnAñadir_Click()
  If Ado_datos.Recordset!estado_codigo <> "ANL" Then
        nuevo = True
        FraNavega.Enabled = False
        FraDet1.Enabled = False
        cargarDatos (1)
        fraAgregarTarea.Caption = "Aregar Tarea"
        
        lblNroPeriodosModif.Visible = False
        txtPeriodosModif.Visible = False
        chkHabilitar.Visible = False
        txtNroPeriodos.Locked = False
        txtNroPeriodos.backColor = &H80000005
        txtPeriodosModif.Text = 0
        
        fraAgregarTarea.Visible = True
        fraAgregarTarea.Enabled = True
  Else
    MsgBox "No se puede Agregar una nueva tarea", vbExclamation
  End If
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

Private Sub btnCancelarTarea_Click()
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    fraAgregarTarea.Visible = False
    fraAgregarTarea.Enabled = False
End Sub

Private Sub btnEliminar_Click()
On Error GoTo UpdateErr
    If Ado_datos.Recordset.RecordCount < 1 Or Ado_detalle1.Recordset.RecordCount < 1 Then
        MsgBox "No existen Tareas para anular, seleccione una  y vuelva a intentar", vbExclamation, "ANULAR"
    Else
        If MsgBox("Seguro que quiere ANULAR la tarea " & Ado_detalle1.Recordset!IdTareaInst & _
                        " del tipo de equipo " & Ado_datos.Recordset!tipo_eqp & "?", vbQuestion + vbYesNo, "ANULAR") = vbYes Then
                        
            db.Execute "UPDATE tc_tareas_crono_instalacion SET estado_codigo = 'ANL' WHERE tipo_eqp = '" & Ado_datos.Recordset!tipo_eqp & _
            "' AND IdTareaInst = " & Ado_detalle1.Recordset!IdTareaInst
            
            MsgBox "Se anuló la tarea " & Ado_detalle1.Recordset!IdTareaInst & " del tipo de equipo " & Ado_datos.Recordset!tipo_eqp, vbOKOnly, "TAREA ELIMINADA"
        End If
        ABRIR_TABLA_DET
    End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub btnGrabarTarea_Click()
    If campoValido Then
        Exit Sub
    End If
    Dim nroDias As Double
    If nuevo Then
        If MsgBox("Quiere agregar esta nueva tarea?", vbQuestion + vbYesNo, "AGREGAR") = vbYes Then
            nroDias = CInt(txtNroPeriodos) / 2
            
            db.Execute "INSERT INTO tc_tareas_crono_instalacion VALUES('" & Ado_datos.Recordset!tipo_eqp & _
            "', (SELECT COUNT(*) + 1 AS idTareaInst FROM tc_tareas_crono_instalacion WHERE tipo_eqp = '" & Ado_datos.Recordset!tipo_eqp & _
            "'), '" & txtDescripcion.Text & "', REPLACE('" & nroDias & "',',','.'),  " & CInt(txtNroPeriodos.Text) & ", 'MDIA', " & CInt(txtNroPeriodos.Text) * 4 & _
            ", " & CInt(txtNroPeriodos.Text) & ", '" & cmbUnidadMedida.Text & "', 'REG', 'REG', 'REG', '" & ObtenerFechaServidor & "', '" & glusuario & "')"
            
            MsgBox "Se Agrego una nueva tarea", vbOKOnly, "GUARDADO CON EXITO"
            btnCancelarTarea_Click
        End If
    Else
        If MsgBox("Seguro que quiere modificar la tarea " & Ado_detalle1.Recordset!IdTareaInst & _
                        " del tipo de equipo " & Ado_datos.Recordset!tipo_eqp & "?", vbQuestion + vbYesNo, "MODIFICAR") = vbYes Then
            nroDias = CInt(txtNroPeriodos) / 2
            
            db.Execute "UPDATE tc_tareas_crono_instalacion SET TareaDescripcion = '" & txtDescripcion.Text & "', NroEstimadoDias = REPLACE('" & nroDias & _
            "',',','.'), NroEstimadoHoras = " & CInt(txtNroPeriodos.Text) * 4 & ", NroPeriodosModif = " & CInt(txtPeriodosModif.Text) & _
            ", unidad_codigo = '" & cmbUnidadMedida.Text & "', usr_codigo = '" & glusuario & "', NroTiempoPeriodos = " & CInt(txtNroPeriodos) & _
            " WHERE tipo_eqp = '" & Ado_datos.Recordset!tipo_eqp & "' AND IdTareaInst = " & CInt(txtIdTarea.Text)
            
            MsgBox "Se modificó la tarea " & Ado_detalle1.Recordset!IdTareaInst & _
                        " del tipo de equipo " & Ado_datos.Recordset!tipo_eqp & "", vbOKOnly, "GUARDADO CON EXITO"
            btnCancelarTarea_Click
        End If
    End If
    ABRIR_TABLA_DET
End Sub

Private Function campoValido() As Boolean 'Si algun campo no es valido devuelve TRUE como alarma
    If Len(txtDescripcion.Text) < 10 Then
        campoValido = True
        MsgBox "La descripcion es demasiado corta", vbExclamation, "CAMPOS INCORRECTOS"
        Exit Function
    End If
    If Len(txtDescripcion.Text) > 100 Then
        campoValido = True
        MsgBox "La descripcion es demasiado larga", vbExclamation, "CAMPOS INCORRECTOS"
        Exit Function
    End If
    If Not esNumero(txtNroPeriodos) Then
        campoValido = True
        MsgBox "El número de periodos debe ser UN NUMERO!!!", vbExclamation, "CAMPOS INCORRECTOS"
        Exit Function
    End If
    If Not esNumero(txtPeriodosModif) Then
        campoValido = True
        MsgBox "El número de periodos modificados debe ser UN NUMERO!!!", vbExclamation, "CAMPOS INCORRECTOS"
        Exit Function
    End If
    If (cmbUnidadMedida.Text <> "DNINS") And (cmbUnidadMedida.Text <> "DNAJS") Then
        campoValido = True
        MsgBox "La unidad debe ser comprendida entre DNINS y DNAJS", vbExclamation, "CAMPOS INCORRECTOS"
        Exit Function
    End If
End Function

Private Function esNumero(txtAux As TextBox) As Boolean
    On Error GoTo EditErr
    Dim entero As Integer
    entero = CInt(txtAux.Text)
    esNumero = True
    Exit Function
EditErr:
    esNumero = False
End Function

Private Sub cargarDatos(tipo As Integer)
    borrarCampos
    If tipo = 1 Then
        txtTipoEquipo.Text = Ado_datos.Recordset!tipo_eqp
    Else
        txtTipoEquipo.Text = Ado_datos.Recordset!tipo_eqp
        txtIdTarea.Text = Ado_detalle1.Recordset!IdTareaInst
        txtDescripcion.Text = Ado_detalle1.Recordset!TareaDescripcion
        txtNroPeriodos.Text = Ado_detalle1.Recordset!NroTiempoPeriodos
        cmbUnidadMedida.Text = Ado_detalle1.Recordset!unidad_codigo
    End If
End Sub

Private Sub borrarCampos()
    txtTipoEquipo.Text = ""
    txtIdTarea.Text = "0"
    txtDescripcion.Text = ""
    txtNroPeriodos = 0
    txtUnidadMedida = "Medio Dia"
    cmbUnidadMedida.Text = "DNINS"
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

Private Sub BtnModificar_Click()
  If Ado_detalle1.Recordset.RecordCount = 0 Then
    MsgBox "No existen registros para Modificar, seleccione uno y vuelva a intentar!! ", vbExclamation
    Exit Sub
  End If
On Error GoTo EditErr
    nuevo = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    cargarDatos (0)
    nroPeriodosAux = CInt(txtNroPeriodos.Text)
    fraAgregarTarea.Caption = "Modificar Tarea"
    
    lblNroPeriodosModif.Visible = True
    txtPeriodosModif.Visible = True
    chkHabilitar.Visible = True
    chkHabilitar.Value = False
    txtNroPeriodos.Locked = True
    txtNroPeriodos.backColor = &H80000000
    txtPeriodosModif.Text = Ado_detalle1.Recordset!NroPeriodosModif
    
    fraAgregarTarea.Visible = True
    fraAgregarTarea.Enabled = True
    Exit Sub
EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub chkHabilitar_Click()
    If chkHabilitar.Value Then
        txtNroPeriodos.Locked = False
        txtNroPeriodos.backColor = &H80000005
    Else
        txtNroPeriodos.Locked = True
        txtNroPeriodos.backColor = &H80000000
        txtNroPeriodos.Text = nroPeriodosAux
    End If
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub Form_Load()
    'swnuevo = 0
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
    'Call ABRIR_TABLAS_AUX
    Call OptFilGral2_Click
    
'    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    

        Call SeguridadSet(Me)
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
    Else
        dg_det1.Visible = False
    End If
End Sub

Private Sub rbHabilitar_Click()
    'If rbHabilitar.Value Then
     '   txtNroPeriodos.Locked = False
    'Else
     '   txtNroPeriodos.Locked = True
    'End If
End Sub

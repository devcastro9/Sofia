VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form aw_p_ao_solicitud 
   BackColor       =   &H00000000&
   Caption         =   "Oportunidades de Negocio - Solicitud de Cotización"
   ClientHeight    =   10260
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "aw_p_ao_solicitud.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10260
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1640
      Left            =   120
      Picture         =   "aw_p_ao_solicitud.frx":0A02
      ScaleHeight     =   1575
      ScaleWidth      =   1875
      TabIndex        =   90
      Top             =   8000
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   600
         Picture         =   "aw_p_ao_solicitud.frx":6AE20
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Anula Producto Elegido"
         Top             =   825
         Width           =   765
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   945
         Picture         =   "aw_p_ao_solicitud.frx":6B262
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Modifica Producto Elegido"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   120
         Picture         =   "aw_p_ao_solicitud.frx":6B6A4
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Adiciona Producto"
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1640
      Left            =   120
      Picture         =   "aw_p_ao_solicitud.frx":6BAE6
      ScaleHeight     =   1575
      ScaleWidth      =   1875
      TabIndex        =   86
      Top             =   6240
      Width           =   1935
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   120
         Picture         =   "aw_p_ao_solicitud.frx":D7B18
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Adiciona Detalle"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   945
         Picture         =   "aw_p_ao_solicitud.frx":D7F5A
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAnlDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   600
         Picture         =   "aw_p_ao_solicitud.frx":D839C
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Elimina Detalle Elegido"
         Top             =   840
         Width           =   765
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00000000&
      Caption         =   "Ejecución del POA"
      ForeColor       =   &H00FFFFC0&
      Height          =   1695
      Left            =   2200
      TabIndex        =   60
      Top             =   7900
      Visible         =   0   'False
      Width           =   12855
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "aw_p_ao_solicitud.frx":D87DE
         Height          =   1335
         Left            =   195
         TabIndex        =   61
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   2355
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "CodGrupo"
            Caption         =   "Grupo"
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
            DataField       =   "cod_montador"
            Caption         =   "Sub-Grupo"
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
            DataField       =   "CodDetalle"
            Caption         =   "Codigo Producto"
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
            DataField       =   "DescDetalle"
            Caption         =   "Denominación del Producto"
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
            DataField       =   "cantidad"
            Caption         =   "Cantidad"
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
         BeginProperty Column05 
            DataField       =   "precio_compra"
            Caption         =   "Precio.Actual"
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
            DataField       =   "Total_compra"
            Caption         =   "Total Actual"
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
            DataField       =   "precio_venta"
            Caption         =   "Precio.c/Dscto."
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
            DataField       =   "Total_venta"
            Caption         =   "Total c/Dscto."
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
         BeginProperty Column09 
            DataField       =   "profesion"
            Caption         =   "Caracteristicas del Bien"
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
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00000000&
      Caption         =   "DETALLE DE LA EDIFICACION"
      ForeColor       =   &H00FFFFC0&
      Height          =   1695
      Left            =   2200
      TabIndex        =   54
      Top             =   6180
      Width           =   12855
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "aw_p_ao_solicitud.frx":D87F9
         Height          =   1335
         Left            =   200
         TabIndex        =   55
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   2355
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
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "edif_codigo"
            Caption         =   "Codigo Edificio"
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
            DataField       =   "edif_area_total_m2"
            Caption         =   "Area Total mt2"
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
            DataField       =   "edif_area_util_m2"
            Caption         =   "Area Util mt2"
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
            DataField       =   "edif_num_pisos"
            Caption         =   "Nro.de Pisos"
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
            DataField       =   "edif_num_habit_libres"
            Caption         =   "Nro.Habit.Libres"
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
         BeginProperty Column05 
            DataField       =   "edif_num_habit_ocupadas"
            Caption         =   "Nro.Habit.Ocupadas"
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
            DataField       =   "edif_num_habit_dorm_2"
            Caption         =   "Habit.de 2 Dorm."
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
            DataField       =   "edif_num_habit_dorm_3"
            Caption         =   "Habit.de 3 Dorm."
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
            DataField       =   "edif_num_habit_dorm_4"
            Caption         =   "Habit.>= de 4 Dorm."
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
         BeginProperty Column09 
            DataField       =   "edif_num_salas_may_200m"
            Caption         =   "Nro.Salas >200 mt."
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
            DataField       =   "edif_num_salas_men_200m"
            Caption         =   "Nro.Salas < 200 mt"
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
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "aw_p_ao_solicitud.frx":D8814
      ScaleHeight     =   960
      ScaleWidth      =   14880
      TabIndex        =   37
      Top             =   60
      Width           =   14940
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "aw_p_ao_solicitud.frx":144846
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "aw_p_ao_solicitud.frx":144A50
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "aw_p_ao_solicitud.frx":145074
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "aw_p_ao_solicitud.frx":145654
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   6000
         Picture         =   "aw_p_ao_solicitud.frx":14631E
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4320
         Picture         =   "aw_p_ao_solicitud.frx":146528
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "aw_p_ao_solicitud.frx":146AE5
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "aw_p_ao_solicitud.frx":14709D
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   720
         Left            =   5160
         Picture         =   "aw_p_ao_solicitud.frx":1472A7
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOLICITUD DE COTIZACIÓN"
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
         Left            =   8640
         TabIndex        =   47
         Top             =   300
         Width           =   4185
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "aw_p_ao_solicitud.frx":1476E9
      ScaleHeight     =   915
      ScaleWidth      =   14880
      TabIndex        =   33
      Top             =   60
      Width           =   14940
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "aw_p_ao_solicitud.frx":1B371B
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "aw_p_ao_solicitud.frx":1B3925
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOLICITUD DE COTIZACIÓN"
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
         Left            =   8460
         TabIndex        =   36
         Top             =   300
         Width           =   4185
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "GERENCIA GENERAL"
      ForeColor       =   &H00FFFFC0&
      Height          =   4920
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   5895
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3B2F
         Height          =   4170
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   7355
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "solicitud_codigo"
            Caption         =   "Cód.Neg."
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
            Caption         =   "Cod.Proy."
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
            DataField       =   "solicitud_fecha_solicitud"
            Caption         =   "Fecha.Solic."
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pendientes"
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
         Left            =   1320
         TabIndex        =   79
         Top             =   4520
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFC0&
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
         Left            =   3600
         TabIndex        =   80
         Top             =   4520
         Width           =   915
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   4440
         Width           =   5625
         _ExtentX        =   9922
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
         BackColor       =   16777152
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
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00000000&
      Height          =   4920
      Left            =   6105
      TabIndex        =   15
      Top             =   1200
      Width           =   8895
      Begin VB.TextBox Text7 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   4340
         TabIndex        =   85
         Top             =   495
         Width           =   375
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5740
         TabIndex        =   84
         Top             =   3270
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   8390
         TabIndex        =   83
         Top             =   2655
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   8390
         TabIndex        =   78
         Top             =   1115
         Width           =   375
      End
      Begin VB.TextBox txt_obs 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_observaciones"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   69
         Top             =   1800
         Visible         =   0   'False
         Width           =   1125
      End
      Begin MSDataListLib.DataCombo dtc_aux11 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3B47
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6360
         TabIndex        =   68
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo11 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3B62
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7440
         TabIndex        =   67
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
      Begin VB.TextBox Text5 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   3780
         TabIndex        =   66
         Top             =   3915
         Width           =   375
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   3780
         TabIndex        =   65
         Top             =   3270
         Width           =   375
      End
      Begin VB.TextBox txt_aux9 
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   4380
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Top             =   3660
         Width           =   4365
      End
      Begin MSDataListLib.DataCombo dtc_aux1 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3B7C
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2520
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo3"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   3780
         TabIndex        =   63
         Top             =   2655
         Width           =   375
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3B96
         DataField       =   "solicitud_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3720
         TabIndex        =   50
         Top             =   120
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "solicitud_fecha_solicitud"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   6585
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   84541441
         CurrentDate     =   41678
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3BAF
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3060
         TabIndex        =   7
         Top             =   4440
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3BC9
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo5"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3BE2
         DataField       =   "clasif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7440
         TabIndex        =   31
         Top             =   2520
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "clasif_codigo"
         BoundColumn     =   "clasif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3BFB
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7200
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo7 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3C14
         DataField       =   "etapa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2280
         TabIndex        =   29
         Top             =   3840
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "etapa_codigo"
         BoundColumn     =   "etapa_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo6 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3C2D
         DataField       =   "subproceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2520
         TabIndex        =   28
         Top             =   3120
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "subproceso_codigo"
         BoundColumn     =   "subproceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3C46
         DataField       =   "proceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2880
         TabIndex        =   27
         Top             =   2520
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "proceso_codigo"
         BoundColumn     =   "proceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3C5F
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3120
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox Txt_descripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_justificacion"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1920
         Width           =   7185
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3C78
         DataField       =   "proceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   2640
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "proceso_descripcion"
         BoundColumn     =   "proceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3C91
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   1100
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc7 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3CAA
         DataField       =   "etapa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   3900
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "etapa_descripcion"
         BoundColumn     =   "etapa_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc6 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3CC3
         DataField       =   "subproceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   21
         Top             =   3255
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "subproceso_descripcion"
         BoundColumn     =   "subproceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3CDC
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4380
         TabIndex        =   25
         Top             =   1100
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3CF5
         DataField       =   "clasif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4380
         TabIndex        =   5
         Top             =   2640
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "clasif_descripcion"
         BoundColumn     =   "clasif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3D0E
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3480
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3D28
         DataField       =   "solicitud_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3240
         TabIndex        =   49
         Top             =   120
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   -2147483629
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
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
      Begin MSDataListLib.DataCombo dtc_codigo9 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3D41
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4380
         TabIndex        =   6
         Top             =   3255
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "doc_codigo"
         BoundColumn     =   "doc_codigo"
         Text            =   "Todos"
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc9 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3D5A
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4380
         TabIndex        =   51
         Top             =   3720
         Visible         =   0   'False
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ListField       =   "doc_descripcion"
         BoundColumn     =   "doc_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3D73
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1455
         TabIndex        =   0
         Top             =   480
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3D8C
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   57
         Top             =   4440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc11 
         Bindings        =   "aw_p_ao_solicitud.frx":1B3DA6
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4380
         TabIndex        =   3
         Top             =   1520
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cite Negociación"
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
         Index           =   6
         Left            =   4875
         TabIndex        =   82
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label Txt_campo2 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "36NO"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4800
         TabIndex        =   81
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Proyecto de Edificación"
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
         Left            =   180
         TabIndex        =   77
         Top             =   840
         Width           =   2130
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Justificación"
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
         Left            =   180
         TabIndex        =   76
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lbl_campo10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Actividad del POA"
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
         Left            =   180
         TabIndex        =   75
         Top             =   4440
         Width           =   1635
      End
      Begin VB.Label lbl_campo9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Código de Registro"
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
         Left            =   4380
         TabIndex        =   74
         Top             =   3000
         Width           =   1755
      End
      Begin VB.Label lbl_campo8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Clasificación de Documentos"
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
         Left            =   4380
         TabIndex        =   73
         Top             =   2400
         Width           =   2610
      End
      Begin VB.Label lbl_campo11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Personal de CGI que Recibió la Solicitud Cotiz."
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
         Left            =   180
         TabIndex        =   72
         Top             =   1545
         Width           =   4155
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cliente que Solicita la Cotización"
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
         Left            =   4380
         TabIndex        =   71
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label lbl_campo1 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1480
         TabIndex        =   70
         Top             =   220
         Width           =   1560
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         X1              =   4260
         X2              =   4260
         Y1              =   2340
         Y2              =   4340
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   8880
         Y1              =   4335
         Y2              =   4335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   8880
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   180
         TabIndex        =   59
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   7140
         TabIndex        =   58
         Top             =   3255
         Width           =   1600
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro. Documento"
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
         Index           =   13
         Left            =   7140
         TabIndex        =   53
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha Solicitud"
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
         Index           =   12
         Left            =   6585
         TabIndex        =   52
         Top             =   225
         Width           =   1380
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Etapa de Proceso"
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
         Index           =   5
         Left            =   180
         TabIndex        =   24
         Top             =   3645
         Width           =   1635
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Sub Proceso"
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
         Index           =   4
         Left            =   180
         TabIndex        =   23
         Top             =   3015
         Width           =   1170
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Proceso"
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
         Index           =   3
         Left            =   180
         TabIndex        =   22
         Top             =   2400
         Width           =   765
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   8115
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cod.Negocia"
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
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   2
         Left            =   8115
         TabIndex        =   16
         Top             =   225
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
      ScaleWidth      =   15120
      TabIndex        =   9
      Top             =   10260
      Width           =   15120
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   14
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   9480
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
   Begin Crystal.CrystalReport CR01 
      Left            =   6240
      Top             =   9840
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2160
      Top             =   9480
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
      Left            =   4200
      Top             =   9480
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
      Left            =   6240
      Top             =   9480
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
      Left            =   8280
      Top             =   9480
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
      Left            =   10320
      Top             =   9480
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
      Left            =   12360
      Top             =   9480
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
      Top             =   9840
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
      Left            =   2160
      Top             =   9840
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
      Left            =   4200
      Top             =   9840
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
      Left            =   6240
      Top             =   9840
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
      Left            =   8280
      Top             =   9840
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
      Left            =   10320
      Top             =   9840
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
End
Attribute VB_Name = "aw_p_ao_solicitud"
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

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim queryinicial As String

Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim CodBien As String
Dim VAR_UNI As String
Dim sino As String
Dim parametro As String
Dim VAR_AUX, VAR_CONT2 As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAddDetalle_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos!estado_codigo = "REG" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False
    Select Case dtc_codigo2.Text
        Case "1"
        Case "2"
        Case "3"
            Call ABRIR_TABLA_DET3
            Ado_detalle1.Recordset.AddNew
            aw_p_ao_solicitud_edificacion.txt_codigo.Caption = Me.txt_codigo.Caption
            aw_p_ao_solicitud_edificacion.txt_campo1.Caption = Me.dtc_codigo1.Text
            aw_p_ao_solicitud_edificacion.Txt_descripcion.Caption = Me.dtc_desc1.Text
            'aw_p_ao_solicitud_edificacion.Txt_Correl.Caption = 0
            aw_p_ao_solicitud_edificacion.dtc_codigo1.Text = Me.dtc_codigo3.Text
            aw_p_ao_solicitud_edificacion.dtc_desc1.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
            aw_p_ao_solicitud_edificacion.dtc_aux1.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
            aw_p_ao_solicitud_edificacion.Dtc_aux2.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
            aw_p_ao_solicitud_edificacion.dtc_aux3.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
            aw_p_ao_solicitud_edificacion.Txt_estado.Caption = "REG"
            aw_p_ao_solicitud_edificacion.Show vbModal
        Case "4"
    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnAnlDetalle_Click()
   sino = MsgBox("Está Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_detalle1.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_detalle1.Recordset.Delete 'adAffectAll
'        Ado_detalle1.Recordset("estado_codigo") = "ERR"
'        Ado_detalle1.Recordset("fecha_registro") = Date
'        Ado_detalle1.Recordset("usr_codigo") = GlUsuario
'        Ado_detalle1.Recordset("campo1") = "REG. ANULADO"
'        Ado_detalle1.Recordset.Update  'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If Ado_datos.Recordset!beneficiario_codigo = "0" Or Ado_datos.Recordset!beneficiario_codigo = "" Then
        MsgBox "No se puede APROBAR, debe registrar al Propietario del Proyecto de Edificación: " + lbl_campo4.Caption, vbExclamation, "Validación de Registro"
        Exit Sub
   End If
   Set rs_aux2 = New ADODB.Recordset
   rs_aux2.Open "Select * from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
   If rs_aux2.RecordCount > 0 Then
        VAR_CONT2 = rs_aux2.RecordCount
   End If
   'If rs_datos!estado_codigo = "REG" And Ado_datos.Recordset!correl_edificacion > 0 Then
   If rs_datos!estado_codigo = "REG" And VAR_CONT2 > 0 Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        Select Case dtc_codigo2.Text
            Case "1"
            Case "2"
            Case "3"
                Set rs_aux1 = New ADODB.Recordset
                'SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_detalle1.Recordset!edif_codigo & "'  "
                SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   "
                rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                'If rs_aux1.RecordCount > 0 Then
                '    MsgBox "El código ya existe, consulte con el administrador del Sistema..."
                '    var_cod = 0
                '    Exit Sub
                'Else
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    'rs_aux2.Open "Select max(trafico_codigo) as Codigo from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    rs_aux2.Open "Select max(trafico_codigo) as Codigo from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    End If
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    rs_aux2.Open "Select edif_capacidad_min_trafico as Codigo from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        VAR_AUX = rs_aux2!Codigo
                    End If
                    rs_aux1.AddNew
                    'var_cod = rs_aux1.RecordCount + 1
                    rs_aux1!ges_gestion = Year(Date)
                    rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
                    rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
                    rs_aux1!edif_codigo = Ado_detalle1.Recordset!edif_codigo
                    rs_aux1!trafico_codigo = var_cod
                    rs_aux1!trafico_h_capacidad_trafico_parametro = Round(VAR_AUX, 2)
                    rs_aux1!estado_codigo = "REG"
                    rs_aux1!fecha_registro = Date
                    rs_aux1!usr_codigo = GlUsuario
                    rs_aux1.Update
                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
                'End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
            Case "4"
        End Select
        Set rs_aux2 = New ADODB.Recordset
        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9.Text & "'  "
        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
            txt_campo1.Caption = rs_aux2!correl_doc
            rs_aux2.Update
        End If
        rs_datos!doc_numero = txt_campo1.Caption
        'REVISAR !!! JQA 2014_07_08
        'VAR_ARCH = RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
        VAR_ARCH = "COM_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
        rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
        rs_datos!archivo_respaldo_cargado = "N"
        rs_datos!estado_codigo = "APR"
        rs_datos!fecha_registro = Date
        rs_datos!usr_codigo = GlUsuario
        rs_datos.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene DETALLE ...", vbExclamation, "Validación de Registro"
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
        Fra_datos.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        'txt_codigo.Enabled = True
        VAR_SW = ""
        dtc_codigo9.Enabled = True
    End If
'    dtc_desc1.Visible = True
'    lbl_aux1.Visible = False
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
   If ExisteReg(Ado_datos.Recordset!edif_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
   If rs_datos!estado_codigo = "APR" Then
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "ERR"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = GlUsuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
      MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
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
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = GlUsuario
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
        'SQL_FOR = "select * from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & var_cod & "  "
        SQL_FOR = "select * from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            var_cod = rs_aux1.RecordCount + 1
            'MsgBox "El código ya existe, consulte con el administrador del Sistema..."
            'var_cod = 0
            'Exit Sub
        Else
            'var_cod = rs_datos.RecordCount '+ 1
            var_cod = 0
        End If
        'var_cod = RTrim(RTrim(dtc_codigo2.Text) + "-") + LTrim(Str(Val(dtc_aux2) + 1))
        txt_codigo.Caption = var_cod
        rs_datos!solicitud_codigo = var_cod
        rs_datos!estado_codigo = "REG"      'no cambia
        rs_datos!ges_gestion = Year(Date)   'no cambia
        rs_datos!unidad_codigo = VAR_UNI
        'Actualiza correaltivo ...
        db.Execute "Update gc_unidad_ejecutora Set correl_solicitud = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "'   "
        rs_datos!doc_numero = "0"    'txt_campo1.Caption
        'rs_datos!correl_edificacion = 0
        rs_datos!archivo_respaldo = "sin_nombre"
        rs_datos!archivo_respaldo_cargado = "N"
     End If
     rs_datos!solicitud_fecha_solicitud = DTPfecha1.Value
     rs_datos!solicitud_tipo = dtc_codigo2.Text
     rs_datos!edif_codigo = dtc_codigo3.Text
     If dtc_codigo4.Text = "" Or dtc_codigo4.Text = "0" Then
        rs_datos!beneficiario_codigo = dtc_aux3.Text
     Else
        rs_datos!beneficiario_codigo = dtc_codigo4.Text
     End If
     rs_datos!solicitud_justificacion = Txt_descripcion.Text
     
     Select Case dtc_codigo2.Text
        Case "1"    'Solo Compras
        Case "2"    'Solo Ventas
        Case "3"    'Compras y Ventas de Equipos
        Case "4"    ' Ventas Nuevas
            rs_datos!proceso_codigo = IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = IIf(dtc_codigo6.Text = "", "COM-01", dtc_codigo6.Text)
            rs_datos!etapa_codigo = IIf(dtc_codigo7.Text = "", "COM-01-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = IIf(dtc_codigo8.Text = "", "COM", dtc_codigo8.Text)
            rs_datos!doc_codigo = IIf(dtc_codigo9.Text = "", "R-234", dtc_codigo9.Text)
        Case "5"    'Modernizacion
     End Select
     rs_datos!poa_codigo = dtc_codigo10.Text
     rs_datos!solicitud_observaciones = txt_obs.Text
     rs_datos!solicitud_fecha_recepción = DTPfecha1.Value
     rs_datos!beneficiario_codigo_resp = dtc_codigo11.Text
     
     rs_datos!ges_gestion_ant = Year(Date)
     rs_datos!unidad_codigo_ant = "-"
     rs_datos!solicitud_codigo_ant = 0
     rs_datos!usr_codigo_aprueba = ""
     rs_datos!fecha_aprueba = Date
     rs_datos!hora_aprueba = ""
     'rs_datos!Foto = Date
     'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
     'rs_datos!archivo_foto_cargado = "N"
     'hora_registro
     rs_datos!fecha_registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(GlUsuario = "", "ADMIN", GlUsuario) 'no cambia
     rs_datos.Update    'Batch 'adAffectAll
    
     Call ABRIR_TABLA
     rs_datos.MoveLast
     mbDataChanged = False
      
     Fra_datos.Enabled = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
'     dtc_desc1.BackColor = &HFFFFC0
     VAR_SW = ""
     dtc_codigo9.Enabled = True
      
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
  If (dtc_codigo8.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo9.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
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
        CR01.ReportFileName = App.Path & "\Reportes\comercial\ar_solicitud_cotizacion.rpt"
        CR01.WindowShowPrintSetupBtn = True
        CR01.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          'CR01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
          'CR01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
        'Call CREAVISTAF11          'JQA JUN-2008
        CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR01.PrintReport
        If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atención"
    End If
Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnModDetalle_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
    swnuevo = 2
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False

    Select Case dtc_codigo2.Text
        Case "1"
        Case "2"
        Case "3"
            Call ABRIR_TABLA_DET3
            aw_p_ao_solicitud_edificacion.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
            aw_p_ao_solicitud_edificacion.txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
            aw_p_ao_solicitud_edificacion.Txt_descripcion.Caption = Me.dtc_desc1.Text
            'aw_p_ao_solicitud_edificacion.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
            'aw_p_ao_solicitud_edificacion.Txt_estado.Caption = "REG"
            aw_p_ao_solicitud_edificacion.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("edif_codigo")
            aw_p_ao_solicitud_edificacion.dtc_desc1.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
            aw_p_ao_solicitud_edificacion.dtc_aux1.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
            aw_p_ao_solicitud_edificacion.Dtc_aux2.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
            aw_p_ao_solicitud_edificacion.dtc_aux3.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
            
            aw_p_ao_solicitud_edificacion.Txt_campo2.Text = Me.Ado_detalle1.Recordset("edif_area_total_m2")
            aw_p_ao_solicitud_edificacion.Txt_campo3.Text = Me.Ado_detalle1.Recordset("edif_area_util_m2")
            aw_p_ao_solicitud_edificacion.Txt_campo4.Text = Me.Ado_detalle1.Recordset("edif_num_pisos")
            aw_p_ao_solicitud_edificacion.Txt_campo5.Text = Me.Ado_detalle1.Recordset("edif_num_salas_may_200m")
            aw_p_ao_solicitud_edificacion.Txt_campo6.Text = Me.Ado_detalle1.Recordset("edif_num_salas_men_200m")
            aw_p_ao_solicitud_edificacion.Txt_campo7.Text = Me.Ado_detalle1.Recordset("edif_num_habit_libres")
            aw_p_ao_solicitud_edificacion.Txt_campo8.Text = Me.Ado_detalle1.Recordset("edif_num_habit_ocupadas")
            aw_p_ao_solicitud_edificacion.Txt_campo9.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_2")
            aw_p_ao_solicitud_edificacion.Txt_campo10.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_3")
            aw_p_ao_solicitud_edificacion.txt_campo11.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_4")
            aw_p_ao_solicitud_edificacion.txt_campo12.Caption = Me.Ado_detalle1.Recordset("edif_indicador_min_trafico")
            aw_p_ao_solicitud_edificacion.Txt_campo13.Caption = Me.Ado_detalle1.Recordset("edif_capacidad_min_trafico")
        
            aw_p_ao_solicitud_edificacion.Show vbModal
        Case "4"
           
    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If

End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        Fra_datos.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        VAR_SW = "MOD"
    '    dtc_desc1.Visible = False
    '    lbl_aux1.Visible = True
    '    lbl_aux1.Caption = dtc_desc1.Text
        dtc_desc4.SetFocus
    '    BtnVer.Visible = True
        dtc_codigo9.Enabled = False
    Else
      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
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
    dtc_desc9.BoundText = dtc_codigo9.BoundText
End Sub

Private Sub dtc_codigo9_LostFocus()
'  If VAR_SW = "ADD" Then
'    Set rs_aux2 = New ADODB.Recordset
'    SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9.Text & "'  "
'    rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'    If rs_aux2.RecordCount > 0 Then
'        rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'        txt_campo1.Caption = rs_aux2!correl_doc
'        rs_aux2.Update
'    End If
'  End If
  txt_aux9.Text = dtc_desc9.Text
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    Call pnivel1(dtc_codigo1.BoundText)
    dtc_desc10.Enabled = True
    Call pnivel11(dtc_codigo1.BoundText)
    dtc_desc11.Enabled = True
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
  
Private Sub pnivel11(codigo1 As String)
   Dim strConsultaF As String
   'strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
   strConsultaF = "Select * from gv_personal_contratado where unidad_codigo = '" & codigo1 & "' order by beneficiario_denominacion"
   
   Set dtc_codigo11.RowSource = Nothing
   Set dtc_codigo11.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_codigo11.ReFill
   dtc_codigo11.BoundText = Empty
   
   Set dtc_desc11.RowSource = Nothing
   Set dtc_desc11.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_desc11.ReFill
   dtc_desc11.BoundText = Empty
End Sub

Private Sub dtc_desc1_LostFocus()
'    dtc_codigo5.Text = dtc_aux1.Text
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
    Call pnivel5(dtc_codigo5.BoundText)
    dtc_desc6.Enabled = True
End Sub

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
    Txt_descripcion.Text = "SOLICITUD DE COTIZACION - " + dtc_desc3.Text
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
'    Call pnivel5(dtc_codigo5.BoundText)
'    dtc_desc6.Enabled = True
End Sub
   
Private Sub pnivel5(codigo5 As String)
   'Dim strConsultaF As String
   'strConsultaF = "select * from gc_proceso_nivel2 where proceso_codigo = '" & codigo5 & "'"
   
   Set dtc_codigo6.RowSource = Nothing
   'Set dtc_codigo6.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_codigo6.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel2 '" & codigo5 & "' ")
   dtc_codigo6.ReFill
   dtc_codigo6.BoundText = Empty
   
   Set dtc_desc6.RowSource = Nothing
   'Set dtc_desc6.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_desc6.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel2 '" & codigo5 & "' ")
   dtc_desc6.ReFill
   dtc_desc6.BoundText = Empty
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
'    Call pnivel6(dtc_codigo6.BoundText)
'    dtc_desc7.Enabled = True
End Sub
  
Private Sub pnivel6(codigo6 As String)
   Dim strConsultaF As String
   strConsultaF = "select * from gc_proceso_nivel3 where subproceso_codigo = '" & codigo6 & "'"
   
   Set dtc_codigo7.RowSource = Nothing
   Set dtc_codigo7.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo7.RowSource = db.Execute("EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
   dtc_codigo7.ReFill
   dtc_codigo7.BoundText = Empty
   
   Set dtc_desc7.RowSource = Nothing
   Set dtc_desc7.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo7.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
   dtc_desc7.ReFill
   dtc_desc7.BoundText = Empty
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc7.BoundText
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
    Call pnivel8(dtc_codigo8.BoundText)
    'dtc_desc9.Enabled = True
    dtc_codigo9.Enabled = True
End Sub
   
Private Sub pnivel8(codigo8 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_documentos_respaldo where clasif_codigo = '" & codigo8 & "'"
   
   Set dtc_codigo9.RowSource = Nothing
   Set dtc_codigo9.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo9.ReFill
   dtc_codigo9.BoundText = Empty
   
   Set dtc_desc9.RowSource = Nothing
   Set dtc_desc9.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc9.ReFill
   dtc_desc9.BoundText = Empty
End Sub

Private Sub dtc_desc9_Click(Area As Integer)
    dtc_codigo9.BoundText = dtc_codigo9.BoundText
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    'parametro = "estado_codigo" + " = " + "'REG'"
    '
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    'Call ABRIR_TABLA
    'Call ABRIR_TABLA_AUX2
    'Call ABRIR_TABLA_DET3
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    db.Execute (" EXEC gp_actualiza_beneficiario_edif ")
'    lbl_aux1.Visible = False
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    'rs_datos2.Open "Select * from gc_tipo_solicitud order by solicitud_tipo", db, adOpenStatic
    rs_datos2.Open "gp_listar_apr_gc_tipo_solicitud", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from gc_edificaciones order by edif_denominacion", db, adOpenStatic
    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "Select * from gc_beneficiario where (tipoben_codigo < 20 and tipoben_codigo <> 1) order by beneficiario_denominacion", db, adOpenStatic
    rs_datos4.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    'rs_datos5.Open "Select * from gc_proceso_nivel1 order by proceso_descripcion", db, adOpenStatic
    rs_datos5.Open "gp_listar_apr_gc_proceso_nivel1", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    'rs_datos6.Open "Select * from gc_proceso_nivel2 order by subproceso_descripcion", db, adOpenStatic
    rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    'rs_datos7.Open "Select * from gc_proceso_nivel3 order by etapa_descripcion", db, adOpenStatic
    rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
          
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    'rs_datos8.Open "Select * from gc_documentos_clasificacion order by clasif_codigo", db, adOpenStatic
    rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    'rs_datos9.Open "Select * from gc_documentos_respaldo order by doc_codigo", db, adOpenStatic
    rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_desc9.BoundText = dtc_codigo9.BoundText
    
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    'rs_datos10.Open "Select * from pc_poa_actividad order by poa_codigo", db, adOpenStatic
    rs_datos10.Open "pp_listar_apr_pc_poa_actividad", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText
End Sub

Private Sub ABRIR_TABLA()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "select solicitud_codigo, unidad_codigo, solicitud_justificacion, solicitud_observaciones, estado_codigo, fecha_registro, usr_codigo, hora_registro, ges_gestion, solicitud_fecha_solicitud as fecha1,  solicitud_fecha_recepción as fecha2, solicitud_tipo as codigo2, beneficiario_codigo as codigo4, beneficiario_codigo_resp as codigo11, edif_codigo as codigo3, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero As campo1, poa_codigo As codigo10, archivo_respaldo, archivo_respaldo_cargado, ges_gestion_ant, unidad_codigo_ant, solicitud_codigo_ant, usr_codigo_aprueba, fecha_aprueba, hora_aprueba From ao_solicitud WHERE estado_codigo = 'REG' "
    queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub ABRIR_TABLA_DET3()
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    rs_det1.Open "select * from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
    'rs_det1.Open "select * from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    'rs_det1.Open "SELECT ges_gestion, unidad_codigo, solicitud_codigo, estado_codigo, fecha_registro, hora_registro, usr_codigo, edif_codigo , edif_area_total_m2, edif_area_util_m2, edif_num_salas_may_200m, edif_num_salas_men_200m, edif_num_pisos, edif_num_habit_libres, edif_num_habit_ocupadas, edif_num_habit_dorm_2, edif_num_habit_dorm_3, edif_num_habit_dorm_4, edif_indicador_min_trafico, edif_capacidad_min_trafico, archivo_edificio, archivo_edificio_cargado, archivo_plano, archivo_plano_cargado From ao_solicitud_edificacion WHERE unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " ", db, adOpenKeyset, adLockOptimistic, adCmdText
    'rs_det1.Open "pp_listar_apr_pc_poa_actividad", db, adOpenStatic
    Set Ado_detalle1.Recordset = rs_det1
    Set dg_det1.DataSource = Ado_detalle1.Recordset
End Sub

Private Sub ABRIR_TABLA_AUX2()
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    'rs_datos11.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
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
    'Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
    ' <-- Inicio                Identificación del Cliente                Fin -->   'esto es de Caption
    'Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
    'Image2 = Img_Foto
'    If Ado_datos.Recordset!archivo_foto_cargado = "S" Then
'        'BtnVer.Visible = True
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
'        Image2 = Img_Foto
'    Else
'        'BtnVer.Visible = False
'        'chkEstado.Value = vbUnchecked
'    End If
    If VAR_SW <> "ADD" Then
        Select Case dtc_codigo2.Text
            Case "1"
            Case "2"
            Case "3"
                Call ABRIR_TABLA_DET3
            Case "4"
           
        End Select
        Call ABRIR_TABLA_AUX2
    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det1.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
    txt_aux9.Text = dtc_desc9.Text
    If Ado_datos.Recordset!estado_codigo = "APR" Then
            FrmABMDet.Visible = False
    Else
            FrmABMDet.Visible = True
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

Private Sub BtnAñadir_Click()
  On Error GoTo AddErr
    VAR_SW = "ADD"
    'lblStatus.Caption = "Agregar registro"
    Fra_datos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    'txt_codigo.Enabled = False
    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    rs_datos.AddNew
    dtc_desc1.SetFocus
    'dtc_desc1.BackColor = &H80000005
    dtc_codigo2.Text = 3
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    dtc_codigo5.Text = "COM"
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_codigo6.Text = "COM-01"
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    dtc_codigo7.Text = "COM-01-02"
    dtc_desc7.BoundText = dtc_codigo7.BoundText
'    BtnVer.Visible = False
    dtc_codigo9.Enabled = False
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
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE dgral_codigo = '" & Unidad & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub OptFilGral1_Click()
    parametro = "estado_codigo" + " = " + "'REG'"
    Call ABRIR_TABLA
End Sub

Private Sub OptFilGral2_Click()
    parametro = "estado_codigo" + " <> " + "'0'"
    Call ABRIR_TABLA
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_obs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

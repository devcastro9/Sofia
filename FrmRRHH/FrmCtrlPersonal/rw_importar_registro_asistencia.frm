VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rw_importar_registro_asistencia 
   BackColor       =   &H00000000&
   Caption         =   "Reportes RRHH"
   ClientHeight    =   8790
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   14835
   Icon            =   "rw_importar_registro_asistencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   14835
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Parámetros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2595
      Left            =   600
      TabIndex        =   40
      Top             =   2520
      Visible         =   0   'False
      Width           =   9540
      Begin VB.TextBox txt_mes 
         BackColor       =   &H00000000&
         DataField       =   "mes_grupo"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "0"
         Top             =   1320
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.CommandButton CmdElim2 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3000
         Picture         =   "rw_importar_registro_asistencia.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Anula Registro Activo"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdApr2 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5160
         Picture         =   "rw_importar_registro_asistencia.frx":1404
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Aprueba Registro Activo"
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cbo_mes_rep 
         Height          =   315
         ItemData        =   "rw_importar_registro_asistencia.frx":198E
         Left            =   5280
         List            =   "rw_importar_registro_asistencia.frx":19B6
         TabIndex        =   45
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox cmb_gestion 
         Height          =   315
         ItemData        =   "rw_importar_registro_asistencia.frx":1A1F
         Left            =   1920
         List            =   "rw_importar_registro_asistencia.frx":1A44
         TabIndex        =   44
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "TODAS LAS PLANILLAS"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   6480
         TabIndex        =   43
         Top             =   1920
         Width           =   2115
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "TODAS INTERIOR"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   6480
         TabIndex        =   42
         Top             =   2280
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.ComboBox cb_aguinaldo 
         Height          =   315
         ItemData        =   "rw_importar_registro_asistencia.frx":1A8A
         Left            =   5280
         List            =   "rw_importar_registro_asistencia.frx":1A94
         TabIndex        =   41
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo dtc_rep_det 
         Bindings        =   "rw_importar_registro_asistencia.frx":1AB2
         DataField       =   "planilla_codigo"
         Height          =   315
         Left            =   2880
         TabIndex        =   46
         Top             =   1920
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "planilla_descripcion"
         BoundColumn     =   "planilla_codigo"
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
      Begin MSDataListLib.DataCombo dtc_rep_cod 
         Bindings        =   "rw_importar_registro_asistencia.frx":1ACE
         DataField       =   "planilla_codigo"
         Height          =   315
         Left            =   1920
         TabIndex        =   47
         Top             =   1920
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "planilla_codigo"
         BoundColumn     =   "planilla_codigo"
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
      Begin MSDataListLib.DataCombo dtc_depto 
         Bindings        =   "rw_importar_registro_asistencia.frx":1AEA
         DataField       =   "planilla_codigo"
         Height          =   315
         Left            =   1920
         TabIndex        =   48
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "depto_codigo"
         BoundColumn     =   "planilla_codigo"
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
      Begin MSAdodcLib.Adodc Ado_datos_rep 
         Height          =   330
         Left            =   0
         Top             =   0
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
         Caption         =   "Ado_cuenta"
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
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "ACEPTAR"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   5220
         TabIndex        =   55
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "CANCELAR"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   3000
         TabIndex        =   54
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label32 
         BackColor       =   &H00000000&
         Caption         =   "GESTIÓN"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   960
         TabIndex        =   51
         Top             =   1335
         Width           =   735
      End
      Begin VB.Label Label33 
         BackColor       =   &H00000000&
         Caption         =   "MES"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   4800
         TabIndex        =   50
         Top             =   1335
         Width           =   735
      End
      Begin VB.Label Label34 
         BackColor       =   &H00000000&
         Caption         =   "PLANILLA"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   960
         TabIndex        =   49
         Top             =   1935
         Width           =   855
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   14
      Top             =   0
      Width           =   20280
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000010&
         Caption         =   "ELIMINAR ASISTENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8640
         Picture         =   "rw_importar_registro_asistencia.frx":1B06
         TabIndex        =   39
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Picture         =   "rw_importar_registro_asistencia.frx":2508
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10800
         Picture         =   "rw_importar_registro_asistencia.frx":2712
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Picture         =   "rw_importar_registro_asistencia.frx":2B54
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   21
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
         Picture         =   "rw_importar_registro_asistencia.frx":3313
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   1430
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "rw_importar_registro_asistencia.frx":3C28
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6960
         Picture         =   "rw_importar_registro_asistencia.frx":4374
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4080
         Picture         =   "rw_importar_registro_asistencia.frx":4BA7
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5520
         Picture         =   "rw_importar_registro_asistencia.frx":535C
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "rw_importar_registro_asistencia.frx":5C29
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   15
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
         TabIndex        =   24
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
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4275
         Picture         =   "rw_importar_registro_asistencia.frx":63EB
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   12
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2880
         Picture         =   "rw_importar_registro_asistencia.frx":6CD7
         ScaleHeight     =   615
         ScaleWidth      =   1305
         TabIndex        =   11
         Top             =   0
         Width           =   1300
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
         Left            =   14175
         TabIndex        =   13
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFC0&
      Height          =   7320
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   8415
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   6840
         Width           =   8145
         _ExtentX        =   14367
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
         BackColor       =   16777215
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
         Caption         =   " <-- Inicio                                                  Asistencia                                              Fin -->"
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
         Bindings        =   "rw_importar_registro_asistencia.frx":74AD
         Height          =   6495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   11456
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Fecha_control"
            Caption         =   "Fecha"
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "CI"
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
            DataField       =   "Nombre"
            Caption         =   "Nombre"
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
            DataField       =   "HoraTres"
            Caption         =   "Marca Ing."
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
            DataField       =   "HoraCuatro"
            Caption         =   "Marca Sal."
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
            DataField       =   "Tardanza"
            Caption         =   "Tardanza"
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
            DataField       =   "siFalta"
            Caption         =   "Falta"
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
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   -1  'True
               ColumnWidth     =   645.165
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fra_ABM 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7320
      Left            =   8640
      TabIndex        =   6
      Top             =   720
      Width           =   5805
      Begin VB.ComboBox cmb_gestion_rep 
         Height          =   315
         ItemData        =   "rw_importar_registro_asistencia.frx":74C5
         Left            =   2400
         List            =   "rw_importar_registro_asistencia.frx":74EA
         TabIndex        =   38
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox cmb_mes_ini 
         DataField       =   "mes_inicio_crono"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "rw_importar_registro_asistencia.frx":7530
         Left            =   360
         List            =   "rw_importar_registro_asistencia.frx":7558
         TabIndex        =   37
         Top             =   1200
         Width           =   2100
      End
      Begin VB.TextBox LblMensaje 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   1320
         TabIndex        =   36
         Text            =   "IMPORTANDO DATOS ..."
         Top             =   6360
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.ComboBox cmb_equipo 
         Height          =   315
         ItemData        =   "rw_importar_registro_asistencia.frx":75C1
         Left            =   3240
         List            =   "rw_importar_registro_asistencia.frx":75C3
         TabIndex        =   31
         Top             =   2280
         Width           =   2175
      End
      Begin VB.ComboBox cmb_departamento 
         Height          =   315
         Left            =   360
         TabIndex        =   30
         Top             =   2280
         Width           =   2055
      End
      Begin VB.OptionButton rbtMes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Es por mes"
         Height          =   375
         Left            =   3720
         TabIndex        =   29
         Top             =   1320
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton rbtDia 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Es por día"
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   28
         Top             =   1080
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   285
         Left            =   360
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         _Version        =   393216
         Format          =   119668737
         CurrentDate     =   42570
      End
      Begin VB.CommandButton btnImportarDato 
         BackColor       =   &H80000010&
         Caption         =   "Importar Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2400
         Picture         =   "rw_importar_registro_asistencia.frx":75C5
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton btnCargarArchivo 
         BackColor       =   &H80000010&
         Caption         =   "Cargar Archivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2400
         Picture         =   "rw_importar_registro_asistencia.frx":7FC7
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   1440
         Picture         =   "rw_importar_registro_asistencia.frx":8409
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1440
         Picture         =   "rw_importar_registro_asistencia.frx":8713
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lbl_inicialq 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Equipo Biométrico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   3360
         TabIndex        =   34
         Top             =   1920
         Width           =   1905
      End
      Begin VB.Label lbl_inicialw 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lugar (Departamento)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   33
         Top             =   1920
         Width           =   2280
      End
      Begin VB.Label lbl_inicialr 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Elija ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   3720
         TabIndex        =   32
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lbl_inicial 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Elija Dia a Procesar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   2280
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
      ScaleWidth      =   14835
      TabIndex        =   0
      Top             =   8790
      Width           =   14835
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
      Left            =   2400
      Top             =   6480
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
      Top             =   6480
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   6840
      TabIndex        =   35
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "rw_importar_registro_asistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NombreArchivo As String
Dim SiEstaImportado As Boolean
Dim Mensaje As String
Dim Fecha As Date
Dim rs_aux7 As New ADODB.Recordset

Private Sub cbo_mes_rep_Change()
 txt_mes.Text = cbo_mes_rep.ListIndex
 txt_mes.Text = Val(txt_mes.Text) + 1
End Sub

Private Sub cmb_mes_ini_Click()
txt_mes.Text = cmb_mes_ini.ListIndex
txt_mes.Text = Val(txt_mes.Text) + 1
End Sub

Private Sub CmdApr2_Click()

sino = MsgBox("¿Está Seguro de Eliminar el Registro de asistencia?", vbYesNo + vbQuestion, "Atención")
If sino = vbYes Then
db.Execute "rp_borrar_asistencia '" & cmb_gestion.Text & "', '" & txt_mes.Text & "', '" & dtc_rep_cod.Text & "'"
MsgBox "Se elimino la asistencia"
Frame1.Visible = False

End If

End Sub

Private Sub CmdElim2_Click()
Frame1.Visible = False
End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
End Sub

Private Sub dtc_rep_cod_Click(Area As Integer)
dtc_rep_det.BoundText = dtc_rep_cod.BoundText
dtc_rep_det.BoundText = dtc_depto.BoundText
Option1.Value = False
End Sub

Private Sub dtc_rep_det_Click(Area As Integer)
dtc_rep_cod.BoundText = dtc_rep_det.BoundText
dtc_depto.BoundText = dtc_rep_det.BoundText
Option1.Value = False
End Sub

Private Sub Form_Load()
    Call CargarControles
    NombreArchivo = ""
    SiEstaImportado = False
    Call limpiar
      
    Set rs_aux7 = New ADODB.Recordset
    If rs_aux7.State = 1 Then rs_aux7.Close
    rs_aux7.Open "SELECT * FROM rc_planilla_grupo", db, adOpenStatic
    Set Ado_datos_rep.Recordset = rs_aux7
    dtc_rep_det.BoundText = dtc_rep_cod.BoundText
        Call SeguridadSet(Me)
End Sub

Private Sub BtnAñadir_Click()
    Call limpiar
    lblMensaje.Visible = False
    Fra_ABM.Enabled = True
    btnImportarDato.Visible = False
    Image1.Visible = True
    cmb_gestion_rep.Text = Year(Date)
End Sub

Private Sub limpiar()
Mensaje = ""
    SiEstaImportado = False
    
    btnCargarArchivo.Enabled = True
    btnImportarDato.Enabled = False
    cmb_departamento = ""
    cmb_equipo = ""
    DtpFecha.Value = Date
    
End Sub

Private Sub btnCargarArchivo_Click()

    Dim rutaArchivo As String
    rutaArchivo = App.Path & "\ASISTENCIA\"
    lblMensaje.Visible = False
    Dim existeRuta As Boolean
    Dim oDir As New Scripting.FileSystemObject
    existeRuta = oDir.FolderExists(rutaArchivo)
    
    ' Valida si existe ruta destino.
    If existeRuta = Falso Then
      ' Consulta no existe ruta.
      sino = MsgBox("No existe ruta destino 'ASISTENCIA' ¿ Desea crearla ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
            Dim f As FileSystemObject
            Set f = New FileSystemObject
            f.CreateFolder (rutaArchivo)
            existeRuta = True
      End If
   End If
   
   If existeRuta Then
     ' Carga archivo.
     Dim rsCantExistente As New ADODB.Recordset
     Dim esValido As Boolean
     esValio = True
     Call valida_campos(esValio)
    
     If esValio Then
If rbtMes.Value = True Then
sino = MsgBox("¿Esta seguro de subir la asistencia del MES con los siguientes datos?" & vbCrLf & "Gestion: " & cmb_gestion_rep.Text & vbCrLf & "Mes:" & cmb_mes_ini.Text & vbCrLf & "Equipo Biométrico: " & cmb_equipo.Text & vbCrLf & "Departamento: " & cmb_departamento.Text, vbYesNo + vbQuestion, "Atención")
End If
If rbtDia(0).Value = True Then
sino = MsgBox("¿Esta seguro de subir la asistencia de un DÍA con los siguientes datos?" & vbCrLf & "Fecha:" & DtpFecha.Value & vbCrLf & "Equipo Biométrico: " & cmb_equipo.Text & vbCrLf & "Departamento: " & cmb_departamento.Text, vbYesNo + vbQuestion, "Atención")
End If
If sino = vbYes Then
        GLCarpeta = ""
        Dim dia As String, mes As String
        
        Fecha = DtpFecha.Value
        Call ObtenerDiaMes(DatePart("m", Fecha), mes)
        ' Tipo de exportación por mes o dia.
        If rbtMes.Value = True Then
           NombreArchivo = UCase(Trim$(Replace(cmb_departamento, " ", "")) & "_" & Trim$(cmb_equipo) & "_" & cmb_gestion_rep.Text & txt_mes.Text)
        Else
           Call ObtenerDiaMes(DatePart("d", Fecha), dia)
           NombreArchivo = UCase(Trim$(Replace(cmb_departamento, " ", "")) & "_" & Trim$(cmb_equipo) & "_" & DatePart("yyyy", Fecha) & mes & dia)
        End If
        ' Asigna nombre archivo a variable global
        GLCarpeta2 = NombreArchivo
        rutaArchivo = App.Path & "\ASISTENCIA\"
        GlArch = "ASIS"
        Frmexporta.DirDestino.Path = rutaArchivo
        Frmexporta.DirDestino2.Path = rutaArchivo
        Frmexporta.Show vbModal
        ' Verifica si nombre de hoja es diferente a vacio.
        If GLCarpeta <> "" Then
             MsgBox "El archivo " & NombreArchivo & " se copio correctamente."
             btnImportarDato.Enabled = True
        End If
        
        ' Consulta verifica si los datos del archivo con NombreArchivo se registraron.
        rsCantExistente.Open "SELECT COUNT(*) AS 'cuantos' FROM auxiliar_asistencia AS ax INNER JOIN ro_controlasistencia AS ctr ON ax.Id_AuxAsis =ctr.Id_AuxAsis WHERE ax.Nombre_Archivo = '" & NombreArchivo & "' ", db, adOpenStatic
        rsCantExistente.MoveFirst
        
        If rsCantExistente![Cuantos] > 0 Then SiEstaImportado = True Else SiEstaImportado = False
        
        rsCantExistente.Close
    
    db.Execute "delete auxiliar_asistencia "
    btnImportarDato.Visible = True
    Image1.Visible = False
    Image2.Visible = True
     End If

    End If
    
End If
End Sub

Private Sub CargarControles()
    Dim rsDepartamento As New ADODB.Recordset
    Dim rsEquipo As New ADODB.Recordset
    rsDepartamento.Open "SELECT DISTINCT * FROM gc_departamento ", db, adOpenStatic
    rsDepartamento.MoveFirst
    With Me.cmb_departamento
        .Clear
        Do
            .AddItem rsDepartamento![depto_descripcion]
            rsDepartamento.MoveNext
        Loop Until rsDepartamento.EOF
    End With
    ' Equipo
    rsEquipo.Open "SELECT * FROM rc_equipo_asistencia ", db, adOpenStatic
    rsEquipo.MoveFirst
    With Me.cmb_equipo
        .Clear
        Do
            .AddItem rsEquipo![descripcion_asist]
            rsEquipo.MoveNext
        Loop Until rsEquipo.EOF
    End With
    
'UserForm_Initialize_Exit:
    On Error Resume Next
    rsDepartamento.Close
    rsEquipo.Close
End Sub


Private Sub valida_campos(esValio)
  Dim inicial As Integer
  If rbtDia(0).Value = True Then
  If DtpFecha.Value = "" Then
    MsgBox " El campo Fecha es requerido."
    esValio = False
  End If
  
  End If
   If rbtMes.Value = True Then
   If txt_mes.Text = "0" Or txt_mes.Text = "" Then
    MsgBox " El campo Mes requerido."
    esValio = False
  End If
   
   End If
  If cmb_departamento = "" Then
    MsgBox " Seleccione un departamento."
    esValio = False
  End If
   If cmb_equipo = "" Then
    MsgBox " Seleccione un equipo."
    esValio = False
  End If

End Sub


Private Sub btnImportarDato_Click()
        btnCargarArchivo.Enabled = False
        btnImportarDato.Enabled = False
        
        If SiEstaImportado Then
                sino = MsgBox("¿Existen datos para '" & NombreArchivo & "',desea reemplazarlos?", vbQuestion + vbYesNo, "Confirmando Impresión... ")
                If sino = vbYes Then
                   Call EliminarDatoAnterior
                   MsgBox "Los datos anteriores se eliminaron"
                    Call ImportarDato
                    db.Execute "UPDATE ro_controlasistencia SET ges_gestion = year(Fecha_control), Mes_control = month(Fecha_control), Dia_control= day(Fecha_control)"
                End If
        Else
           Call ImportarDato
           db.Execute "UPDATE ro_controlasistencia SET ges_gestion = year(Fecha_control), Mes_control = month(Fecha_control), Dia_control= day(Fecha_control)"
        End If
        Fra_ABM.Enabled = False
        ProgressBar1.Visible = False
        Image2.Visible = False
End Sub

' Eliminar datos de anterior importacion
Private Sub EliminarDatoAnterior()
     db.Execute " DELETE FROM ro_controlasistencia WHERE Id_AuxAsis IN (SELECT Id_AuxAsis FROM auxiliar_asistencia WHERE Nombre_Archivo = '" & NombreArchivo & "') "
     db.Execute " DELETE FROM auxiliar_asistencia WHERE Nombre_Archivo = '" & NombreArchivo & "' "
End Sub


' Importar excel
Private Sub ImportarDato()
  On Error GoTo ErrorHandler
            
        lblMensaje.Visible = True
        MsgBox " Se inicia el proceso de importación de datos."
                
        Dim conExcel As New ADODB.Connection
        Dim rsExcel As New ADODB.Recordset
        
        Dim rsTablaAuxiliar As ADODB.Recordset
        
        Dim sqlDatosAux As String
        Dim indice As Integer
        
        If conExcel.State = adStateOpen Then conExcel.Close
        If rsExcel.State = adStateOpen Then rsExcel.Close
        
        Dim origenExcel As String
        Dim ruta As String
        origenExcel = NombreArchivo '
        ' ruta = App.Path & "\ASISTENCIA\" & NombreArchivo & ".xls"
        ruta = App.Path & "\ASISTENCIA\" & NombreArchivo & "." & GlExtension
        
        '--------------------------------- Obtiene nombre de hoja
'        Dim ObjExcel As Excel.Application
'        Dim ObjExcelLibro As Excel.Workbook
'        Set ObjExcel = New Excel.Application
'        Set ObjExcelLibro = ObjExcel.Workbooks.Open(ruta)
'
'        If ObjExcelLibro.Sheets.Count > 0 Then
'            ' Asigna nombre de primera hoja.
'            GLCarpeta = ObjExcelLibro.Sheets(1).Name
'        End If
'        ObjExcelLibro.Close
'        ObjExcel.Quit
'        Set ObjExcelLibro = Nothing
'        Set ObjExcel = Nothing
       
        '---------------------------------
        
        ' Coneccion a excel
'        conExcel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'            "Data Source= " & ruta & ";" & _
'                "Extended Properties=""Excel 8.0;"";"
                
'         conExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'            "Data Source= " & ruta & ";" & _
'                "Extended Properties=""Excel 12.0 Xml;"";"
        
        'conExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source= '" & ruta & "'; Extended Properties= Excel 12.0 Xml; HDR=YES; IMEX=1"
        
        conExcel.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source= '" & ruta & "'; Extended Properties= Excel 12.0 Xml; HDR=YES; IMEX=1"
        
        'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\myFolder\myExcel2007file.xlsx;
        'Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=1";
        
        ' Consulta obtiene datos de excel.
        ' GLCarpeta contiene nombre de hoja desde frmexport
         rsExcel.Open "SELECT * FROM [" & GLCarpeta & "$]", conExcel, 3, 1
        
        ' INSERTA REGISTROS A TABLA AUXILIAR
         indice = 0
        ' Variables de registros auxiliar
        Dim nroxl As Integer, cantRegistro As Integer
        Dim sql As String
        Dim sqlValue As String
        cantRegistro = 1
        'JQ
       CANTOT = rsExcel.RecordCount
       ProgressBar1.Visible = True
        With ProgressBar1
            .Max = CANTOT     'rs_datos6.RecordCount
            .Min = 0
            .Value = 0
        End With
        While Not rsExcel.EOF
                If rsExcel.Fields(0) <> "" Or rsExcel.Fields(0) <> Nulo Then
                    For indice = 0 To rsExcel.Fields.Count - 1
                        sqlDatosAux = sqlDatosAux & "'" & rsExcel.Fields(indice).Value & "',"
                    Next
                End If
                
                If sqlValue = "" And Trim$(sqlDatosAux) <> "" Then
                     sqlValue = " (" & Mid(sqlDatosAux, 1, Len(sqlDatosAux) - 1) & " ,'" & GLCarpeta2 & "' )"
                Else
                     If Trim$(sqlDatosAux) <> "" Then
                           sqlValue = sqlValue & ", (" & Mid(sqlDatosAux, 1, Len(sqlDatosAux) - 1) & " ,'" & GLCarpeta2 & "' )"
                     End If
                End If
                ' Sql server solo permite registrar 1000 registros por insert.
                'If cantRegistro = 1000 Then
                If cantRegistro = 100 Then
                     sql = sql & " INSERT INTO auxiliar_asistencia (Nro,AC_No,Cedula_No,Nombre,Auto_asigna,Fecha,Horario,HoraEnt,HoraSal,Marc_Ent,Marc_Sal,Normal,TiemReal,Tardanza,SalioTempr,Falta,HoraExtra,WorkTime,Excepcion,Debe_C_In,Debe_C_Sal,Depto,NDays,FinSemana,Feriado,TiemAsist,NDiasOT,FinSemanaOT,FeriadoOT, Nombre_Archivo) VALUES  " & sqlValue & " ;"
                     cantRegistro = 0
                     sqlValue = ""
                      ' Inserta registros.
                     db.Execute sql
                     sql = ""
                End If
                
              sqlDatosAux = ""
              rsExcel.MoveNext
              cantRegistro = cantRegistro + 1
              ProgressBar1.Value = ProgressBar1.Value + 1
        Wend
        
        If sqlValue <> "" Then
            sql = sql & " INSERT INTO auxiliar_asistencia (Nro,AC_No,Cedula_No,Nombre,Auto_asigna,Fecha,Horario,HoraEnt,HoraSal,Marc_Ent,Marc_Sal,Normal,TiemReal,Tardanza,SalioTempr,Falta,HoraExtra,WorkTime,Excepcion,Debe_C_In,Debe_C_Sal,Depto,NDays,FinSemana,Feriado,TiemAsist,NDiasOT,FinSemanaOT,FeriadoOT, Nombre_Archivo) VALUES  " & sqlValue & " ;"
        End If
        If sql <> "" Then
             ' Inserta registros.
            db.Execute sql
        End If
        
         ' INSERTA REGISTROS A TABLA OFICIAL            WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
         Set rsTablaAuxiliar = New ADODB.Recordset
         
           If rsTablaAuxiliar.State = 1 Then rsTablaAuxiliar.Close
            Dim sqlSelect As String
            ' Tipo de exportación por mes o dia.
            If rbtMes.Value = True Then
               ' Consulta por mes
                sqlSelect = "SELECT * FROM auxiliar_asistencia WHERE MONTH(Fecha) = '" & txt_mes.Text & "' AND YEAR(Fecha) = '" & cmb_gestion_rep.Text & "' AND Nombre_Archivo = '" & NombreArchivo & "' "
            Else
               ' Consulta por dia
                sqlSelect = "SELECT * FROM auxiliar_asistencia WHERE DAY(Fecha) = DAY('" & Fecha & "') AND MONTH(Fecha) = MONTH('" & Fecha & "') AND YEAR(Fecha) = YEAR('" & Fecha & "') AND Nombre_Archivo = '" & NombreArchivo & "' "
            End If
 
            rsTablaAuxiliar.Open sqlSelect, db, 3, 1
          
           sqlValue = ""
           cantRegistro = 1
           sql = ""
           ' Recorre registros de auxiliar asistencia
           Dim strValorInser As String
           Dim esdebein As String, esfalta As String, esdebesal As String
           Dim Nro As String, ac_no As String
           Dim tardanzaval As String
           Dim normal As String, tiemporeal As String, nday As String, ndiasot As String, tardanza As String
           Dim minutoTardanza As Integer
           Dim Formato As String
           Formato = "#,##0"
           
           If rsTablaAuxiliar.RecordCount > 0 Then
             rsTablaAuxiliar.MoveFirst
             While Not rsTablaAuxiliar.EOF
                   
                   Call ObtenerValorNumero(rsTablaAuxiliar!Nro, Nro)
                   Call ObtenerValorNumero(rsTablaAuxiliar!ac_no, ac_no)
                   Call ObtenerValorNumero(rsTablaAuxiliar!normal, normal)
                   Call ObtenerValorNumero(rsTablaAuxiliar!TiemReal, tiemporeal)
                   Call ObtenerValorBool(rsTablaAuxiliar!Falta, esfalta)
                   Call ObtenerValorBool(rsTablaAuxiliar!Debe_C_In, esdebein)
                   Call ObtenerValorBool(rsTablaAuxiliar!Debe_C_Sal, esdebesal)
                   Call ObtenerValorNumero(rsTablaAuxiliar!NDays, nday)
                   Call ObtenerValorNumero(rsTablaAuxiliar!ndiasot, ndiasot)
                   
                   tardanzaval = rsTablaAuxiliar!tardanza
                   
                   If rsTablaAuxiliar!tardanza = "NULL" Then
                        tardanzaval = "00:00"
                   End If
                   If Trim(rsTablaAuxiliar!tardanza) = "" Then
                        tardanzaval = "00:00"
                   End If
                   
                   minutoTardanza = Format(DateDiff("n", "00:00", tardanzaval), Formato)
                                      
                   Dim tardanzaCadena As String
                   tardanzaCadena = rsTablaAuxiliar!tardanza
                   If tardanzaCadena = "" Then
                    tardanzaCadena = "0000"
                   Else
                    tardanzaCadena = Replace(rsTablaAuxiliar!tardanza, ":", "")
                   End If
                   
                   ' Cadena de datos para insert.
                strValorInser = " " & Nro & ", " & ac_no & ", '" & rsTablaAuxiliar!Cedula_No & "', " & _
                                " '" & rsTablaAuxiliar!Nombre & "', '" & CStr(rsTablaAuxiliar!Auto_asigna) & "', '" & CStr(rsTablaAuxiliar!Fecha) & "', " & _
                                " '" & CStr(rsTablaAuxiliar!Horario) & "', '" & Replace(rsTablaAuxiliar!HoraEnt, ":", "") & "', '" & CStr(rsTablaAuxiliar!HoraEnt) & "', " & _
                                " '" & Replace(rsTablaAuxiliar!horaSal, ":", "") & "', '" & CStr(rsTablaAuxiliar!horaSal) & "', '" & Replace(rsTablaAuxiliar!Marc_Ent, ":", "") & "', " & _
                                " '" & CStr(rsTablaAuxiliar!Marc_Ent) & "', '" & Replace(rsTablaAuxiliar!Marc_Sal, ":", "") & "', '" & CStr(rsTablaAuxiliar!Marc_Sal) & "', " & _
                                 " " & Replace(normal, ",", ".") & ", " & Replace(tiemporeal, ",", ".") & ", '" & tardanzaval & "', " & _
                                 " '" & CStr(rsTablaAuxiliar!SalioTempr) & "', " & esfalta & ", '" & Trim$(Replace(Replace(CStr(rsTablaAuxiliar!HoraExtra), "a.m.", ""), "p.m.", "")) & "', " & _
                                 " '" & CStr(rsTablaAuxiliar!WorkTime) & "', '" & CStr(rsTablaAuxiliar!Excepcion) & "', " & esdebein & ", " & _
                                 "  " & esdebesal & ", '" & CStr(rsTablaAuxiliar!Depto) & "', " & Replace(nday, ",", ".") & ", " & _
                                 " '" & CStr(rsTablaAuxiliar!FinSemana) & "', '" & CStr(rsTablaAuxiliar!Feriado) & "', '" & CStr(rsTablaAuxiliar!TiemAsist) & "', " & _
                                 "  " & Replace(ndiasot, ",", ".") & ", '" & CStr(rsTablaAuxiliar!FinSemanaOT) & "', '" & CStr(rsTablaAuxiliar!FeriadoOT) & "', " & rsTablaAuxiliar!Id_AuxAsis & " , " & _
                                 " '" & tardanzaCadena & "', '" & Replace(rsTablaAuxiliar!TiemAsist, ":", "") & "' " & " , " & _
                                 " " & minutoTardanza & " "
                
                If Nro <> "NULL" Then
                    If sqlValue = "" Then
                         sqlValue = " (" & strValorInser & ")"
                    Else
                         sqlValue = sqlValue & ", (" & strValorInser & ") "
                    End If
                End If
                 ' Sql server solo permite registrar 1000 registros por insert.
                'If cantRegistro = 1000 Then
                If cantRegistro = 100 Then
                     sql = sql & " INSERT INTO ro_controlasistencia (Correl,Correl_ac,beneficiario_codigo,Nombre,Autoasigna,Fecha_control,TipoHorario,Hora1, HoraUno,Hora2,HoraDos,Hora3,HoraTres,Hora4,HoraCuatro,Normal,TiemReal,Tardanza,SalioTempr,EsFalta,HoraExtra,WorkTime,Excepcion,Debe_C_In,Debe_C_Sal,Depto,NDays,FinSemana,Feriado,TiemAsist,NDiasOT,FinSemanaOT,FeriadoOT, Id_AuxAsis,TardanzaCadena,TiempoTrabajoCadena, AtrasoMin1) VALUES  " & sqlValue & " ;"
                     cantRegistro = 0
                     sqlValue = ""
                End If
                rsTablaAuxiliar.MoveNext
                cantRegistro = cantRegistro + 1
             Wend
             
              If sqlValue <> "" Then
                    sql = sql & " INSERT INTO ro_controlasistencia (Correl,Correl_ac,beneficiario_codigo,Nombre,Autoasigna,Fecha_control,TipoHorario,Hora1, HoraUno,Hora2,HoraDos,Hora3,HoraTres,Hora4,HoraCuatro,Normal,TiemReal,Tardanza,SalioTempr,EsFalta,HoraExtra,WorkTime,Excepcion,Debe_C_In,Debe_C_Sal,Depto,NDays,FinSemana,Feriado,TiemAsist,NDiasOT,FinSemanaOT,FeriadoOT, Id_AuxAsis,TardanzaCadena,TiempoTrabajoCadena, AtrasoMin1) VALUES " & sqlValue & " ;"
              End If
                
              If sql <> "" Then
                     ' Inserta registros.
                    db.Execute sql
              End If
             lblMensaje.Visible = False
             MsgBox "Los datos se registraron en la tabla oficial."
             
        Else
          MsgBox " No existen datos coincidentes."
             
         End If
           
           Call ABRIR_TABLA
           ' VERIFICA REGISTROS
ErrorHandler:
    If Trim(Err.Description) <> "" Then
       lblMensaje.Visible = False
       MsgBox Err.Description, , "Error"
    End If
End Sub

' Validacion cabecera.
Private Function ValidarCabecera(registros As ADODB.Recordset) As String ' Notice the As String
            Dim Mensaje As String
            Mensaje = ""
            
            Dim nombreEnc As String
            Dim nomCabecera As String
            For i = 0 To rsExcel.Fields.Count - 1
              nomCabecera = rsExcel.Fields(i).Name
              If i = 0 Then
                 If LTrim$(nomCabecera) <> "No#" Then
                    Mensaje = Mensaje & " La columna " & (i + 1) & " debe nombrarse 'No.'"
                 End If
              End If
              
            Next
            
            'return mensaje
End Function

' Retorna valor por defecto campo decimal o entero vacio
Private Function ObtenerValorNumero(dato As String, rvalor As String) As String
    If LTrim$(dato) = "" Then
       rvalor = "0"
    Else
       rvalor = dato
    End If
End Function

' Retorna valor por defecto campo bool
Private Function ObtenerValorBool(dato As String, rvalor As String) As String
                   If LTrim$(dato) = "True" Then
                        rvalor = "1"
                   ElseIf LTrim$(dato) = "False" Then
                        rvalor = "0"
                   Else
                        rvalor = "NULL"
                   End If
End Function

Private Function ObtenerDiaMes(dato As String, rvalor As String) As String
                   rvalor = Trim$(dato)
                   If Len(dato) = 1 Then
                        rvalor = "0" & Trim$(dato)
                   End If
End Function


Private Sub BtnSalir_Click()

  Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub ABRIR_TABLA()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = " SELECT CASE esfalta WHEN 1 THEN 'SI' ELSE 'NO' END AS siFalta, * FROM ro_controlasistencia WHERE Id_AuxAsis IN (SELECT Id_AuxAsis FROM auxiliar_asistencia WHERE Nombre_Archivo = '" & NombreArchivo & "') "
 
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
dtc_rep_cod.Text = "%"
dtc_rep_det.Text = "TODAS LAS PLANILLAS"
dtc_depto.Text = "%"
Else
dtc_rep_cod.Text = ""
dtc_rep_det.Text = ""
End If
End Sub

Private Sub rbtDia_Click(Index As Integer)
If rbtDia(0).Value = True Then
lbl_inicial(0).Visible = True
DtpFecha.Visible = True
lbl_inicial(1).Visible = False
cmb_mes_ini.Visible = False
cmb_gestion_rep.Visible = False
End If
End Sub

Private Sub rbtMes_Click()

If rbtMes.Value = True Then

lbl_inicial(0).Visible = False
DtpFecha.Visible = False
lbl_inicial(1).Visible = True
cmb_mes_ini.Visible = True
cmb_gestion_rep.Visible = True
End If

End Sub

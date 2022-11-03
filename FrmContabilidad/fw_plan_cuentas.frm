VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fw_plan_cuentas 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Contabilidad - Plan Cuentas"
   ClientHeight    =   9165
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "fw_plan_cuentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra_Det4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle Nivel - 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   720
      TabIndex        =   85
      Top             =   5160
      Visible         =   0   'False
      Width           =   13455
      Begin VB.CommandButton BtnSalir4 
         Caption         =   "Salir"
         Height          =   645
         Left            =   12600
         Picture         =   "fw_plan_cuentas.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   840
         Width           =   765
      End
      Begin VB.CommandButton BtnAceptar4 
         Caption         =   "Aceptar"
         Height          =   645
         Left            =   11760
         Picture         =   "fw_plan_cuentas.frx":0E44
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   840
         Width           =   750
      End
      Begin MSDataGridLib.DataGrid dg_datos4 
         Bindings        =   "fw_plan_cuentas.frx":114E
         Height          =   1215
         Left            =   120
         TabIndex        =   88
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   2143
         _Version        =   393216
         BackColor       =   12648447
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "correl"
            Caption         =   "Correl"
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
            DataField       =   "Cuenta"
            Caption         =   "Cuenta"
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
            DataField       =   "NombreCta"
            Caption         =   "Nombre_Cuenta"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fra_Det3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle Nivel - 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   480
      TabIndex        =   81
      Top             =   5040
      Visible         =   0   'False
      Width           =   13455
      Begin VB.CommandButton BtnSalir3 
         Caption         =   "Salir"
         Height          =   645
         Left            =   12600
         Picture         =   "fw_plan_cuentas.frx":1167
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   840
         Width           =   765
      End
      Begin VB.CommandButton BtnAceptar3 
         Caption         =   "Aceptar"
         Height          =   645
         Left            =   11760
         Picture         =   "fw_plan_cuentas.frx":15A9
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   840
         Width           =   750
      End
      Begin MSDataGridLib.DataGrid dg_datos3 
         Bindings        =   "fw_plan_cuentas.frx":18B3
         Height          =   1215
         Left            =   120
         TabIndex        =   84
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   2143
         _Version        =   393216
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "correl"
            Caption         =   "Correl"
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
            DataField       =   "Cuenta"
            Caption         =   "Cuenta"
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
            DataField       =   "NombreCta"
            Caption         =   "Nombre_Cuenta"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fra_Det2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle Nivel - 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   600
      TabIndex        =   77
      Top             =   5040
      Visible         =   0   'False
      Width           =   13815
      Begin VB.CommandButton BtnAceptar2 
         Caption         =   "Aceptar"
         Height          =   645
         Left            =   12120
         Picture         =   "fw_plan_cuentas.frx":18CC
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   840
         Width           =   750
      End
      Begin VB.CommandButton BtnSalir2 
         Caption         =   "Salir"
         Height          =   645
         Left            =   12960
         Picture         =   "fw_plan_cuentas.frx":1BD6
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   840
         Width           =   765
      End
      Begin MSDataGridLib.DataGrid dg_datos2 
         Bindings        =   "fw_plan_cuentas.frx":2018
         Height          =   1215
         Left            =   120
         TabIndex        =   80
         Top             =   240
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   2143
         _Version        =   393216
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "correl"
            Caption         =   "Correl"
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
            DataField       =   "Cuenta"
            Caption         =   "Cuenta"
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
            DataField       =   "NombreCta"
            Caption         =   "Nombre_Cuenta"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   22
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "fw_plan_cuentas.frx":2031
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   31
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5520
         Picture         =   "fw_plan_cuentas.frx":27F3
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   30
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4200
         Picture         =   "fw_plan_cuentas.frx":30C0
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   29
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6960
         Picture         =   "fw_plan_cuentas.frx":3875
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   28
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "fw_plan_cuentas.frx":40A8
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   27
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1320
         Picture         =   "fw_plan_cuentas.frx":47F4
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   26
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "fw_plan_cuentas.frx":5109
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   25
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10800
         Picture         =   "fw_plan_cuentas.frx":58C8
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Picture         =   "fw_plan_cuentas.frx":5D0A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
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
         TabIndex        =   32
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
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "fw_plan_cuentas.frx":5F14
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   20
         Top             =   0
         Width           =   1455
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         Picture         =   "fw_plan_cuentas.frx":6800
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   19
         Top             =   0
         Width           =   1335
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
         Left            =   13215
         TabIndex        =   21
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO"
      ForeColor       =   &H00800000&
      Height          =   2880
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   14652
      Begin VB.OptionButton OptFilGral3 
         Caption         =   "Aprobadas"
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
         Left            =   10200
         TabIndex        =   16
         Top             =   2490
         Width           =   2175
      End
      Begin VB.OptionButton OptFilGral1 
         Caption         =   "No Aprobadas"
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
         Left            =   6480
         TabIndex        =   15
         Top             =   2490
         Width           =   1815
      End
      Begin VB.OptionButton OptFilGral2 
         Caption         =   "Todas"
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
         Left            =   3120
         TabIndex        =   14
         Top             =   2490
         Width           =   915
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "fw_plan_cuentas.frx":6FD6
         Height          =   2175
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   14280
         _ExtentX        =   25188
         _ExtentY        =   3836
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "Cuenta"
            Caption         =   "Cuenta"
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
            DataField       =   "SubCta1"
            Caption         =   "SubCta1"
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
            DataField       =   "SubCta2"
            Caption         =   "SubCta2"
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
            DataField       =   "NombreCta"
            Caption         =   "Nombre de la Cuenta"
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
            DataField       =   "Aux1"
            Caption         =   "Aux1"
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
            DataField       =   "Aux2"
            Caption         =   "Aux2"
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
            DataField       =   "Aux3"
            Caption         =   "Aux3"
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
         BeginProperty Column07 
            DataField       =   "Mov"
            Caption         =   "Titulo/Mov"
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
         BeginProperty Column08 
            DataField       =   "nivel"
            Caption         =   "Nivel"
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
         BeginProperty Column09 
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
         BeginProperty Column10 
            DataField       =   "fecha_registro"
            Caption         =   "Fecha_Reg."
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
         BeginProperty Column11 
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
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   5760
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   2400
         Width           =   14265
         _ExtentX        =   25162
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
         BackColor       =   -2147483633
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
   Begin VB.Frame Fra_ABM 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Registro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5265
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   14652
      Begin VB.Frame Fra_Det1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle Nivel - 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1575
         Left            =   480
         TabIndex        =   56
         Top             =   1320
         Visible         =   0   'False
         Width           =   13455
         Begin MSDataGridLib.DataGrid dg_datos1 
            Bindings        =   "fw_plan_cuentas.frx":6FEE
            Height          =   1215
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   2143
            _Version        =   393216
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "correl"
               Caption         =   "Correl"
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
               DataField       =   "Cuenta"
               Caption         =   "Cuenta"
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
               DataField       =   "NombreCta"
               Caption         =   "Nombre_Cuenta"
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
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton BtnSalir1 
            Caption         =   "Salir"
            Height          =   645
            Left            =   12480
            Picture         =   "fw_plan_cuentas.frx":7007
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   840
            Width           =   765
         End
         Begin VB.CommandButton BtnAceptar1 
            Caption         =   "Aceptar"
            Height          =   645
            Left            =   11640
            Picture         =   "fw_plan_cuentas.frx":7449
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   840
            Width           =   750
         End
      End
      Begin VB.Frame Fra_ABM2 
         BackColor       =   &H00C0C0C0&
         Height          =   1815
         Left            =   240
         TabIndex        =   60
         Top             =   1080
         Width           =   14175
         Begin VB.Label dtc_Aux4 
            Alignment       =   2  'Center
            BackColor       =   &H80000011&
            Caption         =   "todos"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   1320
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label dtc_Aux3 
            Alignment       =   2  'Center
            BackColor       =   &H80000011&
            Caption         =   "todos"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   960
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label dtc_Aux2 
            Alignment       =   2  'Center
            BackColor       =   &H80000011&
            Caption         =   "todos"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   600
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label dtc_Aux1 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "todos"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label dtc_codigo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "todos"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   76
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label dtc_desc1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3840
            TabIndex        =   75
            Top             =   240
            Width           =   9735
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cuenta Nivel - 1"
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
            Left            =   480
            TabIndex        =   74
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cuenta Nivel - 2"
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
            Left            =   480
            TabIndex        =   73
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cuenta Nivel - 3"
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
            Left            =   480
            TabIndex        =   72
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cuenta Nivel - 4"
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
            Left            =   480
            TabIndex        =   71
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label dtc_codigo2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "todos"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   66
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label dtc_codigo3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "todos"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   65
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label dtc_codigo4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "todos"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   64
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label dtc_desc2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3840
            TabIndex        =   63
            Top             =   600
            Width           =   9735
         End
         Begin VB.Label dtc_desc3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3840
            TabIndex        =   62
            Top             =   960
            Width           =   9735
         End
         Begin VB.Label dtc_desc4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3840
            TabIndex        =   61
            Top             =   1320
            Width           =   9735
         End
      End
      Begin VB.Frame Fra_Aux 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipos de Auxiliares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1095
         Left            =   240
         TabIndex        =   49
         Top             =   4080
         Width           =   14175
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "fw_plan_cuentas.frx":7753
            DataField       =   "aux1"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   50
            Top             =   360
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   14737632
            ListField       =   "aux"
            BoundColumn     =   "aux"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo dtc_codigo6 
            Bindings        =   "fw_plan_cuentas.frx":776C
            DataField       =   "aux2"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   8040
            TabIndex        =   51
            Top             =   360
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   14737632
            ListField       =   "aux"
            BoundColumn     =   "aux"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo dtc_codigo7 
            Bindings        =   "fw_plan_cuentas.frx":7785
            DataField       =   "aux3"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   12600
            TabIndex        =   52
            Top             =   360
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   14737632
            ListField       =   "aux"
            BoundColumn     =   "aux"
            Text            =   "0000"
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "fw_plan_cuentas.frx":779E
            DataField       =   "Aux1"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "descripcion"
            BoundColumn     =   "Aux"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc6 
            Bindings        =   "fw_plan_cuentas.frx":77B7
            DataField       =   "Aux2"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4800
            TabIndex        =   3
            Top             =   720
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "descripcion"
            BoundColumn     =   "Aux"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc7 
            Bindings        =   "fw_plan_cuentas.frx":77D0
            DataField       =   "Aux3"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   9360
            TabIndex        =   4
            Top             =   720
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "descripcion"
            BoundColumn     =   "Aux"
            Text            =   "Todos"
         End
         Begin VB.Label lbl_enlace5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Auxiliar - 1"
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
            TabIndex        =   55
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lbl_enlace6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Auxiliar - 2"
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
            Left            =   4800
            TabIndex        =   54
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lbl_enlace7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Auxiliar - 3"
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
            Left            =   9480
            TabIndex        =   53
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame Fra_ABM3 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FFFF80&
         Height          =   1215
         Left            =   240
         TabIndex        =   39
         Top             =   2880
         Width           =   14175
         Begin VB.TextBox TxtConcepto 
            DataField       =   "NombreCta"
            DataSource      =   "Ado_datos"
            Height          =   405
            Left            =   2880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Top             =   600
            Width           =   10995
         End
         Begin VB.Label Txt_campo4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "00"
            DataField       =   "SubCta2"
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
            Height          =   255
            Left            =   12960
            TabIndex        =   48
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Txt_campo3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "SubCta1"
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
            Height          =   255
            Left            =   7440
            TabIndex        =   47
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sub Cuenta - 1"
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
            Left            =   4920
            TabIndex        =   46
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sub Cuenta - 2"
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
            Left            =   9720
            TabIndex        =   45
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Codigo de la Cuenta"
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
            TabIndex        =   44
            Top             =   240
            Width           =   1830
         End
         Begin VB.Label Txt_campo2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "cuenta"
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2880
            TabIndex        =   43
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Descripcion de la Cuenta"
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
            TabIndex        =   42
            Top             =   720
            Width           =   2250
         End
         Begin VB.Label Txt_campo9 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
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
            Height          =   255
            Left            =   360
            TabIndex        =   41
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Txt_campo5 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
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
            Height          =   255
            Left            =   360
            TabIndex        =   40
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Frame Fra_ABM1 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FFFF80&
         Height          =   855
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   14175
         Begin VB.ComboBox dtc_cuenta 
            DataField       =   "nivel"
            DataSource      =   "Ado_datos"
            Height          =   315
            ItemData        =   "fw_plan_cuentas.frx":77E9
            Left            =   2280
            List            =   "fw_plan_cuentas.frx":77FC
            TabIndex        =   0
            Text            =   "0"
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nivel de la Cuenta"
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
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Estado de Registro"
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
            Left            =   10560
            TabIndex        =   37
            Top             =   360
            Width           =   1740
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Titulo/Subtitulo/Detalle"
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
            Left            =   5400
            TabIndex        =   36
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label Txt_campo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "D"
            DataField       =   "Mov"
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
            Height          =   255
            Left            =   7560
            TabIndex        =   35
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Txt_estado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   12480
            TabIndex        =   34
            Top             =   360
            Width           =   855
         End
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
      TabIndex        =   5
      Top             =   9165
      Width           =   15120
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   10
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   13920
      Top             =   9000
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
      Top             =   9000
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
      Top             =   9000
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
      Top             =   9000
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
      Left            =   6960
      Top             =   9000
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
      Left            =   9240
      Top             =   9000
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
      Left            =   11640
      Top             =   9000
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
      Left            =   120
      Top             =   9360
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
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
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
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "fw_plan_cuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset

Dim rs_det0 As New ADODB.Recordset
Dim rs_det1 As New ADODB.Recordset
Dim rs_det2 As New ADODB.Recordset
Dim rs_det3 As New ADODB.Recordset
Dim rs_det4 As New ADODB.Recordset
Dim rs_det5 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String


Dim var_cod, VAR_COD2, VAR_COD1, VAR_COD3 As String
Dim VAR_VAL As String
Dim VAR_SUB1 As String
Dim VAR_SW As String
Dim VAR_CTA, VAR_SUB2 As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAceptar1_Click()
  fw_plan_cuentas.dtc_aux1.Caption = dg_datos1.Columns(0)
  fw_plan_cuentas.dtc_codigo1.Caption = dg_datos1.Columns(1)
  fw_plan_cuentas.dtc_desc1.Caption = dg_datos1.Columns(2)
  
  Fra_Det2.Visible = False
  TxtConcepto.SetFocus
  
    If dtc_cuenta.Text > 2 Then
        Fra_Det2.Visible = True
        
        dtc_codigo1.Visible = True
        dtc_desc1.Visible = True
        dtc_codigo2.Visible = True
        dtc_desc2.Visible = True
        dtc_codigo3.Visible = False
        dtc_desc3.Visible = False
        dtc_codigo4.Visible = False
        dtc_desc4.Visible = False

        Fra_Aux.Visible = False
'        Fra_ABM3.Enabled = True
'        Call ABRIR_NIVEL1
        Call ABRIR_NIVEL2
        Call ABRIR_NIVEL3
    End If
   Fra_Det1.Visible = False
End Sub

Private Sub BtnAceptar2_Click()
  fw_plan_cuentas.Dtc_aux2.Caption = dg_datos2.Columns(0)
  fw_plan_cuentas.dtc_codigo2.Caption = dg_datos2.Columns(1)
  fw_plan_cuentas.dtc_desc2.Caption = dg_datos2.Columns(2)

  Fra_Det3.Visible = False
   TxtConcepto.SetFocus
   
  If dtc_cuenta.Text > 3 Then
         Fra_Det3.Visible = True
        'Fra_Det2.Visible = False
        
        dtc_codigo1.Visible = True
        dtc_desc1.Visible = True
        dtc_codigo2.Visible = True
        dtc_desc2.Visible = True
        dtc_codigo3.Visible = True
        dtc_desc3.Visible = True
        dtc_codigo4.Visible = False
        dtc_desc4.Visible = False

        Fra_Aux.Visible = False

        'Call ABRIR_NIVEL1
'        Call ABRIR_NIVEL2
        Call ABRIR_NIVEL3
    End If
      Fra_Det2.Visible = False
End Sub

Private Sub BtnAceptar3_Click()
  fw_plan_cuentas.dtc_aux3.Caption = dg_datos3.Columns(0)
  fw_plan_cuentas.dtc_codigo3.Caption = dg_datos3.Columns(1)
  fw_plan_cuentas.dtc_desc3.Caption = dg_datos3.Columns(2)
  
  Fra_Det3.Visible = False
   TxtConcepto.SetFocus
  If dtc_cuenta.Text > 4 Then
         Fra_Det4.Visible = True
'         Fra_Det3.Visible = True
'         Fra_Det1.Visible = True
         
        dtc_codigo1.Visible = True
        dtc_desc1.Visible = True
        dtc_codigo2.Visible = True
        dtc_desc2.Visible = True
        dtc_codigo3.Visible = True
        dtc_desc3.Visible = True
        dtc_codigo4.Visible = True
        dtc_desc4.Visible = True

        Fra_Aux.Visible = False
'        Call ABRIR_NIVEL1
       Call ABRIR_NIVEL4

    End If
    Fra_Det1.Visible = False
End Sub

Private Sub BtnAceptar4_Click()
  fw_plan_cuentas.dtc_aux4.Caption = dg_datos4.Columns(0)
  fw_plan_cuentas.dtc_codigo4.Caption = dg_datos4.Columns(1)
  fw_plan_cuentas.dtc_desc4.Caption = dg_datos4.Columns(2)
  
  Fra_Det4.Visible = False
   TxtConcepto.SetFocus
  If dtc_cuenta.Text > 5 Then

        dtc_codigo1.Visible = True
        dtc_desc1.Visible = True
        dtc_codigo2.Visible = True
        dtc_desc2.Visible = True
        dtc_codigo3.Visible = False
        dtc_desc3.Visible = False
        dtc_codigo4.Visible = False
        dtc_desc4.Visible = False

        Fra_Aux.Visible = False
'        Call ABRIR_NIVEL1
'        Call ABRIR_NIVEL2
'        Call ABRIR_NIVEL3
'        Call ABRIR_NIVEL4
'        Call ABRIR_AUX1
'        Call ABRIR_AUX2
'        Call ABRIR_AUX3
    End If
    Fra_Det1.Visible = False
End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If Ado_datos.Recordset!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         Ado_datos.Recordset!estado_codigo = "APR"
         Ado_datos.Recordset!fecha_registro = Date
         Ado_datos.Recordset!usr_codigo = glusuario
         Ado_datos.Recordset.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado (ERR) o Aprobado (APR) anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnBuscar_Click()
  On Error GoTo UpdateErr
    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexión = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = dg_datos
    ClBuscaGrid.QueryUtilizado = queryinicial
    Set ClBuscaGrid.RecordsetTrabajo = rs_datos
    'ClBuscaGrid.CamposVisibles = "11010011"
    ClBuscaGrid.Ejecutar
    Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelBatch
'        If mvBookMark > 0 Then
'          rs_datos.BookMark = mvBookMark
'        Else
'          rs_datos.MoveFirst
'        End If
        Call OptFilGral2_Click
        
        rs_datos.MoveFirst
        mbDataChanged = False
        Fra_Det4.Visible = False
        Fra_Det3.Visible = False
        Fra_Det2.Visible = False
        Fra_Det1.Visible = False
        FraNavega.Enabled = True
        Fra_ABM.Enabled = False
        Fra_ABM2.Visible = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        txt_codigo.Enabled = True
        dtc_desc1.Enabled = True
        VAR_SW = ""
    End If
     Exit Sub
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
   If ExisteReg(rs_datos!cuenta, rs_datos!subcta1, rs_datos!subcta2) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
   If Ado_datos.Recordset!estado_codigo = "APR" Then
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         Ado_datos.Recordset!estado_codigo = "ANL"
         Ado_datos.Recordset!fecha_registro = Date
         Ado_datos.Recordset!usr_codigo = glusuario
         Ado_datos.Recordset.UpdateBatch adAffectAll
      End If
   Else
      MsgBox "No se puede ANULAR un registro Elaborado (REG) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

'Private Sub BtnDesAprobar_Click()
'  On Error GoTo UpdateErr
'   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
'   If rs_datos!estado_codigo = "APR" Then
'      If sino = vbYes Then
'         rs_datos!estado_codigo = "REG"
'         rs_datos!fecha_registro = Date
'         rs_datos!usr_codigo = glusuario
'         rs_datos.UpdateBatch adAffectAll
'      End If
'   Else
'        MsgBox "No se puede DESAPROBAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
'   End If
'   Exit Sub
'UpdateErr:
'  MsgBox Err.Description
'End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If VAR_SW = "ADD" Then
'        Set rs_aux1 = New ADODB.Recordset
'        'Busca en la tabla actual el codigo del padre
'        SQL_FOR = "select * from gc_documentos_respaldo where clasif_codigo = '" & dtc_codigo1.Text & "'  "
'        'Set rs_aux1.DataSource = db.Execute(" EXEC gp_listar_mediante_codigo_gc_direccion_general '" & txt_codigo.Text & "' ")
'        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
''            MsgBox " CODIGO DUPLICADO, Vuelva a intentar..."
''            Exit Sub
'            var_cod = rs_aux1.RecordCount + 1
'        Else
'            var_cod = 1
'        End If
''        rs_datos!doc_codigo = RTrim(RTrim(dtc_codigo1.Text) + ".") + LTrim(Str(Val(var_cod)))
        VAR_SUB1 = "00"
        VAR_AUX1 = "00"
        VAR_AUX2 = "00"
        VAR_AUX3 = "00"
        If dtc_cuenta.Text = 1 Then
            Set rs_det1 = New ADODB.Recordset
            If rs_det1.State = 1 Then rs_det1.Close
            rs_det1.Open "Select IsNull(max(Cuenta),0) as Cta from cc_plan_nivel1 ", db, adOpenStatic
            VAR_CTA = rs_det1!Cta + 1000
'           Ado_datos.Recordset!mov = "T"
'           Ado_datos.Recordset!Niv_Public = "N"
        End If
        
        If dtc_cuenta.Text = 2 Then
            Set rs_det2 = New ADODB.Recordset
            If rs_det2.State = 1 Then rs_det2.Close
            'rs_det2.Open "Select IsNull(max(Cuenta),'" & dtc_codigo1.Text & "') as Cta from cc_plan_nivel2 where left(Cuenta,1)=  Left('" & dtc_codigo1.Text & "', 1)  ", db, adOpenStatic
            rs_det2.Open "Select IsNull(max(Cuenta),0) as Cta from cc_plan_nivel2 ", db, adOpenStatic
            VAR_CTA = rs_det2!Cta + 100
'            Ado_datos.Recordset!mov = "T"
'            Ado_datos.Recordset!Niv_Public = "N"
        End If
        
        If dtc_cuenta.Text = 3 Then
            Set rs_det3 = New ADODB.Recordset
            If rs_det3.State = 1 Then rs_det3.Close
            'rs_det3.Open "Select IsNull(max(Cuenta),'" & dtc_codigo2.Text & "') as Cta from cc_plan_nivel3 where left(Cuenta,2)=  Left('" & dtc_codigo2.Text & "', 2)  ", db, adOpenStatic
            rs_det3.Open "Select IsNull(max(Cuenta),0) as Cta from cc_plan_nivel3 ", db, adOpenStatic
            VAR_CTA = rs_det3!Cta + 10
'           Ado_datos.Recordset!mov = "T"
'           Ado_datos.Recordset!Niv_Public = "N"
        End If
        
        If dtc_cuenta.Text = 4 Then
            Set rs_det4 = New ADODB.Recordset
            If rs_det4.State = 1 Then rs_det4.Close
            'rs_det4.Open "Select IsNull(max(Cuenta),'" & dtc_codigo3.Text & "') as Cta from cc_plan_nivel4 where left(Cuenta,3)=  Left('" & dtc_codigo3.Text & "', 3)  ", db, adOpenStatic
            rs_det4.Open "Select IsNull(max(Cuenta),0) as Cta from cc_plan_nivel4 ", db, adOpenStatic
            VAR_CTA = rs_det4!Cta + 1
'           Ado_datos.Recordset!mov = "T"
'           Ado_datos.Recordset!Niv_Public = "N"
        End If
        
         If dtc_cuenta.Text = 5 Then
            VAR_CTA = dtc_codigo4

            Set rs_det5 = New ADODB.Recordset
            If rs_det5.State = 1 Then rs_det5.Close
            rs_det5.Open "Select IsNull(max(SubCta1),'00') as Scta from cc_plan_cuentas where cuenta='" & VAR_CTA & "' ", db, adOpenStatic
            If rs_det5!SCta < 10 Then
                VAR_SUB1 = "0" + Trim(Str(rs_det5!SCta + 1))
              Else
            
                VAR_SUB1 = rs_det5!SCta + 1
            End If
        VAR_AUX1 = dtc_codigo5
        VAR_AUX2 = dtc_codigo6
        VAR_AUX3 = dtc_codigo7
        End If

        db.Execute " INSERT INTO cc_plan_cuentas (Cuenta,SubCta1,SubCta2,NombreCta,Aux1,Aux2,Aux3,Mov,Niv_Public,Naturaleza,nivel,t_comision,estado_codigo,Usr_codigo,Fecha_registro) " & _
        " VALUES ('" & VAR_CTA & "','" & VAR_SUB1 & "','00','" & TxtConcepto.Text & "','" & VAR_AUX1 & "','" & VAR_AUX2 & "','" & VAR_AUX3 & "','T','N','D','" & dtc_cuenta.Text & "','N','REG','" & glusuario & "','" & Date & "')"
        
        Set rs_det0 = New ADODB.Recordset
            If rs_det0.State = 1 Then rs_det0.Close
            rs_det0.Open "Select ISNULL(correl,0)  as correl2 from CC_plan_cuentas where Cuenta= '" & VAR_CTA & "' and SubCta1= '00' and SubCta2='00' and nivel = '" & dtc_cuenta.Text & "'", db, adOpenStatic
            If rs_det0.RecordCount > 0 Then
                VAR_COD3 = rs_det0!correl2
            End If
       If dtc_cuenta.Text = 1 Then
     
            db.Execute " INSERT INTO cc_plan_nivel1 (Cuenta,SubCta1,SubCta2,NombreCta,Aux1,Aux2,Aux3,Mov,Niv_Public,Naturaleza,nivel,t_comision,estado_codigo,Usr_codigo,Fecha_registro, correl) " & _
              " VALUES ('" & VAR_CTA & "','00','00','" & TxtConcepto.Text & "','00','00','00','T','N','D','" & dtc_cuenta.Text & "','N','REG','" & glusuario & "','" & Date & "', " & VAR_COD3 & ")"
       End If
     
       If dtc_cuenta.Text = 2 Then
     
            db.Execute " INSERT INTO cc_plan_nivel2 (Cuenta,SubCta1,SubCta2,NombreCta,Aux1,Aux2,Aux3,Mov,Niv_Public,Naturaleza,nivel,t_comision,estado_codigo,Usr_codigo,Fecha_registro, correl) " & _
             " VALUES ('" & VAR_CTA & "','00','00','" & TxtConcepto.Text & "','00','00','00','T','N','D','" & dtc_cuenta.Text & "','N','REG','" & glusuario & "','" & Date & "', " & VAR_COD3 & ")"
       End If
        
       If dtc_cuenta.Text = 3 Then

           db.Execute " INSERT INTO cc_plan_nivel3 (Cuenta,SubCta1,SubCta2,NombreCta,Aux1,Aux2,Aux3,Mov,Niv_Public,Naturaleza,nivel,t_comision,estado_codigo,Usr_codigo,Fecha_registro, correl) " & _
           " VALUES ('" & VAR_CTA & "','00','00','" & TxtConcepto.Text & "','00','00','00','T','N','D','" & dtc_cuenta.Text & "','N','REG','" & glusuario & "','" & Date & "', " & VAR_COD3 & ")"
       End If

       If dtc_cuenta.Text = 4 Then

             db.Execute " INSERT INTO cc_plan_nivel4 (Cuenta,SubCta1,SubCta2,NombreCta,Aux1,Aux2,Aux3,Mov,Niv_Public,Naturaleza,nivel,t_comision,estado_codigo,Usr_codigo,Fecha_registro, correl) " & _
            " VALUES ('" & VAR_CTA & "','00','00','" & TxtConcepto.Text & "','00','00','00','T','N','D','" & dtc_cuenta.Text & "','N','REG','" & glusuario & "','" & Date & "', " & VAR_COD3 & ")"
       End If

'        If dtc_cuenta.Text = 5 Then
'
'             db.Execute " INSERT INTO cc_plan_cuentas (Cuenta,SubCta1,SubCta2,NombreCta,Aux1,Aux2,Aux3,Mov,Niv_Public,Naturaleza,nivel,t_comision,estado_codigo,Usr_codigo,Fecha_registro,correl) " & _
'            " VALUES ('" & VAR_CTA & "','" & VAR_SUB1 & "','00','" & TxtConcepto.Text & "','" & dtc_codigo5.Text & "','" & dtc_codigo6.Text & "','" & dtc_codigo7.Text & "','T','N','D','" & dtc_cuenta.Text & "','N','REG','" & glusuario & "','" & Date & "'," & VAR_COD3 & ")"
'        End If


End If
'        'rs_datos!subproceso_codigo = txt_codigo.Text ' Esto para codigos trascritos
'        Ado_datos.Recordset!estado_codigo = "REG"  ' no cambia
'
'        Ado_datos.Recordset!nivel = dtc_cuenta.Text 'Codigo del padre
'        Ado_datos.Recordset!cuenta = VAR_CTA
'        Ado_datos.Recordset!subcta1 = "00"
'        Ado_datos.Recordset!subcta2 = "00"
'
'        Ado_datos.Recordset!aux1 = "00"
'        Ado_datos.Recordset!AUX2 = "00"
'        Ado_datos.Recordset!aux3 = "00"
'
'        Ado_datos.Recordset!Naturaleza = "D"
'        Ado_datos.Recordset!t_comision = "N"
'
'        'Guarda en el Padre, en el campo ctrl de correlativos para codigos que se generan
''        db.Execute "Update gc_direccion_general Set correl_da = CAST('" & var_cod & "' AS INT) + 1 Where dgral_codigo= '" & dtc_codigo1.Text & "' "
'     End If
'     Ado_datos.Recordset!NombreCta = TxtConcepto.Text
'     Ado_datos.Recordset!fecha_registro = Date     ' no cambia
'     Ado_datos.Recordset!usr_codigo = glusuario    ' no cambia
'     Ado_datos.Recordset.Update     'Batch 'adAffectAll
    If VAR_SW = "MOD" Then
    
        If rs_datos!nivel < 5 Then
            db.Execute " UPDATE CC_plan_cuentas set NombreCta='" & TxtConcepto.Text & "',Usr_codigo='" & glusuario & "', Fecha_registro='" & Date & "' WHERE CC_plan_cuentas.correl='" & Ado_datos.Recordset!CORREL & "' "
            db.Execute " UPDATE cc_plan_nivel1 set NombreCta='" & TxtConcepto.Text & "',Usr_codigo='" & glusuario & "', Fecha_registro='" & Date & "' WHERE cc_plan_nivel1.correl='" & Ado_datos.Recordset!CORREL & "' "
            db.Execute " UPDATE cc_plan_nivel2 set NombreCta='" & TxtConcepto.Text & "',Usr_codigo='" & glusuario & "', Fecha_registro='" & Date & "' WHERE cc_plan_nivel2.correl='" & Ado_datos.Recordset!CORREL & "' "
            db.Execute " UPDATE cc_plan_nivel3 set NombreCta='" & TxtConcepto.Text & "',Usr_codigo='" & glusuario & "', Fecha_registro='" & Date & "' WHERE cc_plan_nivel3.correl='" & Ado_datos.Recordset!CORREL & "' "
            db.Execute " UPDATE cc_plan_nivel4 set NombreCta='" & TxtConcepto.Text & "',Usr_codigo='" & glusuario & "', Fecha_registro='" & Date & "' WHERE cc_plan_nivel4.correl='" & Ado_datos.Recordset!CORREL & "' "
            
            
         End If
         If rs_datos!nivel = 5 Then
         
            db.Execute " UPDATE CC_plan_cuentas set NombreCta='" & TxtConcepto.Text & "', Aux1='" & dtc_codigo5.Text & "', Aux2='" & dtc_codigo6.Text & "', Aux3='" & dtc_codigo7.Text & "', Usr_codigo='" & glusuario & "', Fecha_registro='" & Date & "' WHERE CC_plan_cuentas.correl='" & Ado_datos.Recordset!CORREL & "' "
            
            'db.Execute " UPDATE cc_plan_nivel1 set cc_plan_nivel1.correl=CC_plan_cuentas.correl FROM cc_plan_nivel1 INNER JOIN CC_plan_cuentas ON cc_plan_nivel1.Cuenta=CC_plan_cuentas.Cuenta WHERE cc_plan_nivel1.Cuenta='" & VAR_CTA & "' "
         End If
         
     
     End If
     Call OptFilGral2_Click
     rs_datos.Update
     rs_datos.MoveLast
     mbDataChanged = False
      
      Fra_ABM.Enabled = True
      FraNavega.Enabled = True
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      Fra_ABM2.Visible = False
      dg_datos.Enabled = True
      Fra_Det1.Visible = False
      Fra_Det2.Visible = False
      Fra_Det3.Visible = False
      Fra_Det4.Visible = False
      
      'txt_codigo.Enabled = True
     'dtc_cuenta.Enabled = True
    ' dtc_desc1.Enabled = True
      VAR_SW = ""
 End If

  Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub valida_campos()
'habilitar codigo cuando se transcribe
  If dtc_cuenta.Text = "" Then
    MsgBox "Debe registrar el " + lbl_codigo.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If TxtConcepto.Text = "" Then
    MsgBox "Debe registrar la " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If VAR_SW = "ADD" And Val(dtc_cuenta.Text) > "1" Then
'    If dtc_codigo1.Text = "" Then
'      MsgBox "Debe registrar: " + lbl_enlace1.Caption, vbCritical + vbExclamation, "Validación de datos"
'      VAR_VAL = "ERR"
'      Exit Sub
'    End If
'    If dtc_codigo2.Text = "" And Val(dtc_cuenta.Text) > "2" Then
'      MsgBox "Debe registrar: " + lbl_enlace2.Caption, vbCritical + vbExclamation, "Validación de datos"
'      VAR_VAL = "ERR"
'      Exit Sub
'    End If
'    If dtc_codigo3.Text = "" And Val(dtc_cuenta.Text) > "3" Then
'      MsgBox "Debe registrar: " + lbl_enlace3.Caption, vbCritical + vbExclamation, "Validación de datos"
'      VAR_VAL = "ERR"
'      Exit Sub
'    End If
'    If dtc_codigo4.Text = "" And Val(dtc_cuenta.Text) > "4" Then
'      MsgBox "Debe registrar: " + lbl_enlace4.Caption, vbCritical + vbExclamation, "Validación de datos"
'      VAR_VAL = "ERR"
'      Exit Sub
'    End If
'  End If
  If dtc_cuenta.Text = "5" Then
      If dtc_codigo5.Text = "" Then
        MsgBox "Debe registrar: " + lbl_enlace5.Caption, vbCritical + vbExclamation, "Validación de datos"
        VAR_VAL = "ERR"
        Exit Sub
      End If
    '   If dtc_codigo6.Text = "" Then
    '    MsgBox "Debe registrar: " + lbl_enlace6.Caption, vbCritical + vbExclamation, "Validación de datos"
    '    VAR_VAL = "ERR"
    '    Exit Sub
    '  End If
    '   If dtc_codigo7.Text = "" Then
    '    MsgBox "Debe registrar: " + lbl_enlace7.Caption, vbCritical + vbExclamation, "Validación de datos"
    '    VAR_VAL = "ERR"
    '    Exit Sub
    '  End If
   End If
End Sub

Private Sub BtnImprimir_Click()
  Dim iResult As Integer
  CR01.WindowShowPrintSetupBtn = True
  CR01.WindowShowRefreshBtn = True
  CR01.ReportFileName = App.Path & "\REPORTES\contabilidad\cr_plan_cuentas.rpt"
  iResult = CR01.PrintReport
  If iResult <> 0 Then
      MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
   Exit Sub
  CR01.WindowState = crptMaximized
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
   If rs_datos!estado_codigo = "REG" Then
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    FraNavega.Enabled = False
    dg_datos.Enabled = False
    VAR_SW = "MOD"
         Fra_ABM1.Enabled = False
         Fra_ABM3.Enabled = True
         TxtConcepto.Enabled = True
         Fra_ABM2.Visible = False
        dtc_cuenta.Text = rs_datos!nivel
           Txt_campo2.Caption = var_cod
           txt_campo3.Caption = VAR_COD1
           txt_campo4.Caption = VAR_COD2
        If rs_datos!nivel < 5 Then
            Fra_Aux.Visible = False
        Else
             Fra_Aux.Visible = True
             dtc_desc5.Visible = True
             dtc_desc6.Visible = True
             dtc_desc7.Visible = True
             dtc_codigo5.Visible = True
             dtc_codigo6.Visible = True
             dtc_codigo7.Visible = True
            'Call ABRIR_
            Call ABRIR_AUX2
            Call ABRIR_AUX3
        
        End If
  Else
        MsgBox "No se puede MODIFICAR un registro APROBADO o Errado ...", vbExclamation, "Validación de Registro"
  End If
'  lblStatus.Caption = "Modificar registro"
        'TxtConcepto.SetFocus
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

Private Sub BtnSalir1_Click()
 Unload Me
End Sub

Private Sub BtnSalir2_Click()
 Unload Me
End Sub

Private Sub BtnSalir3_Click()
 Unload Me
End Sub

Private Sub BtnSalir4_Click()
 Unload Me
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

'Function LlenarMedida()
'Set rs = New Recordse
' rs.Open "SELECT * FROM CC_plan_cuentas", db, adOpenStatic, adLockOptimistic
' DataCombo1.BoundColumn = "Cuenta"
' DataCombo1.ListField = "NombreCta"
' Set DataCombo1.RowSource = rs
'
'End Function


'
'Private Sub dtc_aux1_Click(Area As Integer)
''    dtc_desc1.BoundText = dtc_Aux1.BoundText
''    dtc_codigo1.BoundText = dtc_Aux1.BoundText
'End Sub
'
'Private Sub dtc_aux2_Click(Area As Integer)
''    dtc_desc2.BoundText = dtc_Aux2.BoundText
''    dtc_codigo2.BoundText = dtc_Aux2.BoundText
'End Sub
'
'Private Sub dtc_aux3_Click(Area As Integer)
''    dtc_desc3.BoundText = dtc_Aux3.BoundText
''    dtc_codigo3.BoundText = dtc_Aux3.BoundText
'End Sub
'
'Private Sub dtc_aux4_Click(Area As Integer)
''    dtc_desc4.BoundText = dtc_Aux4.BoundText
''    dtc_codigo4.BoundText = dtc_Aux4.BoundText
'End Sub
'
'Private Sub dtc_codigo1_Click(Area As Integer)
''    dtc_desc1.BoundText = dtc_codigo1.BoundText
''    dtc_Aux1.BoundText = dtc_codigo1.BoundText
'End Sub
'
'Private Sub dtc_codigo2_Click(Area As Integer)
''    dtc_desc2.BoundText = dtc_codigo2.BoundText
''    dtc_Aux2.BoundText = dtc_codigo2.BoundText
'End Sub
'
'Private Sub dtc_codigo3_Click(Area As Integer)
''    dtc_desc3.BoundText = dtc_codigo3.BoundText
''    dtc_Aux3.BoundText = dtc_codigo3.BoundText
'End Sub
'
'Private Sub dtc_codigo4_Click(Area As Integer)
''    dtc_desc4.BoundText = dtc_codigo4.BoundText
''    dtc_Aux4.BoundText = dtc_codigo4.BoundText
'End Sub
'
'Private Sub dtc_codigo5_Click(Area As Integer)
''  dtc_desc5.BoundText = dtc_codigo5.BoundText
''  'dtc_aux5.BoundText = dtc_codigo5.BoundText
'End Sub
'
'Private Sub dtc_codigo6_Click(Area As Integer)
''  dtc_desc6.BoundText = dtc_codigo6.BoundText
''    'dtc_aux6.BoundText = dtc_codigo6.BoundText
'End Sub
'
'Private Sub dtc_codigo7_Click(Area As Integer)
''     dtc_desc7.BoundText = dtc_codigo7.BoundText
''    'dtc_aux7.BoundText = dtc_codigo7.BoundText
'End Sub

Private Sub dtc_cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 0 Then
    KeyAscii = 0
    Else
    Exit Sub
    End If
End Sub

Private Sub dtc_cuenta_LostFocus()
    Call Abrir_Aux
End Sub

'Private Sub dtc_codigo5_Click(Area As Integer)
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'End Sub

'Private Sub dtc_codigo6_Click(Area As Integer)
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'End Sub
'
'Private Sub dtc_codigo7_Click(Area As Integer)
'    dtc_desc1.BoundText = dtc_codigo7.BoundText
'    dtc_codigo1.BoundText = dtc_codigo7.BoundText
'    dtc_codigo10.BoundText = dtc_codigo7.BoundText
'End Sub
'
'Private Sub dtc_codigo8_Click(Area As Integer)
'    dtc_desc2.BoundText = dtc_codigo8.BoundText
'    dtc_codigo2.BoundText = dtc_codigo8.BoundText
'    dtc_codigo11.BoundText = dtc_codigo8.BoundText
'End Sub
'
'Private Sub dtc_codigo9_Click(Area As Integer)
'    dtc_desc3.BoundText = dtc_codigo9.BoundText
'    dtc_codigo3.BoundText = dtc_codigo9.BoundText
'    dtc_codigo12.BoundText = dtc_codigo9.BoundText
'End Sub

'Private Sub dtc_desc1_Click(Area As Integer)
''    dtc_codigo1.BoundText = dtc_desc1.BoundText
''    dtc_Aux1.BoundText = dtc_desc1.BoundText
''    If dtc_cuenta.Text > 2 Then
''        Call pnivel2(dtc_codigo1.Text)
''        dtc_desc2.Enabled = True
''    End If
'
'End Sub

Private Sub pnivel2(codigo1 As String)
   Dim strConsultaF As String
     'rs_datos2.Open "Select * from cc_plan_nivel2 order by Cuenta ", db, adOpenStatic
   strConsultaF = "select * from cc_plan_nivel2 where left(Cuenta,1)= '" & Left(codigo1, 1) & "'"
'   Set dtc_codigo2.RowSource = Nothing
'   Set dtc_codigo2.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo2.ReFill
'   dtc_codigo2.BoundText = Empty
   
'   Set dtc_desc2.RowSource = Nothing
'   Set dtc_desc2.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc2.ReFill
'   dtc_desc2.BoundText = Empty

End Sub

'Private Sub dtc_desc2_Click(Area As Integer)
''  dtc_codigo2.BoundText = dtc_desc2.BoundText
''    dtc_Aux2.BoundText = dtc_desc2.BoundText
''    If dtc_cuenta.Text > 3 Then
''        Call pnivel3(dtc_codigo2.Text)
''        dtc_desc3.Enabled = True
''    End If
'End Sub

Private Sub pnivel3(codigo2 As String)
   Dim strConsultaF As String
     'rs_datos2.Open "Select * from cc_plan_nivel2 order by Cuenta ", db, adOpenStatic
   strConsultaF = "select * from cc_plan_nivel3 where left(Cuenta,2)= '" & Left(codigo2, 2) & "'"
'   Set dtc_codigo3.RowSource = Nothing
'   Set dtc_codigo3.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo3.ReFill
'   dtc_codigo3.BoundText = Empty
'
'   Set dtc_desc3.RowSource = Nothing
'   Set dtc_desc3.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc3.ReFill
'   dtc_desc3.BoundText = Empty

End Sub

'Private Sub dtc_desc3_Click(Area As Integer)
'' dtc_codigo3.BoundText = dtc_desc3.BoundText
''    dtc_Aux3.BoundText = dtc_desc3.BoundText
''    If dtc_cuenta.Text > 4 Then
''        Call pnivel4(dtc_codigo3.Text)
''        dtc_desc4.Enabled = True
''    End If
'End Sub
Private Sub pnivel4(codigo3 As String)
   Dim strConsultaF As String
     'rs_datos2.Open "Select * from cc_plan_nivel2 order by Cuenta ", db, adOpenStatic
   strConsultaF = "select * from cc_plan_nivel4 where left(Cuenta,3)= '" & Left(codigo3, 3) & "'"
'   Set dtc_codigo4.RowSource = Nothing
'   Set dtc_codigo4.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo4.ReFill
'   dtc_codigo4.BoundText = Empty
'
'   Set dtc_desc4.RowSource = Nothing
'   Set dtc_desc4.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc4.ReFill
'   dtc_desc4.BoundText = Empty

End Sub

'Private Sub dtc_desc4_Click(Area As Integer)
''    dtc_codigo4.BoundText = dtc_desc4.BoundText
''    dtc_Aux4.BoundText = dtc_desc4.BoundText
''    If dtc_cuenta.Text > 5 Then
''        Call pnivel5(dtc_codigo4.Text)
''        dtc_desc3.Enabled = True
''
''    End If
'End Sub

Private Sub pnivel5(codigo4 As String)
   Dim strConsultaF As String
     'rs_datos2.Open "Select * from cc_plan_nivel2 order by Cuenta ", db, adOpenStatic
   strConsultaF = "select * from cc_tipo_auxiliar where left(Cuenta,4)= '" & Left(codigo4, 4) & "'"
   Set dtc_codigo5.RowSource = Nothing
   Set dtc_codigo5.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo5.ReFill
   dtc_codigo5.BoundText = Empty
   
   Set dtc_desc5.RowSource = Nothing
   Set dtc_desc5.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc5.ReFill
   dtc_desc5.BoundText = Empty

End Sub

'Private Sub dtc_desc5_Click(Area As Integer)
'  dtc_codigo5.BoundText = dtc_desc5.BoundText
'  'dtc_codigo5.BoundText = dtc_desc5.BoundText
''   If dtc_cuenta.Text > 6 Then
''        Call pnivel6(dtc_codigo5.Text)
''        dtc_desc5.Enabled = True
''       dtc_desc5.Enabled = True
''       Fra_Aux.Enabled = True
''    End If
'End Sub

Private Sub pnivel6(codigo5 As String)
   Dim strConsultaF As String
   'rs_datos2.Open "Select * from cc_plan_nivel2 order by Cuenta ", db, adOpenStatic
   strConsultaF = "select * from cc_tipo_auxiliar where left(aux,5)= '" & Left(codigo5, 5) & "'"
   Set dtc_codigo6.RowSource = Nothing
   Set dtc_codigo6.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo6.ReFill
   dtc_codigo6.BoundText = Empty
   
   Set dtc_desc6.RowSource = Nothing
   Set dtc_desc6.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc6.ReFill
   dtc_desc6.BoundText = Empty

End Sub

Private Sub dtc_desc5_Click(Area As Integer)
 dtc_codigo5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
 dtc_codigo6.BoundText = dtc_desc6.BoundText
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
 dtc_codigo7.BoundText = dtc_desc7.BoundText
End Sub

'Private Sub dtc_desc6_Click(Area As Integer)
'dtc_codigo6.BoundText = dtc_desc6.BoundText
'  'dtc_codigo5.BoundText = dtc_desc5.BoundText
''   If dtc_cuenta.Text > 6 Then
''        Call pnivel6(dtc_codigo6.Text)
''        dtc_desc6.Enabled = True
''       dtc_desc6.Enabled = True
''       Fra_Aux.Enabled = True
''    End If
'End Sub

'Private Sub dtc_desc7_Click(Area As Integer)
'dtc_codigo7.BoundText = dtc_desc7.BoundText
'  'dtc_codigo5.BoundText = dtc_desc5.BoundText
''   If dtc_cuenta.Text > 6 Then
''        Call pnivel7(dtc_codigo7.Text)
''        dtc_desc7.Enabled = True
''       dtc_desc7.Enabled = True
''       Fra_Aux.Enabled = True
''    End If
'End Sub
'
'Private Sub dtc_desc7_Click(Area As Integer)
'  dtc_codigo7.BoundText = dtc_desc7.BoundText
'    dtc_aux7.BoundText = dtc_desc7.BoundText
'End Sub

'Private Sub dtc_desc4_Click(Area As Integer)
'    dtc_codigo4.BoundText = dtc_desc4.BoundText
'End Sub

'Private Sub dtc_desc5_Click(Area As Integer)
'    dtc_codigo5.BoundText = dtc_desc5.BoundText
'End Sub

'Private Sub dtc_desc6_Click(Area As Integer)
'    dtc_codigo6.BoundText = dtc_desc6.BoundText
'End Sub

Private Sub Form_Load()
'    Call ABRIR_TABLAS_AUX
    Call OptFilGral2_Click

'   Call ABRIR_TABLA
'   txt_codigo.Enabled = True
    VAR_SW = ""
    mbDataChanged = False
    Fra_ABM.Enabled = False
    dg_datos.Enabled = True
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
  '  txt_Tcta.Visible = False
'    txt_Tscta1.Visible = False
'    txt_Tscta2.Visible = False
   ' txt_desc1.Visible = False
    
'    txt_Tcta2.Visible = False
'    txt_Tscta12.Visible = False
'    txt_Tscta22.Visible = False
'    txt_desc2.Visible = False
	Call SeguridadSet(Me)
End Sub

Private Sub OptFilGral2_Click()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from CC_Plan_Cuentas "      ' todos
  'queryinicial = "select  da_codigo, da_descripcion, dgral_codigo, proceso_codigo, estado_codigo, fecha_registro, usr_codigo, correl_unidad as correl from gc_direccion_administrativa  "
  'queryinicial = "gp_listar_gc_direccion_general "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = rs_datos
End Sub

Private Sub OptFilGral1_Click()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from CC_Plan_Cuentas where estado_codigo= 'REG' "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = rs_datos
End Sub

Private Sub OptFilGral3_Click()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from CC_Plan_Cuentas  where estado_codigo= 'APR' "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = rs_datos
End Sub

Private Sub OptFilGral4_Click()
'  Set rs_datos = New Recordset
'  If rs_datos.State = 1 Then rs_datos.Close
'  queryinicial = "select  * from CC_Plan_Cuentas  where mov= 'D' "
'  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'  Set Ado_datos = rs_datos.DataSource
'  Set dg_datos.DataSource = rs_datos
End Sub

Private Sub ABRIR_TABLAS_AUX()
'    Set rs_datos1 = New ADODB.Recordset
'    If rs_datos1.State = 1 Then rs_datos1.Close
'    'rs_datos1.Open "Select * from CC_Plan_Cuentas WHERE SubCta1 = '00' AND SubCta2 = '00' order by Cuenta ", db, adOpenStatic
'    rs_datos1.Open "Select * from CC_Plan_Cuentas WHERE MOV = 'T' order by Cuenta ", db, adOpenStatic
'    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
'
'    Set rs_datos2 = New ADODB.Recordset
'    If rs_datos2.State = 1 Then rs_datos2.Close
'    'rs_datos2.Open "Select * from CC_Plan_Cuentas WHERE SubCta1 <> '00' AND SubCta2 = '00' order by Cuenta ", db, adOpenStatic
'    rs_datos2.Open "Select * from CC_Plan_Cuentas WHERE MOV = 'S' order by Cuenta ", db, adOpenStatic
'    Set Ado_datos2.Recordset = rs_datos2
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
'
'    Set rs_datos3 = New ADODB.Recordset
'    If rs_datos3.State = 1 Then rs_datos3.Close
'    'rs_datos3.Open "Select * from CC_Plan_Cuentas WHERE SubCta1 <> '00' AND SubCta2 <> '00' order by Cuenta ", db, adOpenStatic
'    rs_datos3.Open "Select * from CC_Plan_Cuentas WHERE MOV = 'D' order by Cuenta ", db, adOpenStatic
'    Set Ado_datos3.Recordset = rs_datos3
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
'
'     Set rs_datos1 = New ADODB.Recordset
'    If rs_datos1.State = 1 Then rs_datos1.Close
'    'rs_datos1.Open "Select * from CC_Plan_Cuentas WHERE SubCta1 = '00' AND SubCta2 = '00' order by Cuenta ", db, adOpenStatic
''    rs_datos1.Open "Select * from CC_Plan_Cuentas WHERE MOV = 'T' order by Cuenta ", db, adOpenStatic
'    rs_datos1.Open "Select * from cc_plan_nivel1 WHERE MOV = 'T' order by Cuenta ", db, adOpenStatic
'    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
'
'    Set rs_datos2 = New ADODB.Recordset
'    If rs_datos2.State = 1 Then rs_datos2.Close
'    'rs_datos2.Open "Select * from CC_Plan_Cuentas WHERE SubCta1 <> '00' AND SubCta2 = '00' order by Cuenta ", db, adOpenStatic
''    rs_datos2.Open "Select * from CC_Plan_Cuentas WHERE MOV = 'S' order by Cuenta ", db, adOpenStatic
'     rs_datos2.Open "Select * from cc_plan_nivel2  order by Cuenta ", db, adOpenStatic
'    Set Ado_datos2.Recordset = rs_datos2
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
'
'    Set rs_datos3 = New ADODB.Recordset
'    If rs_datos3.State = 1 Then rs_datos3.Close
'    'rs_datos3.Open "Select * from CC_Plan_Cuentas WHERE SubCta1 <> '00' AND SubCta2 <> '00' order by Cuenta ", db, adOpenStatic
''   rs_datos3.Open "Select * from CC_Plan_Cuentas WHERE MOV = 'D' order by Cuenta ", db, adOpenStatic
'    rs_datos3.Open "Select * from cc_plan_nivel3 order by Cuenta ", db, adOpenStatic
'    Set Ado_datos3.Recordset = rs_datos3
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
'
'
'    Set rs_datos4 = New ADODB.Recordset
'    If rs_datos4.State = 1 Then rs_datos4.Close
'    'rs_datos4.Open "Select * from CC_Plan_Cuentas WHERE SubCta1 <> '00' AND SubCta2 <> '00' order by Cuenta ", db, adOpenStatic
''   rs_datos4.Open "Select * from CC_Plan_Cuentas WHERE MOV = 'D' order by Cuenta ", db, adOpenStatic
'    rs_datos4.Open "Select * from cc_plan_nivel4  order by Cuenta ", db, adOpenStatic
'    Set Ado_datos4.Recordset = rs_datos4
'    dtc_desc4.BoundText = dtc_codigo4.BoundText
    

'    Set rs_datos4 = New ADODB.Recordset
'    If rs_datos4.State = 1 Then rs_datos4.Close
'    rs_datos4.Open "Select * from cc_tipo_auxiliar order by aux ", db, adOpenStatic
'    Set Ado_datos4.Recordset = rs_datos4
'    'dtc_desc4.BoundText = dtc_codigo4.BoundText
'
'    Set rs_datos5 = New ADODB.Recordset
'    If rs_datos5.State = 1 Then rs_datos5.Close
'    rs_datos5.Open "Select * from cc_tipo_auxiliar order by aux ", db, adOpenStatic
'    Set Ado_datos5.Recordset = rs_datos5
'    'dtc_desc5.BoundText = dtc_codigo5.BoundText
'
'    Set rs_datos6 = New ADODB.Recordset
'    If rs_datos6.State = 1 Then rs_datos6.Close
'    rs_datos6.Open "Select * from cc_tipo_auxiliar order by aux ", db, adOpenStatic
'    Set Ado_datos6.Recordset = rs_datos6
'    'dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub ABRIR_AUX1()
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from cc_tipo_auxiliar order by aux ", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    
End Sub

Private Sub ABRIR_AUX2()
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from cc_tipo_auxiliar order by aux ", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub ABRIR_AUX3()
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from cc_tipo_auxiliar order by aux ", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
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

Private Sub ABRIR_NIVEL1()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'WHERE correl = '" & rs_datos!CORREL & "'
    rs_datos1.Open "Select * from cc_plan_nivel1   order by Cuenta ", db, adOpenKeyset, adLockBatchOptimistic
    'rs_datos1.Open "Select * from cc_plan_cuentas where Nivel='1'   order by NombreCta ", db, adOpenKeyset, adLockBatchOptimistic
    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_Aux1.BoundText
'    dtc_codigo1.BoundText = dtc_Aux1.BoundText
End Sub

Private Sub ABRIR_NIVEL2()
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
   
'   WHERE cuenta = '" & lef(Ado_datos1.Recordset!cuenta,1) & "'
     rs_datos2.Open "Select * from cc_plan_nivel2  WHERE left(cuenta,1) = '" & Left(Ado_datos1.Recordset!cuenta, 1) & "' order by Cuenta ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
'    dtc_desc2.BoundText = dtc_Aux2.BoundText
'    dtc_codigo2.BoundText = dtc_Aux2.BoundText

End Sub

Private Sub ABRIR_NIVEL3()
Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
     rs_datos3.Open "Select * from cc_plan_nivel3 WHERE left(cuenta,2) = '" & Left(Ado_datos2.Recordset!cuenta, 2) & "' order by Cuenta ", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
'     dtc_desc3.BoundText = dtc_Aux3.BoundText
'    dtc_codigo3.BoundText = dtc_Aux3.BoundText
End Sub

Private Sub ABRIR_NIVEL4()
Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
     rs_datos4.Open "Select * from cc_plan_nivel4 WHERE left(cuenta,3) = '" & Left(Ado_datos3.Recordset!cuenta, 3) & "' order by Cuenta ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
'     dtc_desc4.BoundText = dtc_Aux4.BoundText
'    dtc_codigo4.BoundText = dtc_Aux4.BoundText
    
End Sub

Private Sub Abrir_Aux()
    Fra_ABM2.Visible = True
    If dtc_cuenta.Text = 1 Then
       
        dtc_codigo1.Visible = False
        dtc_desc1.Visible = False
        dtc_codigo2.Visible = False
        dtc_desc2.Visible = False
        dtc_codigo3.Visible = False
        dtc_desc3.Visible = False
        dtc_codigo4.Visible = False
        dtc_desc4.Visible = False
        
        Fra_Aux.Visible = False
        Call ABRIR_NIVEL1
    End If
    
    If dtc_cuenta.Text = 2 Then
        Fra_Det1.Visible = True
        
        dtc_codigo1.Visible = True
        dtc_desc1.Visible = True
        dtc_codigo2.Visible = False
        dtc_desc2.Visible = False
        dtc_codigo3.Visible = False
        dtc_desc3.Visible = False
        dtc_codigo4.Visible = False
        dtc_desc4.Visible = False

        Fra_Aux.Visible = False
        Call ABRIR_NIVEL1
        Call ABRIR_NIVEL2
    End If

    If dtc_cuenta.Text = 3 Then
        Fra_Det1.Visible = True

        dtc_codigo1.Visible = True
        dtc_desc1.Visible = True
        dtc_codigo2.Visible = True
        dtc_desc2.Visible = True
        dtc_codigo3.Visible = False
        dtc_desc3.Visible = False
        dtc_codigo4.Visible = False
        dtc_desc4.Visible = False

        Fra_Aux.Visible = False
        Call ABRIR_NIVEL1
        Call ABRIR_NIVEL2
        Call ABRIR_NIVEL3
    End If
'
    If dtc_cuenta.Text = 4 Then
        Fra_Det1.Visible = True

        dtc_codigo1.Visible = True
        dtc_desc1.Visible = True
        dtc_codigo2.Visible = True
        dtc_desc2.Visible = True
        dtc_codigo3.Visible = True
        dtc_desc3.Visible = True
        dtc_codigo4.Visible = False
        dtc_desc4.Visible = False

        Fra_Aux.Visible = False
        Call ABRIR_NIVEL1
        Call ABRIR_NIVEL2
        Call ABRIR_NIVEL3
        Call ABRIR_NIVEL4
    End If

    If dtc_cuenta.Text = 5 Then
        Fra_Det1.Visible = True
    
        dtc_codigo1.Visible = True
        dtc_desc1.Visible = True
        dtc_codigo2.Visible = True
        dtc_desc2.Visible = True
        dtc_codigo3.Visible = True
        dtc_desc3.Visible = True
        dtc_codigo4.Visible = True
        dtc_desc4.Visible = True

        Fra_Aux.Visible = True
        Call ABRIR_NIVEL1
        Call ABRIR_NIVEL2
        Call ABRIR_NIVEL3
        Call ABRIR_NIVEL4
        Call ABRIR_AUX1
        Call ABRIR_AUX2
        Call ABRIR_AUX3

    End If
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If Ado_datos.Recordset.RecordCount > 0 Then
     ' Ado_datos.Caption = rs_datos.AbsolutePosition & " / " & rs_datos.RecordCount
     If VAR_SW = "" Then
        Fra_ABM2.Visible = False
     Else
        Fra_ABM2.Visible = True
     End If
              
'  Call Abrir_Aux
     
'     If Ado_datos.Recordset!mov = "T" Then
'       ' lbl_observacion.Caption = "TITULO"
'        Fra_Aux.Visible = False
''        dtc_codigo1.Visible = True
'        'dtc_codigo7.Visible = True
'       ' dtc_codigo10.Visible = True
''        dtc_desc1.Visible = True
'
''        dtc_codigo2.Visible = True
'       ' dtc_codigo8.Visible = False
'        'dtc_codigo11.Visible = False
''        dtc_desc2.Visible = True
'
''        dtc_codigo3.Visible = True
''        'dtc_codigo9.Visible = False
''       ' dtc_codigo12.Visible = False
''        dtc_desc3.Visible = True
''
''        dtc_codigo4.Visible = True
''        'dtc_codigo9.Visible = False
''       ' dtc_codigo12.Visible = False
''        dtc_desc4.Visible = True
'
'
'        ' txt_Tcta.Visible = False
'        ' txt_Tscta1.Visible = False
'        ' txt_Tscta2.Visible = False
'        ' txt_desc1.Visible = False
'
'        'txt_Tcta2.Visible = False
'        ' txt_Tscta12.Visible = False
'        'txt_Tscta22.Visible = False
'         'txt_desc2.Visible = False
'     Else
'        If Ado_datos.Recordset!mov = "S" Then
'           Fra_Aux.Visible = False
'           'lbl_observacion.Caption = "SUB TITULO"
''           dtc_codigo1.Visible = False
'            'dtc_codigo7.Visible = False
'            'dtc_codigo10.Visible = False
''            dtc_desc1.Visible = False
'
''            dtc_codigo2.Visible = True
''           ' dtc_codigo8.Visible = True
''           ' dtc_codigo11.Visible = True
''            dtc_desc2.Visible = True
''
''            dtc_codigo3.Visible = True
''            'dtc_codigo9.Visible = False
''            'dtc_codigo12.Visible = False
''            dtc_desc3.Visible = True
''
''            dtc_codigo4.Visible = True
''            'dtc_codigo9.Visible = False
''            'dtc_codigo12.Visible = False
''            dtc_desc4.Visible = True
'
'        var_cod = Ado_datos.Recordset!cuenta
''            txt_Tcta.Visible = True
''            txt_Tscta1.Visible = True
''            txt_Tscta2.Visible = True
''            txt_desc1.Visible = True
'
''            If txt_Tcta.Text <> "" Then
'              If rs_aux2.State = 1 Then rs_aux2.Close
'              'rs_aux2.Open "select  * from CC_Plan_Cuentas where mov= 'T' and cuenta = '" & txt_Tcta & "' and SubCta1 = '" & txt_Tscta1 & "' and SubCta2 = '" & txt_Tscta2 & "' ", db, adOpenKeyset, adLockOptimistic
'              'rs_aux2.Open "select  * from CC_Plan_Cuentas where mov= 'T' and cuenta = '" & dtc_codigo1 & "' and SubCta1 = '00' and SubCta2 = '00' ", db, adOpenKeyset, adLockOptimistic
'              rs_aux2.Open "select  * from CC_Plan_Cuentas where mov= 'T' and cuenta = '" & var_cod & "'  ", db, adOpenKeyset, adLockOptimistic
'              If rs_aux2.RecordCount > 0 Then
''                txt_Tcta.Text = rs_aux2("Cuenta")
''                txt_Tscta1.Text = rs_aux2("SubCta1")
''                txt_Tscta2.Text = rs_aux2("SubCta2")
''                txt_desc1.Text = rs_aux2("NombreCta")
'              End If
''            Else
''                txt_desc1.Text = dtc_desc1.Text
''            End If
'
'
'        '    If DtCCta_codigo.Text <> "01" Then
'        '      If rstdestino.State = 1 Then rstdestino.Close
'        '      rstFc_cuenta_bancaria.Find " cta_codigo = '" & DtCCta_codigo & "'", , adSearchForward, 1
'        '      If Not rstFc_cuenta_bancaria.EOF Then
'        '        fte_codigo1 = rstFc_cuenta_bancaria("fte_codigo")
'        '      Else
'        '      End If
'        '    Else
'        '        fte_codigo1 = Me.DtCFte_codigo.Text
'        '    End If
'        Else
'           Fra_Aux.Visible = True
'
'           'lbl_observacion.Caption = "DETALLE"
''           dtc_codigo1.Visible = False
'''            dtc_codigo7.Visible = False
'''            dtc_codigo10.Visible = False
''            dtc_desc1.Visible = False
''
''            dtc_codigo2.Visible = True
''            'dtc_codigo8.Visible = False
''            'dtc_codigo11.Visible = False
''            dtc_desc2.Visible = False
''
''            dtc_codigo3.Visible = True
''            'dtc_codigo9.Visible = True
''           ' dtc_codigo12.Visible = True
''            dtc_desc3.Visible = True
''
''            dtc_codigo4.Visible = True
''            'dtc_codigo9.Visible = True
''           ' dtc_codigo12.Visible = True
''            dtc_desc4.Visible = True
'
''            txt_Tcta.Visible = True
''            txt_Tscta1.Visible = True
''            txt_Tscta2.Visible = True
''            txt_desc1.Visible = True
'
''            txt_Tcta2.Visible = True
''            txt_Tscta12.Visible = True
''            txt_Tscta22.Visible = True
''            txt_desc2.Visible = True
'
'            var_cod = Ado_datos.Recordset!cuenta
'            VAR_COD1 = Ado_datos.Recordset!subcta1
'            VAR_COD2 = Ado_datos.Recordset!subcta2
'
''            If txt_Tcta2.Text <> "" Then
'              If rs_aux1.State = 1 Then rs_aux1.Close
'              'rs_aux1.Open "select  * from CC_Plan_Cuentas where mov= 'S' and cuenta = '" & dtc_codigo3 & "' and SubCta1 = '" & dtc_codigo9 & "' and SubCta2 = '00' ", db, adOpenKeyset, adLockOptimistic
'              rs_aux1.Open "select  * from CC_Plan_Cuentas where mov= 'S' and cuenta = '" & var_cod & "' and subcta1 = '" & VAR_COD1 & "'  ", db, adOpenKeyset, adLockOptimistic
'              If rs_aux1.RecordCount > 0 Then
''                txt_Tcta2.Text = rs_aux1("Cuenta")
''                txt_Tscta12.Text = rs_aux1("SubCta1")
''                txt_Tscta22.Text = rs_aux1("SubCta2")
''                txt_desc2.Text = rs_aux1("NombreCta")
'              End If
'
'              If rs_aux2.State = 1 Then rs_aux2.Close
'              'rs_aux2.Open "select  * from CC_Plan_Cuentas where mov= 'T' and cuenta = '" & txt_Tcta & "' and SubCta1 = '00' and SubCta2 = '00' ", db, adOpenKeyset, adLockOptimistic
'              rs_aux2.Open "select  * from CC_Plan_Cuentas where mov= 'T' and cuenta = '" & var_cod & "' ", db, adOpenKeyset, adLockOptimistic
'              If rs_aux2.RecordCount > 0 Then
''                  txt_Tcta.Text = rs_aux2("Cuenta")
''                  txt_Tscta1.Text = rs_aux2("SubCta1")
''                  txt_Tscta2.Text = rs_aux2("SubCta2")
''                  txt_desc1.Text = rs_aux2("NombreCta")
'              End If
            If Ado_datos.Recordset!aux1 = "00" Then
'                Chkaux1.Value = 0
                dtc_codigo5.Visible = False
                dtc_desc5.Visible = False
                'dtc_desc4.Visible = False
            Else
'                Chkaux1.Value = 1
                Call ABRIR_AUX1
                dtc_codigo5.Visible = True
                dtc_desc5.Visible = True
            End If
            If Ado_datos.Recordset!AUX2 = "00" Then
'                Chkaux2.Value = 0
                dtc_codigo6.Visible = False
                dtc_desc6.Visible = False
            Else
'                Chkaux2.Value = 1
                Call ABRIR_AUX2
                dtc_codigo6.Visible = True
                dtc_desc6.Visible = True
            End If
            If Ado_datos.Recordset!aux3 = "00" Then
'                Chkaux3.Value = 0
                 dtc_codigo7.Visible = False
                dtc_desc7.Visible = False
            Else
'                Chkaux3.Value = 1
                Call ABRIR_AUX3
                dtc_codigo7.Visible = True
                dtc_desc7.Visible = True
            End If
'        End If
'     End If
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
    Call OptFilGral2_Click
    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    rs_datos.AddNew
    'lblStatus.Caption = "Agregar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    Fra_ABM2.Visible = True
    FraNavega.Enabled = False
'   dg_datos.Enabled = False
    VAR_SW = "ADD"
    
'   txt_codigo.Enabled = False

    dtc_cuenta.SetFocus
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

Private Function ExisteReg(cuenta2 As String, scuenta1 As String, scuenta2 As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE dgral_codigo = '" & Unidad & "'"
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM co_diario WHERE D_Cuenta = '" & cuenta2 & "' and D_Subcta1= '" & scuenta1 & "' and D_SubCta2= '" & scuenta2 & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub



VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_ro_LiquidaMain 
   BackColor       =   &H00000000&
   Caption         =   "RRHH - Proceso de Pagos por Planilla e Individuales"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "frm_ro_LiquidaMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "frm_ro_LiquidaMain.frx":0A02
      ScaleHeight     =   960
      ScaleWidth      =   15300
      TabIndex        =   31
      Top             =   120
      Width           =   15360
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   2640
         Picture         =   "frm_ro_LiquidaMain.frx":6CA34
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnVer2 
         BackColor       =   &H00808000&
         Caption         =   "Alcance"
         Height          =   720
         Left            =   6000
         Picture         =   "frm_ro_LiquidaMain.frx":6CFEC
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Regitra Alcance del Contrato"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   720
         Left            =   4320
         Picture         =   "frm_ro_LiquidaMain.frx":6DCB6
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "frm_ro_LiquidaMain.frx":6E0F8
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
         Caption         =   "p/Depto"
         Height          =   720
         Left            =   3480
         Picture         =   "frm_ro_LiquidaMain.frx":6E302
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Planilla por Depto. y Mes"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   6840
         Picture         =   "frm_ro_LiquidaMain.frx":6E8BF
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   960
         Picture         =   "frm_ro_LiquidaMain.frx":6EAC9
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar Tramitre"
         Height          =   720
         Left            =   5160
         Picture         =   "frm_ro_LiquidaMain.frx":6F793
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Cerrar Tramite y Archivarlo"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   1800
         Picture         =   "frm_ro_LiquidaMain.frx":70195
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   120
         Picture         =   "frm_ro_LiquidaMain.frx":7039F
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLANILLAS"
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
         Left            =   10335
         TabIndex        =   42
         Top             =   300
         Width           =   1755
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_ro_LiquidaMain.frx":7097F
      ScaleHeight     =   915
      ScaleWidth      =   15300
      TabIndex        =   27
      Top             =   120
      Width           =   15360
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "frm_ro_LiquidaMain.frx":DC9B1
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "frm_ro_LiquidaMain.frx":DCBBB
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENTAS"
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
         Left            =   9915
         TabIndex        =   30
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFC0&
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
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
         Left            =   2880
         TabIndex        =   65
         Top             =   7320
         Width           =   915
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
         Left            =   840
         TabIndex        =   64
         Top             =   7320
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   7200
         Width           =   4545
         _ExtentX        =   8017
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
         Caption         =   " <-- Inicio                                              Fin -->"
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
      Begin MSDataGridLib.DataGrid grdPrincipal 
         Bindings        =   "frm_ro_LiquidaMain.frx":DCDC5
         Height          =   6930
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   12224
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
            DataField       =   "ges_gestion"
            Caption         =   "Gestion"
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
            DataField       =   "planilla_codigo"
            Caption         =   "Planilla"
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
            DataField       =   "mes_grupo"
            Caption         =   "Mes"
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
         BeginProperty Column04 
            DataField       =   "descripcion_grupo"
            Caption         =   "Descripcion.Planilla"
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
            DataField       =   "depto_codigo"
            Caption         =   "Depto"
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
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   -1  'True
               ColumnWidth     =   2145.26
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraDatosProponente 
      BackColor       =   &H00404040&
      Caption         =   "Sub-Grupo por Unidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   3855
      Left            =   5040
      TabIndex        =   13
      Top             =   1200
      Width           =   10455
      Begin VB.PictureBox FrmABMDet2 
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   825
         Left            =   120
         Picture         =   "frm_ro_LiquidaMain.frx":DCDDD
         ScaleHeight     =   765
         ScaleWidth      =   10155
         TabIndex        =   46
         Top             =   240
         Width           =   10215
         Begin VB.CommandButton cmdPagoDesaprob 
            BackColor       =   &H80000018&
            Caption         =   "Desapro."
            Height          =   680
            Left            =   3540
            MaskColor       =   &H00404040&
            Picture         =   "frm_ro_LiquidaMain.frx":148E0F
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   40
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton cmdPagoAnulaDev 
            BackColor       =   &H80000018&
            Caption         =   "Anl.Dev"
            Height          =   680
            Left            =   6015
            MaskColor       =   &H00404040&
            Picture         =   "frm_ro_LiquidaMain.frx":149019
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Anula Registro Activo"
            Top             =   40
            Width           =   765
         End
         Begin VB.CommandButton cmdImprimeLiquida 
            BackColor       =   &H80000018&
            Caption         =   "p.Unidad"
            Height          =   680
            Left            =   7575
            MaskColor       =   &H00404040&
            Picture         =   "frm_ro_LiquidaMain.frx":149CE3
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Listado de Ventas por Servicio"
            Top             =   40
            Width           =   765
         End
         Begin VB.CommandButton cmdPagoDeven 
            BackColor       =   &H80000018&
            Caption         =   "Apr.Dev"
            Height          =   680
            Left            =   5175
            MaskColor       =   &H00404040&
            Picture         =   "frm_ro_LiquidaMain.frx":14A2A0
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Aprueba Registro"
            Top             =   40
            Width           =   765
         End
         Begin VB.CommandButton cmdPagoAprob 
            BackColor       =   &H80000018&
            Caption         =   "Verificar"
            Height          =   680
            Left            =   2680
            MaskColor       =   &H00404040&
            Picture         =   "frm_ro_LiquidaMain.frx":14A4AA
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Aprueba Registro"
            Top             =   40
            Width           =   765
         End
         Begin VB.CommandButton cmdPagoNuevo 
            BackColor       =   &H80000018&
            Caption         =   "Nuevo"
            Height          =   680
            Left            =   120
            Picture         =   "frm_ro_LiquidaMain.frx":14A6B4
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Adiciona Detalle"
            Top             =   40
            Width           =   765
         End
         Begin VB.CommandButton cmdPagoEditar 
            BackColor       =   &H80000018&
            Caption         =   "Modificar"
            Height          =   680
            Left            =   980
            Picture         =   "frm_ro_LiquidaMain.frx":14AAF6
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Modifica Detalle Elegido"
            Top             =   40
            Width           =   765
         End
         Begin VB.CommandButton cmdPagoAnular 
            BackColor       =   &H80000018&
            Caption         =   "Borrar"
            Height          =   680
            Left            =   1830
            Picture         =   "frm_ro_LiquidaMain.frx":14AF38
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Elimina Detalle Elegido"
            Top             =   40
            Visible         =   0   'False
            Width           =   765
         End
      End
      Begin Crystal.CrystalReport CRCrono 
         Left            =   9840
         Top             =   3360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   1
         WindowControlBox=   -1  'True
         WindowMaxButton =   0   'False
         WindowMinButton =   0   'False
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCancelBtn=   0   'False
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSDataGridLib.DataGrid grdLiquida 
         Bindings        =   "frm_ro_LiquidaMain.frx":14B37A
         Height          =   2370
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4180
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "ges_gestion"
            Caption         =   "Gestion"
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
            DataField       =   "planilla_codigo"
            Caption         =   "Planilla"
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
            DataField       =   "mes_grupo"
            Caption         =   "Mes"
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
            DataField       =   "numero_pago"
            Caption         =   "Pago"
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
            DataField       =   "codigo_unidad_pln"
            Caption         =   "Unidad"
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
            DataField       =   "concepto"
            Caption         =   "Concepto"
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
            DataField       =   "fecha_estimada_liq"
            Caption         =   "Fecha.Liq."
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   6419.906
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoLiquida 
         Height          =   330
         Left            =   8160
         Top             =   3480
         Width           =   2145
         _ExtentX        =   3784
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
         Caption         =   " <--       -->"
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
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Totales:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   3480
         TabIndex        =   17
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label lblTotalUSLiq 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$US:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4320
         TabIndex        =   16
         Top             =   3480
         Width           =   1830
      End
      Begin VB.Label lblTotalBSLiq 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bs.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6240
         TabIndex        =   15
         Top             =   3480
         Width           =   1830
      End
      Begin VB.Label lblEstadoLiquida 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblEstadoLiquida"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   3480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblNroLiquida 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro. registros:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   3480
         Width           =   3135
      End
   End
   Begin VB.Frame fraBeneficiario 
      BackColor       =   &H00000000&
      Caption         =   "Personal de la Planilla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   3735
      Left            =   5040
      TabIndex        =   19
      Top             =   5040
      Width           =   10455
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   825
         Left            =   120
         Picture         =   "frm_ro_LiquidaMain.frx":14B393
         ScaleHeight     =   765
         ScaleWidth      =   10155
         TabIndex        =   55
         Top             =   240
         Width           =   10215
         Begin VB.CommandButton cmdBenAnular 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Elim.Pers"
            Height          =   680
            Left            =   1830
            Picture         =   "frm_ro_LiquidaMain.frx":1B73C5
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Elimina Detalle Elegido"
            Top             =   40
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton cmdBenMonto 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Montos"
            Height          =   680
            Left            =   980
            Picture         =   "frm_ro_LiquidaMain.frx":1B7807
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Modifica Detalle Elegido"
            Top             =   40
            Width           =   765
         End
         Begin VB.CommandButton cmdBenNuevo 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Personal"
            Height          =   680
            Left            =   120
            Picture         =   "frm_ro_LiquidaMain.frx":1B7C49
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Adiciona Detalle"
            Top             =   40
            Width           =   765
         End
         Begin VB.CommandButton cmdBenConf 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Conform."
            Height          =   680
            Left            =   2680
            Picture         =   "frm_ro_LiquidaMain.frx":1B808B
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Aprueba Registro"
            Top             =   40
            Width           =   765
         End
         Begin VB.CommandButton cmdDatosContrato 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Contrato"
            Height          =   680
            Left            =   5175
            Picture         =   "frm_ro_LiquidaMain.frx":1B8295
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Aprueba Registro"
            Top             =   40
            Width           =   765
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Cmpbte."
            Height          =   680
            Left            =   7575
            Picture         =   "frm_ro_LiquidaMain.frx":1B849F
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Listado de Ventas por Servicio"
            Top             =   40
            Width           =   765
         End
         Begin VB.CommandButton cmdBenAnularTodo 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Anl.Todo"
            Height          =   680
            Left            =   6015
            Picture         =   "frm_ro_LiquidaMain.frx":1B8A5C
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Anula Registro Activo"
            Top             =   40
            Width           =   765
         End
         Begin VB.CommandButton cmdBenFact 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Factura?"
            Height          =   680
            Left            =   3540
            Picture         =   "frm_ro_LiquidaMain.frx":1B9726
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   40
            Visible         =   0   'False
            Width           =   765
         End
      End
      Begin MSDataGridLib.DataGrid grdBeneficiario 
         Bindings        =   "frm_ro_LiquidaMain.frx":1B9930
         Height          =   2250
         Left            =   120
         TabIndex        =   45
         Top             =   1080
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   3969
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
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
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "beneficiario_codigo"
            Caption         =   "CI"
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
            DataField       =   "sueldo_basico"
            Caption         =   "Sueldo.Basico"
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
            DataField       =   "monto_refrigerio"
            Caption         =   "Refrigerio"
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
            DataField       =   "total_ganado"
            Caption         =   "Total.Ganado"
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
            DataField       =   "total_dsctos"
            Caption         =   "Total.Dsctos."
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
            DataField       =   "liquido_pagable_bs"
            Caption         =   "Liq.Pagable"
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
         BeginProperty Column07 
            DataField       =   "beneficiario_denominacion"
            Caption         =   "Apellidos y Nombres"
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
            DataField       =   "bono_antiguedad"
            Caption         =   "Antiguedad"
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
            DataField       =   "anticipo_sueldo"
            Caption         =   "Anticipo.S.B."
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
            DataField       =   "anticipo_refrigerio"
            Caption         =   "Anticipo.Refr."
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
         BeginProperty Column11 
            DataField       =   "prestamo"
            Caption         =   "Prestamo"
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
         BeginProperty Column12 
            DataField       =   "afp2"
            Caption         =   "FUTURO"
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
         BeginProperty Column13 
            DataField       =   "rciva"
            Caption         =   "RC-IVA"
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
         BeginProperty Column14 
            DataField       =   "otros_dsctos"
            Caption         =   "Otros.Dsctos."
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
               Alignment       =   2
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2385.071
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column12 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1214.929
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoBeneficiario 
         Height          =   330
         Left            =   8160
         Top             =   3360
         Width           =   2145
         _ExtentX        =   3784
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
         BackColor       =   16761024
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
         Caption         =   " <--       -->"
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
      Begin VB.Label lblEstadoBeneficiario 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblEstadoBeneficiario"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblNroBeneficiario 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro. registros:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Totales:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   3480
         TabIndex        =   22
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblTotalUS 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$US:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4320
         TabIndex        =   21
         Top             =   3360
         Width           =   1830
      End
      Begin VB.Label lblTotalBS 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bs.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6240
         TabIndex        =   20
         Top             =   3360
         Width           =   1830
      End
   End
   Begin VB.Label lblUsuario 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuario: XXXXX"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7440
      TabIndex        =   26
      Top             =   9120
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label lblDA 
      Caption         =   "D.A.: XXXXX"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10680
      TabIndex        =   25
      Top             =   8760
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblEstadoGrupoLiq 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblEstadoGrupoLiq"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   9000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblCodUniSol 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   9240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblCodGrupo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      DataField       =   "planilla_codigo"
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
      Height          =   255
      Left            =   6360
      TabIndex        =   10
      Top             =   9600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblDesGrupo 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   9
      Top             =   9600
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Grupo Liquidación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   9600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "Form.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6720
      TabIndex        =   7
      Top             =   8880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFormulario 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "mes_grupo"
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
      Height          =   255
      Left            =   7320
      TabIndex        =   6
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblGestion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      DataField       =   "ges_gestion"
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
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   8880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label30 
      Caption         =   "Gestión:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   8880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblDesUniSol 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   9240
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label Label28 
      Caption         =   "Uni. Sol.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   9240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblNroPrincipal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro. registros:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   8880
      Visible         =   0   'False
      Width           =   4815
   End
End
Attribute VB_Name = "frm_ro_LiquidaMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQLs As String ' usado para la elaboración de los querys
Dim filtro As String ' usado para la elaboración de los querys
Dim nro_reg As Integer
Dim rs_grdPrincipal As ADODB.Recordset ' usado para navegar sobre el grid principal
Dim rs_grdLiquida As ADODB.Recordset ' usado para navegar sobre el grid
Dim rs_grdBeneficiario As ADODB.Recordset ' usado para navegar sobre el grid
Dim rsNada As ADODB.Recordset

    Dim Cad As String '
    Dim swGuardar As Integer ' usado para saber si efectivamente se almaceno o elimino los datos en la base
                          ' swGuarda -> 0 si se realizo el proceso satisfactoriamente
                          ' swGuarda -> 1 si se produjo un evento de cancelar por parte del usuario en el proceso
                          ' swGuarda -> 2 si se produjo un error de integridad de la base de datos en el servidor por el proceso
    Dim RegPuntero As Long ' usada para guardar el código de registro para poder apuntar el el registro seleccionado luego de un refresh
    Dim fechax As String
    Dim horax  As String
    Dim i As Integer

Dim VAR_SW As String
Dim marca1 As String

'Private Sub cmdActualiza_Click()
'    Dim Gestion As String
'    Dim CodUni  As String
'    Dim CodGrupo  As Integer
'
'    Screen.MousePointer = vbHourglass
'
'    cboUnidadSol.BoundText = ""
'    If rs_grdPrincipal.RecordCount > 0 Then
'        ' se guarda los datos de registro para poder ubicar luego el registro
'        Gestion = rs_grdPrincipal!ges_gestion
'        CodUni = rs_grdPrincipal!unidad_codigo
'        CodGrupo = rs_grdPrincipal!CODIGO_GRUPO
'    End If
'    Call pl_RefrescaListaPrincipal ' se refresca el recordset para mostrar los datos originales
'
'    If Len(Gestion) > 0 Then ' si se tiene un registro activo
'        rs_grdPrincipal.Find " ges_gestion ='" & Gestion & "'" ' el puntero de registro se ubica en la posicion guardada
'        rs_grdPrincipal.Find " unidad_codigo ='" & CodUni & "'" 'el puntero de registro se ubica en la posicion guardada
'        rs_grdPrincipal.Find " codigo_grupo =" & CodGrupo ' el puntero de registro se ubica en la posicion guardada
'    End If
'    Call grdPrincipal_RowColChange(0, 0)
'    Screen.MousePointer = vbDefault
'
'End Sub

Private Sub cmdAnulaDev_Click()
    Call pl_OpcionesGenericas("Tool_AnulaOrdLiq", "Liquida")
End Sub

Private Sub BtnEliminar_Click()
    Call pl_OpcionesGenericas("Tool_Anular", "GrupoLiq")
End Sub

Private Sub cmdBenAnularTodo_Click()
    Call pl_OpcionesGenericas("Tool_AnularTodo", "Beneficiario")
End Sub

Private Sub cmdBenConf_Click()
    Call pl_OpcionesGenericas("Tool_Conformidad", "Beneficiario")
End Sub

Private Sub cmdBenAnular_Click()
    Call pl_OpcionesGenericas("Tool_Anular", "Beneficiario")
End Sub

Private Sub cmdBenFact_Click()
    Call pl_OpcionesGenericas("Tool_Factura", "Beneficiario")
End Sub

Private Sub cmdBenMonto_Click()
    Call pl_OpcionesGenericas("Tool_Monto", "Beneficiario")
End Sub

Private Sub cmdBenNuevo_Click()
'    Call pl_OpcionesGenericas("Tool_Nuevo", "Beneficiario")
   If Ado_datos.Recordset.RecordCount > 0 Then
        If AdoLiquida.Recordset.RecordCount > 0 Then
            marca1 = Ado_datos.Recordset.Bookmark
            AdoBeneficiario.Recordset.AddNew
            ro_Personal_Planilla.txtSW = "ADD"
            ro_Personal_Planilla.TxtGestion = Ado_datos.Recordset!ges_gestion
            ro_Personal_Planilla.txtBenef = Ado_datos.Recordset!planilla_codigo
            ro_Personal_Planilla.TxtInicial = Ado_datos.Recordset!mes_grupo
            ro_Personal_Planilla.TxtAprob = "REG"
            ro_Personal_Planilla.TxtLquida = AdoLiquida.Recordset!NUMERO_PAGO
'            ro_Personal_Planilla.Txtpago3 = adoLista.Recordset!beneficiario_haber_mensual
            ro_Personal_Planilla.Show vbModal
            'Call abrirtabla
            'AdoLiquidacion.Refresh
        End If
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Personal"
   End If

End Sub

Private Sub cmdBenNuevoSegun_Click()
    Call pl_OpcionesGenericas("Tool_NuevoSegun", "Beneficiario")
End Sub

'Private Sub cmdBusca_Click()
'    If rs_grdPrincipal.RecordCount > 0 Then
'        Call pg_BuscaTdbGrid(grdPrincipal, rs_grdPrincipal, grdPrincipal.Columns(grdPrincipal.Col).DataField)
'        Call pl_PersonalizaGridPrincipal
'      Else
'        MsgBox "No existen registros para búscar.", vbInformation, "Aviso"
'    End If
'
'End Sub

Private Sub cmdDatosContrato_Click()
    Call pl_OpcionesGenericas("Tool_Contrato", "Beneficiario")
End Sub

Private Sub BtnModificar_Click()
    'Call pl_OpcionesGenericas("Tool_Editar", "GrupoLiq")
    
    If rs_grdPrincipal.RecordCount > 0 Then
         VAR_SW = "MOD"
                ' ***********************
                ' se llama al formulario
                ' ***********************
                Screen.MousePointer = vbHourglass
'                lblEstadoGrupoLiq.Caption = "E" ' se esta en modo de edicion del reg. actual
                frm_ro_LiquidaAdiGrupo.Show vbModal
                
'                If lblEstadoGrupoLiq.Caption <> "E" Then  ' si el proceso la modificacion
                    Cad = lblCodUniSol.Caption   ' codigo unidad
                    RegPuntero = lblCodGrupo.Caption ' codigo grupo
                    frm_ro_LiquidaAdiGrupo.txtCodGrupo = lblCodGrupo.Caption        ' codigo grupo
                    frm_ro_LiquidaAdiGrupo.txtCodUnidad = lblFormulario.Caption     'Mes Grupo
                    Call pl_RefrescaListaPrincipal
                    rs_grdPrincipal.Find "unidad_codigo = '" & Cad & "'" ' se posisicona en el registro editado
                    rs_grdPrincipal.Find "codigo_grupo = " & RegPuntero ' se posisicona en el registro editado
                    Call grdPrincipal_RowColChange(0, 0)
                    MsgBox "Los cambios se realizaron satisfactoriamente", vbInformation, "Aviso"
                    grdPrincipal.SetFocus
'                End If
'                lblEstadoGrupoLiq.Caption = "" ' no se esta editando ni adicionando registros
              Else
                MsgBox "No existen registro de grupos de liquidación para ser editado.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
    End If
End Sub

'Private Sub cmdFiltro_Click()
'    'PROPÓSITO      : Realiza la filtración de una especificación sobre la columna de la celda activa
'
'    Dim CadFiltro As String ' usada para almacenar la cedena de filtración
'    Dim micriterio As String ' critetio de filtración
'    Dim CampoAct As Integer ' nombre del campo activo
'    Dim ColFiltro As String ' usada para almacenar la columna activa por el cual se filtrará
'    Dim a As Integer ' usada para ver el formato de cadena a filtrar
'
'    If rs_grdPrincipal.RecordCount > 0 Then
'
'        micriterio = "Digite " & LCase(grdPrincipal.Columns(grdPrincipal.Col).Caption) & " a filtrar"
'        CadFiltro = pg_QuitaEspBlanco(UCase(InputBox(micriterio, "Filtración")))
'        ' verificamos que la cadena sea del tipo *a* donde a representa cualquier secuencia de caracteres
'        Select Case Len(CadFiltro)
'          Case Is >= 3
'            If Left(CadFiltro, 1) = "*" And Right(CadFiltro, 1) <> "*" Then
'                ' completamos la cadena al tipo *a*
'                CadFiltro = CadFiltro & "*"
'              Else
'                ' es del tipo a* o a que son cadenas validas
'                'CadFiltro = "*" & CadFiltro
'            End If
'          Case 2
'            If Left(CadFiltro, 1) = "*" And Right(CadFiltro, 1) <> "*" Then
'                CadFiltro = CadFiltro & "*"
'              Else
'                If Left(CadFiltro, 1) = "*" And Right(CadFiltro, 1) = "*" Then
'                    ' si ambos son *
'                    CadFiltro = ""
'                  Else
'                    ' es del tipo a* o a que son cadenas validas
'                    'CadFiltro = "*" & CadFiltro
'                End If
'            End If
'          Case 1
'            If CadFiltro = "*" Then CadFiltro = ""
'        End Select
'
'        On Error GoTo EtiqError
'
'        If Len(CadFiltro) > 0 Then ' si introdujo una cadena a filtrar
'            CampoAct = grdPrincipal.Col
'            ColFiltro = grdPrincipal.Columns(grdPrincipal.Col).DataField
'            ' verificamos si la longitud coincide con el tamaño del campo
'            If Len(CadFiltro) <= rs_grdPrincipal.Fields(ColFiltro).DefinedSize Or rs_grdPrincipal.Fields(ColFiltro).Type = 3 Then
'                If rs_grdPrincipal.Filter = 0 Then ' es la primera filtración
'                    rs_grdPrincipal.Filter = ColFiltro & " like " & Chr(39) & CadFiltro & Chr(39)
'                    filtro = grdPrincipal.Columns(CampoAct).Caption & " -> " & Chr(39) & CadFiltro & Chr(39) ' concatenamos la cadena de filtración
'                  Else
'                    rs_grdPrincipal.Filter = rs_grdPrincipal.Filter & " AND " & ColFiltro & " like " & Chr(39) & CadFiltro & Chr(39)
'                    filtro = filtro & ", " & grdPrincipal.Columns(CampoAct).Caption & " -> " & Chr(39) & CadFiltro & Chr(39)  ' concatenamos la cadena de filtración
'                End If
'                lblNroPrincipal.Caption = "Nro. de Liq: " & rs_grdPrincipal.RecordCount & " Filtro ( " & filtro & " )"
'
'                If rs_grdPrincipal.RecordCount = 0 Then ' no se encontraron coincidencias
'                    Call grdPrincipal_RowColChange(0, 0)
'                    MsgBox "No se encontró ninguna coincidencia con " & filtro, vbInformation, "Información"
'
'                End If
'                grdPrincipal.SetFocus
'
'              Else ' la longitud de la cadena a filtrar es mayor a la longitud del campo
'                MsgBox "La longitud de la cadena a filtrar -> " & CadFiltro & " es mayor a la longitud del permitido por " & grdPrincipal.Columns(CampoAct).Caption, vbInformation, "Información"
'            End If
'            grdPrincipal.SetFocus
'          Else ' solo tiene el foco
'            grdPrincipal.SetFocus
'        End If
'      Else
'        MsgBox "No existen registros para ser filtrados.", vbInformation, "Aviso"
'    End If
'
'    Call pl_PersonalizaGridPrincipal
'
'    On Error GoTo 0 ' desactiva el manejador de errores
'    Exit Sub
'
'EtiqError:
'    Select Case Err.Number
'      Case -2147352571
'        MsgBox "Error: No se pueden filtrar los datos, los tipos no coinciden." & Chr(13) & Chr(13) & "No se realizo la filtración de datos." & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description, vbCritical, "Error"
'      Case Else ' si se produjo otro tipo de error
'        MsgBox "Error: No se realizo la filtración de datos." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
'    End Select
'
'End Sub

'Private Sub cmdImprime_Click()
'
'    If rs_grdPrincipal.RecordCount > 0 Then
'        Call pg_Imprimir(grdPrincipal, grdPrincipal.Caption)
'      Else
'        MsgBox "No existen registros para ser impresos.", vbInformation, "Aviso"
'    End If
'
'End Sub

Private Sub cmdImprimeLiquida_Click()
'    If rs_grdPrincipal.RecordCount > 0 Then
'        If rs_grdLiquida.RecordCount > 0 Then
'            ''** llama al formulario
'            lblEstadoBeneficiario.Tag = rs_grdLiquida!NUMERO_PAGO & ""
'            frm_ro_PagosPrintOrdenPago.Show vbModal
'          Else
'            MsgBox "No existe resgistros para liquidación." & Chr(13) & "Corrija el error e intente imprimir nuevamente.", vbInformation, "Aviso"
'        End If
'      Else
'        MsgBox "No existen registros para ser impresos.", vbInformation, "Aviso"
'    End If

End Sub

Private Sub BtnAñadir_Click()
    Call pl_OpcionesGenericas("Tool_Nuevo", "GrupoLiq")
End Sub

'Private Sub cmdOrdAZ_Click()
'    If rs_grdPrincipal.RecordCount > 0 Then
'        Call pg_OrdenaTdbGrid(grdPrincipal, rs_grdPrincipal, True)
'        Call pl_PersonalizaGridPrincipal
'      Else
'        MsgBox "No existen registros para ordenar.", vbInformation, "Aviso"
'    End If
'
'End Sub
'
'Private Sub cmdOrdZA_Click()
'    If rs_grdPrincipal.RecordCount > 0 Then
'        Call pg_OrdenaTdbGrid(grdPrincipal, rs_grdPrincipal, False)
'        Call pl_PersonalizaGridPrincipal
'      Else
'        MsgBox "No existen registros para ordenar.", vbInformation, "Aviso"
'    End If
'
'End Sub

Private Sub cmdPagoAnulaDev_Click()
    Call pl_OpcionesGenericas("Tool_AnulaDevengar", "Liquida")
End Sub

Private Sub cmdPagoanular_Click()
    Call pl_OpcionesGenericas("Tool_Anular", "Liquida")
End Sub

Private Sub cmdPagoAprob_Click()
    Call pl_OpcionesGenericas("Tool_Aprobar", "Liquida")
End Sub

Private Sub cmdPagoDesaprob_Click()
    Call pl_OpcionesGenericas("Tool_Desaprobar", "Liquida")
End Sub

Private Sub cmdPagoDeven_Click()
    Call pl_OpcionesGenericas("Tool_Devengar", "Liquida")
End Sub

Private Sub cmdPagoEditar_Click()
    Call pl_OpcionesGenericas("Tool_Editar", "Liquida")
End Sub

Private Sub cmdPagoNuevo_Click()
    Call pl_OpcionesGenericas("Tool_Nuevo", "Liquida")
End Sub

Private Sub BtnVer2_Click()
'    If rs_grdPrincipal.RecordCount > 0 Then
'        If Not (rs_grdPrincipal.EOF Or rs_grdPrincipal.BOF) Then
'            frm_ro_HistRelCronoCompro.xGes_Gestion = rs_grdPrincipal!ges_gestion
'            frm_ro_HistRelCronoCompro.xunidad_codigo = rs_grdPrincipal!unidad_codigo
'            frm_ro_HistRelCronoCompro.xCodigo_Grupo = rs_grdPrincipal!CODIGO_GRUPO
'        End If
'    End If
'    If rs_grdBeneficiario.RecordCount > 0 Then
'        If Not (rs_grdBeneficiario.EOF Or rs_grdBeneficiario.BOF) Then
'            frm_ro_HistRelCronoCompro.xNumero_Pago = rs_grdBeneficiario!NUMERO_PAGO
'            'frm_ro_HistRelCronoCompro.Xcodigo_beneficiario = rs_grdBeneficiario!codigo_beneficiario
'        End If
'    End If
'    frm_ro_HistRelCronoCompro.Show vbModal

End Sub

Private Sub BtnSalir_Click()
    Call pl_OpcionesGenericas("Tool_Salir", "GrupoLiq")
End Sub

Private Sub Form_Load()
    ' obtiene direccion administrativa en funcion del usuario
'    Set rstTemp = New ADODB.Recordset
'    SQLs = "SELECT gc_usuarios.da_codigo, gc_direccion_administrativa.da_descripcion "
'    SQLs = SQLs & "FROM gc_usuarios INNER JOIN gc_direccion_administrativa ON gc_usuarios.da_codigo = gc_direccion_administrativa.da_codigo "
'    SQLs = SQLs & "WHERE gc_usuarios.usr_codigo = '" & glusuario & "' AND gc_usuarios.estado_codigo = 'APR' "
'    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
'    If rstTemp.RecordCount > 0 Then
'        GldaCodigo = rstTemp!da_codigo & ""
'        GldaDescrip = rstTemp!da_descripcion & ""
'      Else
'        GldaCodigo = ""
'        GldaDescrip = ""
'        MsgBox "Error: No existe relación entre el usuario y una Dirección Administrativa." & Chr(13) & "Esto puede causar muchos errores." & Chr(13) & "Anote el error y comuniquese con el admisnitrador del sistema.", vbError, "Aviso"
'    End If
    GldaCodigo = "1.1"
    GldaDescrip = "GERENCIA ADMINISTRATIVA"

'    Call pl_Llena_Combos_Base 'llena los combos base

    Call pl_RefrescaListaPrincipal 'refresca la lista principal

'    Call pl_ValoresDefecto

    Call grdPrincipal_RowColChange(0, 0)

    VAR_SW = "NNN"
'    Call ABRIR_TABLAS_AUX
'    Call OptFilGral1_Click

'''/***
''DE.Edson.Open
''DE.Edson.Execute "SET DATEFORMAT dmy"
'''**/

	Call SeguridadSet(Me)
End Sub

Private Sub pl_Llena_Combos_Base()
    ' llena los combos y listas base para la carga del formulario
    
'    Dim rstTemp As ADODB.Recordset ' usado para la carga de los combos de base
'
'    ' unidad solicitante - para realizar filtro sobre el grid principal
'    Set rstTemp = New ADODB.Recordset
'    SQLs = "select unidad_codigo, unidad_codigo + ' - '+ uni_descripcion_larga as des_unidad from gc_unidad_ejecutora where uni_activo = 'S' ORDER BY unidad_codigo"
'    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
'    If rstTemp.RecordCount > 0 Then
'        Set cboUnidadSol.RowSource = rstTemp
'        cboUnidadSol.BoundColumn = "unidad_codigo"
'        cboUnidadSol.ListField = "des_unidad"
'
'      Else
'        MsgBox "El catalogo de unidad solicitante no esta actualizado.", vbInformation, "Aviso"
'    End If
'
'    Set rstTemp = Nothing

End Sub

Private Sub pl_ValoresDefecto()
    'PROPOSITO:             Permite establecer los valores por defecto de los elementos del formulario
    
    lblUsuario.Caption = "Usuario: " & glusuario ' usuario
    lblDA.Caption = "Dir. Adm.: " & GldaDescrip ' descripcion direc. adm.
    
    Select Case GldaCodigo
      Case "01" ' si da es DGAARRYHH
        
        Me.Caption = "SAF - Proceso de Liquidación de Consultor"
        lbl_titulo.Caption = "PROCESO DE LIQUIDACIÓN CONSULTOR"
      
      Case "52" ' si da es DAP
        
        Me.Caption = "SAF - Proceso de Liquidación de Consultor"
        lbl_titulo.Caption = "PROCESO DE LIQUIDACIÓN CONSULTOR"
      
      Case "1.1" ' es para recursos humanos RRYHH
        
        Me.Caption = "Proceso de Pagos por Planillas"
        lbl_titulo.Caption = "PERSONAL PERMANENTE"
    
    End Select
    
End Sub

Private Sub pl_RefrescaListaPrincipal()
  ' Onjetivo: Procedimiento refrescar o actualizar la lista principal de solicitudes
  
    On Error GoTo 0 ' activamos el manejador de errores
  
    Screen.MousePointer = vbHourglass
'    cboUnidadSol.BoundText = ""
    
'    SQLs = "SELECT ao_pagos_grupos.unidad_codigo, ao_pagos_grupos.codigo_solicitud, ao_pagos_grupos.codigo_grupo, ao_pagos_grupos.descripcion_grupo, "
'    SQLs = SQLs & "'ModPago' = case when ao_pagos_grupos.modalidad_pago = 'P' then 'Planilla' else 'Invividual' end, ao_pagos_grupos.estado_codigo, gc_unidad_ejecutora.Uni_descripcion_larga, ac_tipo_tramite.denominacion_tipo, ao_pagos_grupos.ges_gestion, ao_pagos_grupos.modalidad_pago, "
'    SQLs = SQLs & "ao_pagos_grupos.formulario , ao_pagos_grupos.da, ao_pagos_grupos.numero_consultoria, ao_pagos_grupos.correl_grupo_da "
'    SQLs = SQLs & "FROM ac_tipo_tramite INNER JOIN ao_pagos_grupos ON ac_tipo_tramite.tipo_formulario = ao_pagos_grupos.formulario LEFT OUTER JOIN gc_unidad_ejecutora ON "
'    SQLs = SQLs & "ao_pagos_grupos.unidad_codigo = gc_unidad_ejecutora.unidad_codigo "
'    Select Case glProceso
'      Case "F05"
'        SQLs = SQLs & "WHERE ao_pagos_grupos.estado_codigo <> 'E' AND ao_pagos_grupos.formulario = 'F05' and ao_pagos_grupos.da = '" & GldaCodigo & "'"
'      Case "F10"
'        SQLs = SQLs & "WHERE ao_pagos_grupos.estado_codigo <> 'E' AND ao_pagos_grupos.formulario = 'F10' "
'    End Select
'    SQLs = SQLs & "ORDER BY ao_pagos_grupos.ges_gestion, ao_pagos_grupos.unidad_codigo, ao_pagos_grupos.codigo_grupo"
    SQLs = "SELECT * from ro_pagos_grupos order by ges_gestion, mes_grupo, planilla_codigo "
    Set rs_grdPrincipal = New ADODB.Recordset
    rs_grdPrincipal.Open SQLs, db, adOpenKeyset, adLockOptimistic 'adOpenStatic, adLockReadOnly
    Set Ado_datos.Recordset = rs_grdPrincipal.DataSource
    Set grdPrincipal.DataSource = Ado_datos.Recordset
    
'    Set grdPrincipal.DataSource = rs_grdPrincipal
'    Call pl_PersonalizaGridPrincipal
'    lblNroPrincipal.Caption = "Nro. Planilla: " & rs_grdPrincipal.RecordCount
    
    Screen.MousePointer = vbDefault
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
      Exit Sub
  Else
      If Ado_datos.Recordset.RecordCount > 0 Then
        If VAR_SW <> "MOD" Then
    '        Call ABRIR_TABLA_DET
            Call pl_RefrescaLiquidacion
            Call pl_RefrescaBeneficiario
        Else
            'Set rs_det1 = New ADODB.Recordset
            Set grdLiquida.DataSource = rsNada
            'Set DtgLaborales.DataSource = rsNada
        End If
        
        'FraDet1.Caption = FraDet1.Caption + parametro
    '    txt_aux9.Text = dtc_desc9.Text
        If Ado_datos.Recordset!estado_codigo = "REG" Then
    '            FrmABMDet2.Visible = True
    '            FrmABMDet.Visible = True
    '            FrmABMDet3.Visible = True
        Else
    '            FrmABMDet2.Visible = False
    '            FrmABMDet.Visible = False
    '            FrmABMDet3.Visible = False
        End If
      End If
  End If
End Sub

Private Sub pl_PersonalizaGridPrincipal()
    'TITULO:                Procedimiento pl_PersonalizaGridPrincipal
    'PROPOSITO:             Personalizar los captions, anchos, etc. del grid
    'EJEMPLO DE LLAMADA:    call pl_PersonalizaGridPrincipal
        
    Dim i As Integer
    
    ' define ancho de columnas y titulo de la cabecera
    grdPrincipal.Columns(0).Width = 750 ' codigo unidad
    grdPrincipal.Columns(0).Caption = "Unidad"
    grdPrincipal.Columns(1).Width = 700 ' codigo solicitud
    grdPrincipal.Columns(1).Caption = "Solicitud Original"
    grdPrincipal.Columns(2).Width = 600 ' cod grupo
    grdPrincipal.Columns(2).Caption = "Código Grupo"
    grdPrincipal.Columns(3).Width = 1800 ' des grupo
    grdPrincipal.Columns(3).Caption = "Descripción Grupo"
    grdPrincipal.Columns(4).Width = 900 ' Modalidad pago
    grdPrincipal.Columns(4).Caption = "Modalidad Pago"
    grdPrincipal.Columns(5).Width = 800 ' estado aprobado
    grdPrincipal.Columns(5).Caption = "Est. Aprobado"
    
    For i = 6 To rs_grdPrincipal.Fields.Count - 1
        grdPrincipal.Columns(i).Visible = False
        grdPrincipal.Columns(i).AllowSizing = False
    Next i
    
    
End Sub

Private Sub pl_OpcionesGenericas(TipoOpcion As String, Proceso As String)
    'TITULO:                Procedimiento pl_OpcionesGenericas
    'PROPOSITO:             Ejecuta una opcion del toolbar
    'EJEMPLO DE LLAMADA:    call pl_OpcionesGenericas(TipoOpcion)
    'ENTRADAS:              TipoOpcion = Opción a elegir (Grabar,Editar, etc.)
                            ' Realiza una acción según TipoOpcion

'    Dim Cad As String '
'    Dim swGuardar As Integer ' usado para saber si efectivamente se almaceno o elimino los datos en la base
'                          ' swGuarda -> 0 si se realizo el proceso satisfactoriamente
'                          ' swGuarda -> 1 si se produjo un evento de cancelar por parte del usuario en el proceso
'                          ' swGuarda -> 2 si se produjo un error de integridad de la base de datos en el servidor por el proceso
'    Dim RegPuntero As Long ' usada para guardar el código de registro para poder apuntar el el registro seleccionado luego de un refresh
'    Dim fechax As String
'    Dim horax  As String
'    Dim i As Integer
    
    On Error GoTo EtiqError
    
    Select Case Proceso
      
      ' ********************************************************
      ' opciones genericas de la ficha: GRUPOS
      ' ********************************************************
      
      Case "GrupoLiq" ' procesa la ficha GRUPOS
        Select Case TipoOpcion
          Case "Tool_Nuevo"
            ' no se procesa en este modulo
            
          Case "Tool_Editar"
            
            If rs_grdPrincipal.RecordCount > 0 Then
                
                ' ***********************
                ' se llama al formulario
                ' ***********************
                Screen.MousePointer = vbHourglass
'                lblEstadoGrupoLiq.Caption = "E" ' se esta en modo de edicion del reg. actual
                frm_ro_LiquidaAdiGrupo.Show vbModal
                
'                If lblEstadoGrupoLiq.Caption <> "E" Then  ' si el proceso la modificacion
                    Cad = lblCodUniSol.Caption   ' codigo unidad
                    RegPuntero = Val(lblCodGrupo.Caption) ' codigo grupo
                    Call pl_RefrescaListaPrincipal
                    rs_grdPrincipal.Find "unidad_codigo = '" & Cad & "'" ' se posisicona en el registro editado
                    rs_grdPrincipal.Find "codigo_grupo = " & RegPuntero ' se posisicona en el registro editado
                    Call grdPrincipal_RowColChange(0, 0)
                    MsgBox "Los cambios se realizaron satisfactoriamente", vbInformation, "Aviso"
                    grdPrincipal.SetFocus
'                End If
'                lblEstadoGrupoLiq.Caption = "" ' no se esta editando ni adicionando registros
              Else
                MsgBox "No existen registro de grupos de liquidación para ser editado.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
            
          Case "Tool_Anular"
            
            If rs_grdPrincipal.RecordCount > 0 Then ' si existen registros
                
                If fl_VerificaEliminaGrupo Then
                    Screen.MousePointer = vbHourglass
                    Cad = "Se eliminará el grupo de liquidación, se anulará el comprometido,se desaprueba la adjudicación y registro de contrato:" & Chr(13)
                    Cad = Cad & "Grupo:[" & lblCodGrupo.Caption & " - " & lblDesGrupo.Caption & "]" & Chr(13) & "Unidad: [" & lblCodUniSol.Caption & " - " & lblDesUniSol.Caption & "]." & Chr(13) & "Desea continuar no podrá revertir el proceso?"
                    If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de elimnación") Then
                        'JQ QR
                        'DE.dbo_ap_PagosBorraGrupo rs_grdPrincipal!ges_gestion, rs_grdPrincipal!unidad_codigo, rs_grdPrincipal!codigo_grupo, rs_grdPrincipal!numero_consultoria
                        Call pl_RefrescaListaPrincipal
                        Call grdPrincipal_RowColChange(0, 0)
                        MsgBox "La información se eliminó satisfactoriamente", vbInformation, "Aviso"
                    End If
                  Else
                    grdPrincipal.SetFocus
                    
                End If
                Screen.MousePointer = vbDefault
              Else
                MsgBox "No existen registros para ser eliminados.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
                    
          Case "Tool_Salir"
            Unload Me
            
        End Select
    
      ' ********************************************************
      ' opciones genericas de la: LIQUIDACION
      ' ********************************************************
      
      Case "Liquida" ' procesa la ficha LIQUIDACION
        
        Select Case TipoOpcion
          Case "Tool_Nuevo"
            ' ***********************
            ' se llama al formulario
            ' ***********************
            Screen.MousePointer = vbHourglass
            lblEstadoLiquida.Caption = "REG" ' se esta en modo de adicion de nuevo registro
            frm_ro_LiquidaAdiPago.Show vbModal
            
            If lblEstadoLiquida.Caption <> "REG" Then ' si el proceso realizo la adicion de un registro
                RegPuntero = Val(lblEstadoLiquida.Caption)  ' codigo
                Call grdPrincipal_RowColChange(0, 0)
                rs_grdLiquida.Find "numero_pago = " & RegPuntero ' se posisicona en el registro editado
                Call grdLiquida_RowColChange(0, 0)
                MsgBox "Los cambios se realizaron satisfactoriamente", vbInformation, "Aviso"
                grdLiquida.SetFocus
            End If
            lblEstadoLiquida.Caption = "" ' no se esta editando ni adicionando registros
            
          Case "Tool_Editar"
            
            If rs_grdLiquida.RecordCount > 0 Then

                ' ***********************
                ' se llama al formulario
                ' ***********************
                Screen.MousePointer = vbHourglass
                lblEstadoLiquida.Caption = "E" ' se esta en modo de edicion del reg. actual
                lblEstadoLiquida.Tag = rs_grdLiquida!NUMERO_PAGO ' guara el nuemro de pago q se usa como parametro
                frm_ro_LiquidaAdiPago.Show vbModal

                If lblEstadoLiquida.Caption <> "E" Then   ' si el proceso realizo la edicion o moedidificcacion de los montos
                    RegPuntero = rs_grdLiquida!NUMERO_PAGO   ' numero pago
                    Call grdPrincipal_RowColChange(0, 0)
                    rs_grdLiquida.Find "numero_pago = " & RegPuntero ' se posisicona en el registro editado
                    Call grdLiquida_RowColChange(0, 0)
                    MsgBox "Los cambios se realizaron satisfactoriamente", vbInformation, "Aviso"
                    grdLiquida.SetFocus
                End If
                lblEstadoLiquida.Caption = "" ' no se esta editando ni adicionando registros
              Else
                MsgBox "No existen el registro de liquidación para ser editado.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
            
          Case "Tool_Anular"
            
            If rs_grdPrincipal.RecordCount > 0 Then ' si existen registros
                If fl_VerificaEliminaLiquida Then
                    Screen.MousePointer = vbHourglass
                    Cad = "Se eliminará el registro de liquidación número: [" & rs_grdLiquida!NUMERO_PAGO & "] del grupo: [" & lblCodGrupo.Caption & " - " & lblDesGrupo.Caption & "], Unidad: [" & rs_grdPrincipal!unidad_codigo & "]."
                    If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de elimnación") Then
                        'JQ QR
                        'DE.dbo_ap_PagosBorraPago rs_grdPrincipal!ges_gestion, rs_grdPrincipal!unidad_codigo, rs_grdPrincipal!codigo_grupo, rs_grdLiquida!NUMERO_PAGO
                        Call pl_RefrescaLiquidacion
                        Call grdLiquida_RowColChange(0, 0)
                        MsgBox "La información se eliminó satisfactoriamente", vbInformation, "Aviso"
                    End If
                    Screen.MousePointer = vbDefault
                  Else
                    grdPrincipal.SetFocus
                End If
              Else
                MsgBox "No existen registros para ser eliminados.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
          
          Case "Tool_Aprobar"
          
            If rs_grdLiquida.RecordCount > 0 Then
                If rs_grdBeneficiario.RecordCount > 0 Then
                    
                    If fl_VerificaAprobar Then
                        Screen.MousePointer = vbHourglass
                        Cad = "Se aprobará la liquidación correspondiente a: " & Chr(13)
                        Cad = Cad & "Número de liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "] [" & rs_grdLiquida!Concepto & "]" & Chr(13)
                        Cad = Cad & "Beneficiarios: [" & rs_grdBeneficiario.RecordCount & "]" & Chr(13)
                        Cad = Cad & "Desea aprobar la liquidación?. No podrá modificar mas datos."
                        
                        If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de aprobación") Then
                            ' se actualiza las banderas de aprobacion a 'S' para q no puedan ser modificadas
                            'JQ QR
                            'DE.dbo_ap_GetServDateTime fechax, horax
                            
                            If rs_grdPrincipal!estado_codigo <> "APR" Then
                                SQLs = "UPDATE ro_pagos_grupos SET estado_codigo ='APR', usr_aprueba = '" & glusuario & "', fecha_aprueba = '" & fechax & "', hora_aprueba = '" & horax & "' "
                                SQLs = SQLs & "WHERE ges_gestion = '" & LblGestion.Caption & "' and unidad_codigo = '" & lblCodUniSol.Caption & "' and planilla_codigo = " & Val(lblCodGrupo.Caption)
                                'JQ QR
                                'DE.dbo_apGeneralSearching SQLs
                            End If
                            
                            SQLs = "UPDATE ro_pagos_cronograma SET estado_codigo ='APR', usr_aprueba = '" & glusuario & "', fecha_aprueba = '" & fechax & "', hora_aprueba = '" & horax & "' "
                            SQLs = SQLs & "WHERE ges_gestion = '" & LblGestion.Caption & "' and unidad_codigo = '" & lblCodUniSol.Caption & "' and planilla_codigo = " & Val(lblCodGrupo.Caption) & " and numero_pago = " & rs_grdLiquida!NUMERO_PAGO
                            'JQ QR
                            'DE.dbo_apGeneralSearching SQLs
                            
                            If rs_grdPrincipal!estado_codigo <> "APR" Then
                                fechax = LblGestion.Caption
                                Cad = lblCodUniSol.Caption
                                RegPuntero = rs_grdPrincipal!CODIGO_GRUPO
                                nro_reg = rs_grdLiquida!NUMERO_PAGO
                                ' se posisicona en el registro editado
                                Call pl_RefrescaListaPrincipal
                                rs_grdPrincipal.Find "ges_gestion = '" & fechax & "'"
                                rs_grdPrincipal.Find "unidad_codigo ='" & Cad & "'"
                                rs_grdPrincipal.Find "codigo_grupo =" & RegPuntero
                                Call grdPrincipal_RowColChange(0, 0)
                                rs_grdLiquida.Find "numero_pago =" & nro_reg
                                Call grdLiquida_RowColChange(0, 0)
                                MsgBox "Se aprobo la liquidación correspondiente.", vbInformation, "Aviso"
'                                grdLiquida.SetFocus
                              Else
                                RegPuntero = rs_grdLiquida!NUMERO_PAGO
                                Call grdPrincipal_RowColChange(0, 0)
                                rs_grdLiquida.Find "numero_pago =" & RegPuntero ' se posiciona en el registro editado
                                Call grdLiquida_RowColChange(0, 0)
                                MsgBox "Se aprobo la liquidación correspondiente.", vbInformation, "Aviso"
                            End If
                            
                         End If
                        grdBeneficiario.SetFocus
                        Screen.MousePointer = vbDefault
                    End If
                    
                  Else
                    MsgBox "No existe registro de beneficiario para ser procesado.", vbInformation, "Aviso"
                End If
              Else
                MsgBox "No existe registro de liquidación para ser procesado.", vbInformation, "Aviso"
            End If
            
          Case "Tool_Desaprobar"
          
            If rs_grdLiquida.RecordCount > 0 Then
                If rs_grdBeneficiario.RecordCount > 0 Then
                    Screen.MousePointer = vbHourglass
                    Cad = "Se desaprobará la liquidación correspondiente a: " & Chr(13)
                    Cad = Cad & "Número de liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "] [" & rs_grdLiquida!Concepto & "]" & Chr(13)
                    Cad = Cad & "Beneficiarios: [" & rs_grdBeneficiario.RecordCount & "]" & Chr(13)
                    Cad = Cad & "Desea desaprobar la liquidación?. Podrá modificar los datos."
                    
                    If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación") Then
                        'JQ QR
                        'DE.dbo_ap_GetServDateTime fechax, horax
                        ' se actualiza las banderas de aprobacion a 'N'
                        SQLs = "SELECT * FROM ro_pagos_cronograma WHERE estado_DEVENGADO ='REG' AND ges_gestion = '" & LblGestion.Caption & "' and unidad_codigo = '" & lblCodUniSol.Caption & "' and planilla_codigo = " & Val(lblCodGrupo.Caption)
                        Set rstTemp = New ADODB.Recordset
                        rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
                        If rstTemp.RecordCount = 0 Then
                            SQLs = "UPDATE ro_pagos_grupos SET estado_codigo ='REG', usr_aprueba = '" & glusuario & "', fecha_aprueba = '" & fechax & "', hora_aprueba = '" & horax & "' "
                            SQLs = SQLs & "WHERE ges_gestion = '" & LblGestion.Caption & "' and unidad_codigo = '" & lblCodUniSol.Caption & "' and planilla_codigo = " & Val(lblCodGrupo.Caption)
                            'JQ QR
                            'DE.dbo_apGeneralSearching SQLs
                            
                            SQLs = "UPDATE ro_pagos_cronograma SET estado_codigo ='REG', usr_aprueba = '" & glusuario & "', fecha_aprueba = '" & fechax & "', hora_aprueba = '" & horax & "' "
                            SQLs = SQLs & "WHERE ges_gestion = '" & LblGestion.Caption & "' and unidad_codigo = '" & lblCodUniSol.Caption & "' and planilla_codigo = " & Val(lblCodGrupo.Caption) & " and numero_pago = " & rs_grdLiquida!NUMERO_PAGO
                            'JQ QR
                            'DE.dbo_apGeneralSearching SQLs
                            
                            fechax = LblGestion.Caption
                            Cad = lblCodUniSol.Caption
                            RegPuntero = rs_grdPrincipal!CODIGO_GRUPO
                            nro_reg = rs_grdLiquida!NUMERO_PAGO
                            ' se posiciona en el registro editado
                            Call pl_RefrescaListaPrincipal
                            rs_grdPrincipal.Find "ges_gestion = '" & fechax & "'"
                            rs_grdPrincipal.Find "unidad_codigo ='" & Cad & "'"
                            rs_grdPrincipal.Find "codigo_grupo =" & RegPuntero
                            Call grdPrincipal_RowColChange(0, 0)
                            rs_grdLiquida.Find "numero_pago =" & nro_reg
                            Call grdLiquida_RowColChange(0, 0)
                            MsgBox "Se desaprobo la liquidación correpondiente.", vbInformation, "Aviso"
'                            grdLiquida.SetFocus
                          Else
                        
                            SQLs = "UPDATE ro_pagos_cronograma SET estado_codigo ='REG', usr_aprueba = '" & glusuario & "', fecha_aprueba = '" & fechax & "', hora_aprueba = '" & horax & "' "
                            SQLs = SQLs & "WHERE ges_gestion = '" & LblGestion.Caption & "' and unidad_codigo = '" & lblCodUniSol.Caption & "' and planilla_codigo = " & Val(lblCodGrupo.Caption) & " and numero_pago = " & rs_grdLiquida!NUMERO_PAGO
                            'JQ QR
                            'DE.dbo_apGeneralSearching SQLs
                            
                            RegPuntero = rs_grdLiquida!NUMERO_PAGO
                            Call grdPrincipal_RowColChange(0, 0)
                            rs_grdLiquida.Find "numero_pago =" & RegPuntero ' se posiciona en el registro editado
                            Call grdLiquida_RowColChange(0, 0)
                            MsgBox "Se desaprobo la liquidación correspondiente", vbInformation, "Aviso"
                        
                        End If
                    End If
                    
                    grdBeneficiario.SetFocus
                    Screen.MousePointer = vbDefault
                  Else
                    MsgBox "No existe registro de beneficiario para ser procesado.", vbInformation, "Aviso"
                End If
              Else
                MsgBox "No existe registro de liquidación para ser procesado.", vbInformation, "Aviso"
            End If
          
          Case "Tool_Devengar"
            
            If fl_VerificaDevengar Then
                Screen.MousePointer = vbHourglass
                Cad = "Se devengará la liquidación correspondiente a: " & Chr(13)
                Cad = Cad & "Número de liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "] [" & rs_grdLiquida!Concepto & "]" & Chr(13)
                Cad = Cad & "Beneficiarios: [" & rs_grdBeneficiario.RecordCount & "]" & Chr(13)
                Cad = Cad & "Desea devengar la liquidación?."
                
                If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de aprobación") Then
                    
                    ' genera el devengado del comprometido para todo el pago
                    Call pl_GeneraDevengado
                    
                    RegPuntero = rs_grdLiquida!NUMERO_PAGO
                    Call grdPrincipal_RowColChange(0, 0)
                    rs_grdLiquida.Find "numero_pago =" & RegPuntero ' se posiciona en el registro editado
                    Call grdLiquida_RowColChange(0, 0)
                    
                    If rs_grdLiquida!estado_devengado = "APR" Then
                        MsgBox "Se ha generado el Devengado del Comprometido con todo éxito", vbInformation, "Aviso"
                    End If
                    
                 End If
                 grdBeneficiario.SetFocus
                 Screen.MousePointer = vbDefault
            End If
          
          Case "Tool_AnulaDevengar"
            
            If rs_grdPrincipal.RecordCount > 0 Then
                If fl_VerificaAnulaOrdLiq Then
                    Screen.MousePointer = vbHourglass
                    Cad = "Se anulará la orden de liquidación y el devengado correspondiente a: " & Chr(13)
                    Cad = Cad & "Número de liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "] [" & rs_grdLiquida!Concepto & "]" & Chr(13)
                    Cad = Cad & "Beneficiarios: [" & rs_grdBeneficiario.RecordCount & "]" & Chr(13)
                    Cad = Cad & "Desea anular la orden de liquidación y su devengado correspondiente?."
                    
                    If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de aprobación") Then
                        
                        ' anula la orden de liquidación y el devengado del comprometido para todo el pago
                        'JQ QR
                        'DE.dbo_ap_PagosAnulaOrdLiq_c lblGestion.Caption, lblCodUniSol.Caption, Val(lblCodGrupo.Caption), rs_grdLiquida!NUMERO_PAGO, rs_grdLiquida!correlativo_reg
                        RegPuntero = rs_grdLiquida!NUMERO_PAGO
                        Call pl_RefrescaLiquidacion
                        MsgBox "Se anulo la Orden de Liquidación y el Devengado del Comprometido con todo éxito.", vbInformation, "Aviso"
                     
                     End If
                     grdBeneficiario.SetFocus
                     Screen.MousePointer = vbDefault
                End If
              Else
                MsgBox "No existen registros para procesar.", vbInformation, "Aviso"
            End If
            
          Case "Tool_Salir"
            Unload Me
            
        End Select
    
      ' ********************************************************
      ' opciones genericas de: BENEFICIARIOS
      ' ********************************************************
      
      Case "Beneficiario" ' procesa la cuadro BENEFICIARIOS
        Select Case TipoOpcion
          Case "Tool_Contrato"
            If rs_grdBeneficiario.RecordCount > 0 Then
                    
                ' ***********************
                ' se llama al formulario
                ' ***********************
                Screen.MousePointer = vbHourglass
                grdBeneficiario.SetFocus
                lblEstadoBeneficiario.Caption = rs_grdBeneficiario!Nombre & " " & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno  ' codigo de beneficario usadocomo parametro
                lblEstadoBeneficiario.Tag = rs_grdBeneficiario!codigo_beneficiario ' codigo de beneficario usadocomo parametro
               
                frm_ro_ConfirmaFechasContrato.Show vbModal

              Else
                MsgBox "No existen el registro de beneficiario para ser editado.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If

          Case "Tool_Nuevo"
            If rs_grdLiquida.RecordCount > 0 Then
                ' adiciona un beneficiario al pago si no tiene beneficiarios o modalidad planilla
                If rs_grdBeneficiario.RecordCount = 0 Or rs_grdPrincipal!modalidad_pago = "P" Then
                    lblEstadoBeneficiario.Caption = "REG"
                    '***************************
                    'llama al formulario
                    '***************************
                    Screen.MousePointer = vbHourglass
                    grdBeneficiario.Tag = rs_grdLiquida!NUMERO_PAGO ' guarda el numero de pago como paremtro
                    lblEstadoBeneficiario.Tag = rs_grdLiquida!correlativo_reg
                    frm_ro_SelecBenLiquida.Show vbModal
                    
                    If lblEstadoBeneficiario.Caption <> "REG" Then ' es distinto se guardo el codigo de beneficiario
                        RegPuntero = rs_grdLiquida!NUMERO_PAGO ' se guarda para luego ser ubicado
                        Cad = lblEstadoBeneficiario.Caption ' se gurada el codigo beneficiario
                        
                        Call pl_RefrescaLiquidacion
                        rs_grdLiquida.Find "numero_pago = " & RegPuntero
                        Call grdLiquida_RowColChange(0, 0)
                        rs_grdBeneficiario.Find "codigo_beneficiario = '" & Cad & "'"  ' se posisicona en el registro editado
                        MsgBox "Se adiciono beneficiario(s).", vbInformation, "Aviso"
                        grdBeneficiario.SetFocus
                        
                    End If
                    
                    lblEstadoBeneficiario.Caption = "" ' no se esta editando ni adicionando registros
                  
                  Else
                    MsgBox "La modalidad de liquidación es planilla individual.", vbInformation, "Aviso"
                    grdLiquida.SetFocus
                    
                End If
              
              Else
                MsgBox "No existen registrado número de Liquidación.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
            
          Case "Tool_Monto"
            
            If rs_grdBeneficiario.RecordCount > 0 Then
            
                ' verifica si el comprobante de pago existe y esta aprobado por tesoreria
                Set rstTemp = New ADODB.Recordset
                SQLs = "SELECT r.aprobotesoreria FROM ac_ben_comprdeven r  WHERE r.codigo_beneficiario = '" & rs_grdBeneficiario!codigo_beneficiario & "' and "
                SQLs = SQLs & "r.gp_unidad_codigo = '" & lblCodUniSol.Caption & "' and "
                SQLs = SQLs & "r.gp_codigo_grupo = '" & lblCodGrupo.Caption & "' and "
                SQLs = SQLs & "r.ges_Gestion     = '" & LblGestion.Caption & "' and "
                SQLs = SQLs & "r.tipocomprobante = 'COM' AND "
                SQLs = SQLs & "r.APROBOTESORERIA IN ('N','APR') "

                rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
                
                If rstTemp.RecordCount > 0 Then
                    If rstTemp!APROBOTESORERIA = "APR" Then ' EN AC_BEN_COMPRDEVEN
                        
                        ' ***********************
                        ' se llama al formulario
                        ' ***********************
                        Screen.MousePointer = vbHourglass
                        grdBeneficiario.SetFocus
                        lblEstadoBeneficiario.Caption = "E" ' se esta en modo de edicion del reg. actual
                        grdLiquida.Tag = rs_grdLiquida!correlativo_reg ' numero correlativo
                        grdBeneficiario.Tag = rs_grdLiquida!NUMERO_PAGO ' numero de pago
                        lblEstadoBeneficiario.Tag = rs_grdBeneficiario!codigo_beneficiario ' codigo de beneficario usadocomo parametro

                        frm_ro_LiquidaMontoBen.Show vbModal

                        If lblEstadoBeneficiario.Caption <> "E" Then  ' si el proceso realizo alguna modificacion
                            RegPuntero = rs_grdLiquida!NUMERO_PAGO ' nujero de pago
                            Cad = rs_grdBeneficiario!codigo_beneficiario ' codigo
                            pl_RefrescaLiquidacion
                            rs_grdLiquida.Find "numero_pago =" & RegPuntero
                            Call grdLiquida_RowColChange(0, 0)
                            rs_grdBeneficiario.Find "codigo_beneficiario = '" & Cad & "'"  ' se posisicona en el registro editado
                            
                            MsgBox "Los cambios se realizaron satisfactoriamente", vbInformation, "Aviso"
                            grdBeneficiario.SetFocus
                                                    
                        End If
                        lblEstadoBeneficiario.Caption = "" ' no se esta editando ni adicionando registros
                        
                      Else
                        MsgBox "El comprobante de liquidación correspondiente a [" & rs_grdBeneficiario!Nombre & " " & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & "]." & Chr(13) & " No se encuentra aprobado en presupuestos." & Chr(13) & Chr(13) & "Verifique el proceso....gracias.", vbCritical, "Error"
                        grdBeneficiario.SetFocus
                    End If
                  Else
                    MsgBox "No existen registro de comprobante de liquidación correspondiente a [" & rs_grdBeneficiario!Nombre & " " & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & "]." & Chr(13) & " para ser procesado." & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
                    grdBeneficiario.SetFocus
                End If
              Else
                MsgBox "No existen el registro de beneficiario para ser editado.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
            
          Case "Tool_Anular"
            
            If rs_grdBeneficiario.RecordCount > 0 Then ' si existen registros
                If rs_grdLiquida!estado_devengado & "" <> "APR" Then
                    Screen.MousePointer = vbHourglass
                    Cad = "Se eliminará el registro del beneficiario [" & rs_grdBeneficiario!Nombre & " " & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & "] correspondiente a:" & Chr(13)
                    Cad = Cad & "Unidad: [" & lblCodUniSol.Caption & " - " & lblDesUniSol.Caption & "]" & Chr(13) & "Grupo: [" & lblCodGrupo.Caption & " - " & lblDesGrupo.Caption & "]" & Chr(13) & "Nro. liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "]." & Chr(13)
                    If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de elimnación") Then
                        'JQ QR
                        'DE.dbo_ap_PagosBorraPagoBenef rs_grdPrincipal!ges_gestion, rs_grdPrincipal!unidad_codigo, rs_grdPrincipal!codigo_grupo, rs_grdLiquida!NUMERO_PAGO, rs_grdBeneficiario!codigo_beneficiario
                        
                        RegPuntero = rs_grdLiquida!NUMERO_PAGO
                        Call pl_RefrescaLiquidacion
                        rs_grdLiquida.Find "numero_pago =" & RegPuntero
                        Call grdLiquida_RowColChange(0, 0)
                        
                        Call pl_RefrescaBeneficiario
                        
                        MsgBox "La información se eliminó satisfactoriamente", vbInformation, "Aviso"
                    End If
                    Screen.MousePointer = vbDefault
                  Else
                    MsgBox "No puede eliminar el registro del beneficiario [" & rs_grdBeneficiario!Nombre & " " & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & "] por que tiene devengado generado.", vbInformation, "Aviso"
                    grdBeneficiario.SetFocus
                End If
              Else
                MsgBox "No existen registros para ser eliminados.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
          
          Case "Tool_AnularTodo"
            
            If rs_grdBeneficiario.RecordCount > 0 Then ' si existen registros
                If rs_grdLiquida!estado_devengado & "" <> "APR" Then
                    Screen.MousePointer = vbHourglass
                    Cad = "Se eliminará TODOS los registros de beneficiarios correspondiente a:" & Chr(13)
                    Cad = Cad & "Unidad: [" & lblCodUniSol.Caption & " - " & lblDesUniSol.Caption & "]" & Chr(13) & "Grupo: [" & lblCodGrupo.Caption & " - " & lblDesGrupo.Caption & "]" & Chr(13) & "Nro. liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "]." & Chr(13)
                    If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de elimnación") Then
                        'JQ QR
                        'DE.dbo_apGeneralSearching "DELETE ro_pagos_cronograma_detalle WHERE ges_gestion = '" & rs_grdPrincipal!ges_gestion & "' and unidad_codigo   = '" & rs_grdPrincipal!unidad_codigo & "' and codigo_grupo = " & rs_grdPrincipal!codigo_grupo & " and numero_pago = " & rs_grdLiquida!NUMERO_PAGO
                        'DE.dbo_apGeneralSearching "UPDATE ro_pagos_cronograma set tipo_moneda = '', monto_us = 0, monto_bs = 0 WHERE ges_gestion = '" & rs_grdPrincipal!ges_gestion & "' and unidad_codigo   = '" & rs_grdPrincipal!unidad_codigo & "' and codigo_grupo = " & rs_grdPrincipal!codigo_grupo & " and numero_pago = " & rs_grdLiquida!NUMERO_PAGO
                        
                        RegPuntero = rs_grdLiquida!NUMERO_PAGO
                        Call pl_RefrescaLiquidacion
                        rs_grdLiquida.Find "numero_pago =" & RegPuntero
                        Call grdLiquida_RowColChange(0, 0)
                        
                        Call pl_RefrescaBeneficiario
                        
                        MsgBox "La información se eliminó satisfactoriamente", vbInformation, "Aviso"
                    End If
                    Screen.MousePointer = vbDefault
                  Else
                    MsgBox "No puede eliminar los beneficiarios por que tiene devengado generado.", vbInformation, "Aviso"
                    grdBeneficiario.SetFocus
                End If
              Else
                MsgBox "No existen registros para ser eliminados.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
                    
          Case "Tool_Conformidad"
            
            If rs_grdLiquida.RecordCount > 0 Then
                If rs_grdLiquida!estado_codigo <> "APR" Then ' verificamos
                    
                    If rs_grdBeneficiario.RecordCount > 0 Then
                    
                        If fl_VerificaConformidad Then

                            ' *********************************************
                            ' se llama al formulario q permite resgistrar cite
                            ' *********************************************
                            Screen.MousePointer = vbHourglass
                            lblEstadoBeneficiario.Caption = "C" ' parametro para registrar conformidad
                            grdBeneficiario.Tag = rs_grdLiquida!NUMERO_PAGO
                            frm_ro_LiquidaConformidad.Show vbModal
                            lblEstadoBeneficiario.Caption = ""
                            
                            Call grdLiquida_RowColChange(0, 0)
                            grdLiquida.SetFocus
                        End If
                      Else
                        MsgBox "No se tiene registro de beneficiarios.", vbInformation, "Aviso"
                        grdLiquida.SetFocus
                    End If
                  Else
                    MsgBox "La liquidación número [" & rs_grdLiquida!NUMERO_PAGO & "] se encuentra aprobada.", vbInformation, "Aviso"
                    grdPrincipal.SetFocus
                End If
              Else
                MsgBox "No existe registro de liquidación para registrar conformidad.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
                
            End If
            
          Case "Tool_Factura"
            
            If rs_grdLiquida.RecordCount > 0 Then
                If rs_grdLiquida!estado_codigo <> "APR" Then ' verificamos si tiene comprobante presupuestario
                    If rs_grdBeneficiario.RecordCount > 0 Then
                    
                        ' *********************************************
                        ' se llama al formulario q permite resgistrar si emite factura
                        ' *********************************************
                        Screen.MousePointer = vbHourglass
                        lblEstadoBeneficiario.Caption = "F" 'parametro para procesar registro de emite o no factura
                        grdBeneficiario.Tag = rs_grdLiquida!NUMERO_PAGO
                        frm_ro_LiquidaConformidad.Show vbModal
                        lblEstadoBeneficiario.Caption = ""
                        
                        Call grdLiquida_RowColChange(0, 0)
                        grdPrincipal.SetFocus
                      Else
                        MsgBox "No se tiene registro de beneficiarios.", vbInformation, "Aviso"
                        grdLiquida.SetFocus
                    End If
                  Else
                    MsgBox "La liquidación número [" & rs_grdLiquida!NUMERO_PAGO & "] se encuentra aprobada.", vbInformation, "Aviso"
                    grdPrincipal.SetFocus
                  
                End If
                
              Else
                MsgBox "No existe registro de liquidación para registrar si emite factura.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
                
            End If
          
          
          Case "Tool_Salir"
            Unload Me
            
        End Select
    
    End Select
      
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    ' si se produjo otro tipo de error
    MsgBox "Error: Se produjo un error." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Function fl_VerificaEliminaGrupo() As Boolean
    'TITULO:                Función fl_VerificaEliminaGrupo
    'PROPOSITO:             Verifica los datos para procesar la elimnacion
    'EJEMPLO DE LLAMADA:    fl_VerificaEliminaGrupo
    Dim rstTemp As ADODB.Recordset ' usado para la carga de los combos de base
    
    fl_VerificaEliminaGrupo = True
    
    ' verifica si tiene compromisos de pago aprobados
    SQLs = "select * from ac_ben_comprDeven where gp_ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and gp_unidad_codigo ='" & rs_grdPrincipal!unidad_codigo & "' and gp_codigo_grupo =" & rs_grdPrincipal!CODIGO_GRUPO & " and tipoComprobante ='COM' and aprobotesoreria='APR'"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        MsgBox "No puede eliminar el grupo [" & lblCodGrupo.Caption & "][" & lblDesGrupo.Caption & "] de liquidación por tener compromiso de pago APROBADO." & Chr(13) & "Comuniquese con el administrador del sistema.", vbInformation, "Aviso"
        grdPrincipal.SetFocus
        fl_VerificaEliminaGrupo = False
        Exit Function
    End If
        
    ' verificamos si tiene algun pago devengado
    SQLs = "select * from ro_pagos_cronograma where ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and unidad_codigo='" & rs_grdPrincipal!unidad_codigo & "' and codigo_grupo=" & rs_grdPrincipal!CODIGO_GRUPO & " and estado_devengado ='APR'"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    If rstTemp.RecordCount > 0 Then
        MsgBox "No puede eliminar el grupo [" & lblCodGrupo.Caption & "][" & lblDesGrupo.Caption & "] porque tiene ordenes de pago elaboradas." & Chr(13) & "Comuniquese con el administrador del sistema.", vbInformation, "Aviso"
        grdPrincipal.SetFocus
        fl_VerificaEliminaGrupo = False
        Exit Function
    End If
    
    Set rstTemp = Nothing
    
End Function

Private Sub pl_HabilitaUnaOpcion(Boton As String, swModo As Boolean, Proceso As String)
    'TITULO:                Procedimiento pl_HabilitaUnaOpcion
    'PROPOSITO:             Habilita o deshabilita el boton especificado del toolbar
    'EJEMPLO DE LLAMADA:    call pl_HabilitaUnaOpcion(NombreBoton, true/false, Proceso)
    
    Select Case Proceso
      Case "GrupoLiq"
        ' habilitamos o deshabilitamos las opciones del menu
        Select Case Boton
          Case "Tool_Nuevo"
            BtnAñadir.Enabled = swModo
          Case "Tool_Editar"
            BtnModificar.Enabled = swModo
          Case "Tool_Anular"
            BtnEliminar.Enabled = swModo
          Case "Tool_Salir"
            BtnSalir.Enabled = swModo
        End Select
      
      Case "Liquida"
        ' habilitamos o deshabilitamos las opciones del menu
        Select Case Boton
          Case "Tool_Nuevo"
            cmdPagoNuevo.Enabled = swModo
          Case "Tool_Editar"
            cmdPagoEditar.Enabled = swModo
          Case "Tool_Anular"
            cmdPagoAnular.Enabled = swModo
          Case "Tool_Aprobar"
            cmdPagoAprob.Enabled = swModo
          Case "Tool_Desaprobar"
            cmdPagoDesaprob.Enabled = swModo
          Case "Tool_Devengar"
            cmdPagoDeven.Enabled = swModo
          Case "Tool_AnulaDevengar"
            cmdPagoAnulaDev.Enabled = swModo
        End Select
    
      Case "Beneficiario"
        ' habilitamos o deshabilitamos las opciones del menu
        Select Case Boton
          Case "Tool_Nuevo"
            cmdBenNuevo.Enabled = swModo
          Case "Tool_Monto"
            cmdBenMonto.Enabled = swModo
          Case "Tool_Anular"
            cmdBenAnular.Enabled = swModo
          Case "Tool_AnularTodo"
            cmdBenAnularTodo.Enabled = swModo
          Case "Tool_Conformidad"
            cmdBenConf.Enabled = swModo
          Case "Tool_Factura"
            cmdBenFact.Enabled = swModo
        End Select
    
    End Select
End Sub

Private Sub grdBeneficiario_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'Call pl_ControlaToolBar("Beneficiario")
End Sub

Private Sub grdPrincipal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'PROPOSITO:             Permite desplazarse sobre el browse actualizando los datos de las fichas
    
    On Error GoTo EtiqError

    ' datos cabecera
    If rs_grdPrincipal.RecordCount = 0 Then
        LblGestion.Caption = "" ' gestion
        lblFormulario.Caption = "" ' tipo de trammite
        lblCodUniSol.Caption = "" ' unidad solicitante
        lblDesUniSol.Caption = "" ' unidad solicitante
        lblCodGrupo.Caption = "" ' codigo grupo
        lblDesGrupo.Caption = "" ' descripcion grupo
        lblCodGrupo.Tag = "" ' para guardar el numero de consultoria
        lblDesGrupo.Tag = "" ' para guardar el tipo de liqudacion planialla o individual
      Else
        LblGestion.Caption = rs_grdPrincipal!ges_gestion & ""    ' gestion
        'lblFormulario.Caption = rs_grdPrincipal!formulario & " - " & rs_grdPrincipal!Denominacion_Tipo    ' tipo de tramite
        lblCodUniSol.Caption = rs_grdPrincipal!unidad_codigo  ' unidad solicitante
        'lblDesUniSol.Caption = rs_grdPrincipal!Uni_descripcion_larga   ' unidad solicitante
        'lblCodGrupo.Caption = rs_grdPrincipal!planilla_codigo & "" ' codigo grupo
        'lblDesGrupo.Caption = rs_grdPrincipal!descripcion_grupo & "" ' descripcion grupo
        'lblCodGrupo.Tag = rs_grdPrincipal!numero_consultoria & "" ' numero conultoria
        'lblDesGrupo.Tag = rs_grdPrincipal!modalidad_pago & "" ' modalidad de pago
        
        lblDesGrupo.Caption = rs_grdLiquida!NUMERO_PAGO & ""
        'lblGestion.Caption = rs_grdLiquida!ges_gestion & "' "
        lblCodGrupo.Caption = rs_grdPrincipal!planilla_codigo & " "
        lblFormulario.Caption = rs_grdPrincipal!mes_grupo & " "
        
    End If
    
    ' GRUPO LIQUIDACION
    Call pl_RefrescaLiquidacion
    Call pl_ControlaToolBar("GrupoLiq")
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub pl_RefrescaLiquidacion()
    'TITULO:                Procedimiento pl_RefrescaLiquidacion
    'PROPOSITO:             Actualiza los datos de la ficha
    'EJEMPLO DE LLAMADA:    call pl_RefrescaLiquidacion
    
    On Error GoTo EtiqError
        
    ' se actualiza el grid
'    SQLS = "SELECT numero_pago, concepto, tipo_moneda, monto_us, monto_bs, estado_codigo, estado_devengado, antecedente, codigo_orden, fecha_estimada_liq, correlativo_reg "
'    SQLS = SQLS & "FROM ro_pagos_cronograma "
'    SQLS = SQLS & "WHERE ges_gestion = '" & lblGestion.Caption & "' AND unidad_codigo ='" & lblCodUniSol.Caption & "' AND codigo_grupo =" & Val(lblCodGrupo.Caption)
'    SQLS = SQLS & " and estado_devengado <> 'E'"
    If LblGestion.Caption = "" Then
        LblGestion.Caption = "2015"
    End If
    If lblCodGrupo.Caption = "" Then
        lblCodGrupo.Caption = "P01"
    End If
    SQLs = "Select * FROM ro_pagos_cronograma WHERE ges_gestion = '" & LblGestion.Caption & "' AND planilla_codigo ='" & lblCodGrupo.Caption & "' AND mes_grupo =" & Ado_datos.Recordset!mes_grupo & " "         'and numero_pago = " & m & " "
    Set rs_grdLiquida = New ADODB.Recordset
    rs_grdLiquida.Open SQLs, db, adOpenStatic, adLockReadOnly
    Set AdoLiquida.Recordset = rs_grdLiquida.DataSource
    Set grdLiquida.DataSource = AdoLiquida.Recordset
    If AdoLiquida.Recordset.RecordCount > 0 Then
        lblDesGrupo.Caption = rs_grdLiquida!NUMERO_PAGO & ""
        LblGestion.Caption = rs_grdLiquida!ges_gestion & "' "
        lblCodGrupo.Caption = rs_grdLiquida!planilla_codigo & " "
        lblFormulario.Caption = rs_grdLiquida!mes_grupo & " "
    Else
        lblDesGrupo.Caption = 0
        LblGestion.Caption = "2015"
        lblCodGrupo.Caption = "P01"
        lblFormulario.Caption = 0
    End If
    grdLiquida.Caption = "Liquidación del grupo: [" & Val(lblCodGrupo.Caption) & " - " & IIf(rs_grdPrincipal.RecordCount = 0, "", rs_grdPrincipal!descripcion_grupo) & "]."
    lblNroLiquida.Caption = "Nro. de liquidaciones: " & rs_grdLiquida.RecordCount
    
    'Call pl_PersonalizaGridLiquida

    ' calcula totales de liquidación
'    SQLs = "SELECT 'total_us' = sum(monto_us), 'total_bs' = sum(monto_bs) "
'    SQLs = SQLs & "FROM ro_pagos_cronograma "
'    SQLs = SQLs & "WHERE ges_gestion = '" & lblGestion.Caption & "' AND unidad_codigo ='" & lblCodUniSol.Caption & "' AND planilla_codigo =" & Val(lblCodGrupo.Caption)
'    SQLs = SQLs & " and estado_devengado <> 'E'"
    
    SQLs = "Select sum(monto_us) as total_us, sum(monto_bs) as total_bs "
    SQLs = SQLs & " FROM ro_pagos_cronograma "
    SQLs = SQLs & " WHERE ges_gestion = '" & LblGestion.Caption & "' AND planilla_codigo = '" & lblCodGrupo.Caption & "' AND mes_grupo =" & Ado_datos.Recordset!mes_grupo & " "
    
    'SQLs = "Select sum(monto_us) as total_us, sum(monto_bs) as total_bs  FROM ro_pagos_cronograma WHERE ges_gestion = '" & lblGestion.Caption & "' AND planilla_codigo = '" & lblCodGrupo.Caption & "' AND mes_grupo =" & Ado_datos.Recordset!mes_grupo & " "         'and numero_pago = " & m & " "
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    'lblTotalUSLiq = Format(IIf(IsNull(rstTemp!total_us), 0, rstTemp!total_us), "##,##0.00") & " $US"
    'lblTotalBSLiq = Format(IIf(IsNull(rstTemp!total_bs), 0, rstTemp!total_bs), "##,##0.00") & " Bs"
    
    Call pl_RefrescaBeneficiario
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub pl_PersonalizaGridLiquida()
    'TITULO:                Procedimiento pl_PersonalizaGridLiquida
    'PROPOSITO:             Personalizar los captions, anchos, etc. del grid
    'EJEMPLO DE LLAMADA:    call pl_PersonalizaGridLiquida
        
    Dim i As Integer
    
    ' define ancho de columnas y titulo de la cabecera
    grdLiquida.Columns(0).Width = 900 ' numero de liq.
    grdLiquida.Columns(0).Caption = "Nro. liquid."
    grdLiquida.Columns(1).Width = 2000 ' concepto
    grdLiquida.Columns(1).Caption = "Concepto"
    grdLiquida.Columns(2).Width = 1000 ' tipo moneda
    grdLiquida.Columns(2).Caption = "Tipo moneda"
    grdLiquida.Columns(3).Width = 1000 ' monto us
    grdLiquida.Columns(3).Caption = "Monto US"
    grdLiquida.Columns(4).Width = 1000 ' monto BS
    grdLiquida.Columns(4).Caption = "Monto BS"
    grdLiquida.Columns(5).Width = 900 ' estado aprobado
    grdLiquida.Columns(5).Caption = "Est. aprobado"
    grdLiquida.Columns(6).Width = 900 ' estado devengado
    grdLiquida.Columns(6).Caption = "Est. devengado"
    grdLiquida.Columns(7).Width = 900 ' antecedente
    grdLiquida.Columns(7).Caption = "Antecedente"
    grdLiquida.Columns(8).Width = 1000 ' codigo orden
    grdLiquida.Columns(8).Caption = "Cod. orden"
    grdLiquida.Columns(9).Width = 1000 ' fecha estimada liquidacion
    grdLiquida.Columns(9).Caption = "F. estimada liquidación"

End Sub

Private Sub grdLiquida_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'PROPOSITO:             Permite desplazarse sobre el browse actualizando los datos de las fichas

    On Error GoTo EtiqError

    If rs_grdLiquida.RecordCount = 0 Then
        lblEstadoBeneficiario.Tag = ""
        Call pl_RefrescaBeneficiario
'        Call grdBeneficiario_RowColChange(0, 0)
'        Call pl_ControlaToolBar("Liquida")
      Else
        'lblEstadoBeneficiario.Tag = rs_grdLiquida!numero_pago & ""
        lblDesGrupo.Caption = rs_grdLiquida!NUMERO_PAGO & ""
        Call pl_RefrescaBeneficiario
        'Call grdBeneficiario_RowColChange(0, 0)
        'Call pl_ControlaToolBar("Liquida")
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Function fl_VerificaEliminaLiquida() As Boolean
    'TITULO:                Función fl_VerificaEliminaLiquida
    'PROPOSITO:             Verifica los datos para procesar la elimnacion
    'EJEMPLO DE LLAMADA:    fl_VerificaEliminaLiquida
    Dim rstTemp As ADODB.Recordset ' usado para la carga de los combos de base
    
    fl_VerificaEliminaLiquida = True

    If rs_grdLiquida!estado_devengado = "APR" Then
        MsgBox "No puede eliminar la liquidación por estar generado correspondiente devengado." & "", vbInformation, "Aviso"
        grdLiquida.SetFocus
        fl_VerificaEliminaLiquida = False
        Exit Function
    End If
    
    If rs_grdLiquida!estado_codigo = "APR" Then
        MsgBox "No puede eliminar la liquidación por estar aprobado." & "", vbInformation, "Aviso"
        grdLiquida.SetFocus
        fl_VerificaEliminaLiquida = False
        Exit Function
    End If
    
    ' verifica si existe pagos superiores
    SQLs = "SELECT * FROM ro_pagos_cronograma "
    SQLs = SQLs & "WHERE ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and unidad_codigo ='" & rs_grdPrincipal!unidad_codigo & "' and codigo_grupo =" & rs_grdPrincipal!CODIGO_GRUPO & " and numero_pago> " & rs_grdLiquida!NUMERO_PAGO & " and estado_codigo <>'E'and estado_devengado <>'E'"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        MsgBox "No puede eliminar la liquidación [" & rs_grdLiquida!NUMERO_PAGO & "] del grupo [" & lblCodGrupo.Caption & " - " & lblDesGrupo.Caption & "] por tener registro de pagos superiores." & Chr(13) & "Corrija el error e intente eliminar nuevamente.", vbInformation, "Aviso"
        grdLiquida.SetFocus
        fl_VerificaEliminaLiquida = False
        Exit Function
    End If
    
    Set rstTemp = Nothing
    
End Function


''*******************************************************
''procesos de la ficha BENEFICIARIOS
''*******************************************************

Private Sub pl_RefrescaBeneficiario()
    'TITULO:                Procedimiento pl_RefrescaBeneficiario
    'PROPOSITO:             Actualiza los datos de la ficha beneficiario
    'EJEMPLO DE LLAMADA:    call pl_RefrescaBeneficiario
    
    On Error GoTo EtiqError
    
    ' obtiene datos de beneficiarios del pago
    ' dependiendo del tipo de proceso si es consultor por F05 ==> "producto - corto plazo"  o F10 ==> consultor por "tiempo - largo pazo"
'    Select Case glProceso
'      Case "F05"
'        SQLs = "SELECT ro_pagos_cronograma_detalle.*, gc_beneficiario.beneficiario_denominacion"
'        SQLs = SQLs & "FROM ro_pagos_cronograma_detalle INNER JOIN gc_beneficiario ON ro_pagos_cronograma_detalle.beneficiario_codigo = gc_beneficiario.beneficiario_codigo"
'        SQLs = SQLs & "WHERE ro_pagos_cronograma_detalle.ges_gestion = '" & lblGestion.Caption & "'  AND ro_pagos_cronograma_detalle.planilla_codigo = '" & Val(lblCodGrupo.Caption) & "' "
'        SQLs = SQLs & "AND ro_pagos_cronograma_detalle.mes_grupo  = " & Ado_datos.Recordset!mes_grupo & "  AND ro_pagos_cronograma_detalle.numero_pago = " & ro_pagos_cronograma_detalle.numero_pago & " "
'        SQLs = SQLs & "ORDER BY gc_beneficiario.beneficiario_denominacion"
'      Case "F10"
'        'SQLs = SQLs & " and ro_pagos_cronograma_detalle.estado_devengado <> 'E' "
    'If AdoBeneficiario.Recordset.RecordCount > 0 Then
        SQLs = "SELECT ro_pagos_cronograma_detalle.*, gc_beneficiario.beneficiario_denominacion"
        SQLs = SQLs & " FROM ro_pagos_cronograma_detalle INNER JOIN gc_beneficiario ON ro_pagos_cronograma_detalle.beneficiario_codigo = gc_beneficiario.beneficiario_codigo "
        SQLs = SQLs & " WHERE ro_pagos_cronograma_detalle.ges_gestion = '" & LblGestion.Caption & "'  AND ro_pagos_cronograma_detalle.planilla_codigo = '" & lblCodGrupo.Caption & "' "
        SQLs = SQLs & " AND ro_pagos_cronograma_detalle.mes_grupo  = " & Ado_datos.Recordset!mes_grupo & "  AND ro_pagos_cronograma_detalle.numero_pago = " & lblDesGrupo.Caption & " "       '" & AdoLiquida.Recordset!numero_pago & " "
        SQLs = SQLs & " ORDER BY gc_beneficiario.beneficiario_denominacion "
'    End Select
        Set rs_grdBeneficiario = New ADODB.Recordset
        rs_grdBeneficiario.Open SQLs, db, adOpenStatic, adLockReadOnly
        'Set grdBeneficiario.DataSource = rs_grdBeneficiario
        Set AdoBeneficiario.Recordset = rs_grdBeneficiario.DataSource
        Set grdBeneficiario.DataSource = AdoBeneficiario.Recordset
    
        grdBeneficiario.Caption = "Beneficiarios del pago numero:[" & IIf(rs_grdLiquida.RecordCount = 0, 0, rs_grdLiquida!NUMERO_PAGO) & "]."
    
'        Call pl_PersonalizaGridBeneficiario
    
        lblNroBeneficiario.Caption = "Nro. de beneficiarios: " & rs_grdBeneficiario.RecordCount
    'Else
    '    lblNroBeneficiario.Caption = "Nro. de beneficiarios: 0"
    'End If
    ' calculamos montos totales por el numero de liquidacion
    
'    SQLs = "SELECT 'total_US' = SUM(ro_pagos_cronograma_detalle.monto_us), 'total_BS' = SUM(ro_pagos_cronograma_detalle.monto_bs) "
'    SQLs = SQLs & "FROM ro_pagos_cronograma_detalle "
'    SQLs = SQLs & "WHERE ro_pagos_cronograma_detalle.ges_gestion = '" & lblGestion.Caption & "' "
'    SQLs = SQLs & " AND ro_pagos_cronograma_detalle.planilla_codigo = " & Val(lblCodGrupo.Caption)
'    SQLs = SQLs & " AND ro_pagos_cronograma_detalle.unidad_codigo = '" & lblCodUniSol.Caption & "' "
'    SQLs = SQLs & " AND ro_pagos_cronograma_detalle.numero_pago = " & IIf(rs_grdLiquida.RecordCount = 0, 9999, rs_grdLiquida!numero_pago) & " and ro_pagos_cronograma_detalle.correlativo_reg = " & IIf(rs_grdLiquida.RecordCount = 0, 999, rs_grdLiquida!correlativo_reg)
'    SQLs = SQLs & " and ro_pagos_cronograma_detalle.estado_devengado <> 'E' "

    SQLs = "SELECT SUM(liquido_pagable_us) as total_US, SUM(liquido_pagable_bs) as total_BS"
    SQLs = SQLs & " FROM ro_pagos_cronograma_detalle "
    SQLs = SQLs & " WHERE ges_gestion = '" & LblGestion.Caption & "'  AND planilla_codigo = '" & lblCodGrupo.Caption & "' "
    SQLs = SQLs & " AND mes_grupo  = " & Ado_datos.Recordset!mes_grupo & "  AND numero_pago = " & lblDesGrupo.Caption & " "

    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    lblTotalUS.Caption = Format(IIf(IsNull(rstTemp!total_us), 0, rstTemp!total_us), "######0.00") & " $US" ' total asignado pie de grid
    lblTotalBS.Caption = Format(IIf(IsNull(rstTemp!total_bs), 0, rstTemp!total_bs), "######0.00") & " Bs" ' total asignado pie de grid
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    'MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub pl_PersonalizaGridBeneficiario()
    'TITULO:                Procedimiento pl_PersonalizaGridBeneficiario
    'PROPOSITO:             Personalizar los captions, anchos, etc. del grid
    'EJEMPLO DE LLAMADA:    call pl_PersonalizaGridBeneficiario

    ' define ancho de columnas y titulo de la cabecera
    grdBeneficiario.Columns(0).Width = 400 ' Nro.
    grdBeneficiario.Columns(0).Caption = "No.Liq."
    grdBeneficiario.Columns(1).Width = 1000 ' paterno
    grdBeneficiario.Columns(1).Caption = "Paterno"
    grdBeneficiario.Columns(2).Width = 1000 ' materno
    grdBeneficiario.Columns(2).Caption = "Materno"
    grdBeneficiario.Columns(3).Width = 1200 ' nombres
    grdBeneficiario.Columns(3).Caption = "Nombre(s)"
    grdBeneficiario.Columns(4).Width = 900 ' codigo ben
    grdBeneficiario.Columns(4).Caption = "Cod. Benef."
    grdBeneficiario.Columns(5).Width = 1000 ' monto us
    grdBeneficiario.Columns(5).Caption = "Monto US"
    grdBeneficiario.Columns(6).Width = 1000 ' monto bs
    grdBeneficiario.Columns(6).Caption = "Monto BS"
    grdBeneficiario.Columns(7).Width = 500 ' tc us
    grdBeneficiario.Columns(7).Caption = "Tc US"
    grdBeneficiario.Columns(8).Width = 600 ' moneda
    grdBeneficiario.Columns(8).Caption = "Moneda"
    grdBeneficiario.Columns(9).Width = 500 ' emite factura
    grdBeneficiario.Columns(9).Caption = "Emite factura"
    grdBeneficiario.Columns(10).Width = 900 ' estado conformidad
    grdBeneficiario.Columns(10).Caption = "Est. conformidad"
    grdBeneficiario.Columns(11).Width = 900 ' estado devengado
    grdBeneficiario.Columns(11).Caption = "Est. devengado"
    grdBeneficiario.Columns(12).Width = 1200 ' nro cite
    grdBeneficiario.Columns(12).Caption = "Nro CITE"
    grdBeneficiario.Columns(13).Width = 1000 ' F. CITE
    grdBeneficiario.Columns(13).Caption = "F. CITE"
    grdBeneficiario.Columns(14).Width = 1000 ' Nro. con. hist
    grdBeneficiario.Columns(14).Caption = "Nro. consul. hist."
    grdBeneficiario.Columns(15).Width = 1000 ' fte. financ hist
    grdBeneficiario.Columns(15).Caption = "Fte. financ.hist."

End Sub

Private Sub pl_ControlaToolBar(Proceso As String)
    'TITULO:                Procedimiento ControlaBotones
    'PROPOSITO:             Permite controlar botones habilitando /deshabilitando para el tipo de proceso q siga

    On Error GoTo EtiqError
    
    If rs_grdPrincipal.RecordCount = 0 Then ' si no existen registros se cancelan todos los botones
        
        Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "GrupoLiq")
        Call pl_HabilitaUnaOpcion("Tool_Editar", False, "GrupoLiq")
        Call pl_HabilitaUnaOpcion("Tool_Anular", False, "GrupoLiq")
        
        Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "Liquida")
        Call pl_HabilitaUnaOpcion("Tool_Editar", False, "Liquida")
        Call pl_HabilitaUnaOpcion("Tool_Anular", False, "Liquida")
        Call pl_HabilitaUnaOpcion("Tool_Aprobar", False, "Liquida")
        Call pl_HabilitaUnaOpcion("Tool_Devengar", False, "Liquida")

        Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "Beneficiario")
        Call pl_HabilitaUnaOpcion("Tool_Monto", False, "Beneficiario")
        Call pl_HabilitaUnaOpcion("Tool_Anular", False, "Beneficiario")
        Call pl_HabilitaUnaOpcion("Tool_AnularTodo", False, "Beneficiario")
        Call pl_HabilitaUnaOpcion("Tool_Conformidad", False, "Beneficiario")
        Call pl_HabilitaUnaOpcion("Tool_Factura", False, "Beneficiario")

        Exit Sub
    End If
    
    Select Case Proceso
      
      ' ********************************************************
      ' controla botones de la: GRUPO DE PAGO
      ' ********************************************************
    
      Case "GrupoLiq" ' botones para procesos de la ficha GRUPO DE PAGO
        
        Select Case LTrim(RTrim(rs_grdPrincipal!estado_codigo)) ' estado aprobación de consultoria
          Case "APR", "A" ' estado de aprobación de la solicitud S=aprobado
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "GrupoLiq")
            Call pl_HabilitaUnaOpcion("Tool_Editar", False, "GrupoLiq")
            Call pl_HabilitaUnaOpcion("Tool_Anular", False, "GrupoLiq")
        
          Case Else ' "", Null = solo solicitado tramite no iniciado
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "GrupoLiq")
            Call pl_HabilitaUnaOpcion("Tool_Editar", True, "GrupoLiq")
            Call pl_HabilitaUnaOpcion("Tool_Anular", True, "GrupoLiq")
                    
        End Select
        
      ' ********************************************************
      ' controla botones de : LIQUIDACION
      ' ********************************************************
        
      Case "Liquida" ' botones para procesos de LIQUIDACION
        
        If rs_grdLiquida.RecordCount = 0 Then
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Editar", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Anular", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Aprobar", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Desaprobar", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Devengar", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_AnulaDevengar", False, "Liquida")
            cmdPagoDeven.Visible = True
            cmdPagoAnulaDev.Visible = False
            cmdPagoAprob.Visible = True
            cmdPagoDesaprob.Visible = False
            Exit Sub
        End If
        
        Select Case LTrim(RTrim(rs_grdLiquida!estado_codigo)) ' estado aprobado
          Case "APR"
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Editar", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Anular", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Aprobar", False, "Liquida")
            
            If LTrim(RTrim(rs_grdLiquida!estado_devengado)) <> "APR" Then ' estado devengado
                Call pl_HabilitaUnaOpcion("Tool_Devengar", True, "Liquida")
                Call pl_HabilitaUnaOpcion("Tool_AnulaDevengar", False, "Liquida")
                Call pl_HabilitaUnaOpcion("Tool_Desaprobar", True, "Liquida")
                cmdPagoDeven.Visible = True
                cmdPagoAnulaDev.Visible = False
                cmdPagoAprob.Visible = False
                cmdPagoDesaprob.Visible = True
            Else
                Call pl_HabilitaUnaOpcion("Tool_Devengar", False, "Liquida")
                
                ' verifica si tiene devengado de pago aprobados
                SQLs = "select * from ac_ben_comprDeven where gp_ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and gp_unidad_codigo ='" & rs_grdPrincipal!unidad_codigo & "' and gp_codigo_grupo =" & rs_grdPrincipal!CODIGO_GRUPO & " and gp_numero_pago = " & rs_grdLiquida!NUMERO_PAGO & " and tipoComprobante ='DEV' and aprobotesoreria='APR'"
                Set rstTemp = New ADODB.Recordset
                rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
                If rstTemp.RecordCount > 0 Then
                    Call pl_HabilitaUnaOpcion("Tool_AnulaDevengar", False, "Liquida")
                  Else
                    Call pl_HabilitaUnaOpcion("Tool_AnulaDevengar", True, "Liquida")
                End If
                
                Call pl_HabilitaUnaOpcion("Tool_Desaprobar", False, "Liquida")
                cmdPagoDeven.Visible = False
                cmdPagoAnulaDev.Visible = True
                cmdPagoAprob.Visible = True
                cmdPagoDesaprob.Visible = False
            End If
            
          Case Else
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Editar", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Anular", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Aprobar", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Desaprobar", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Devengar", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_AnulaDevengar", False, "Liquida")
            cmdPagoDeven.Visible = True
            cmdPagoAnulaDev.Visible = False
            cmdPagoAprob.Visible = True
            cmdPagoDesaprob.Visible = False
        End Select
          
      ' ********************************************************
      ' controla botones de la ficha: BENEFICIARIO
      ' ********************************************************

      Case "Beneficiario" ' botones para procesos de la ficha BENEFICIARIO
        
        If rs_grdLiquida.RecordCount = 0 Then
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Monto", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Anular", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_AnularTodo", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Conformidad", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Factura", False, "Beneficiario")
            
            Exit Sub
        End If
        
        Select Case LTrim(RTrim(rs_grdLiquida!estado_codigo)) ' estado
          Case "APR"
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Monto", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Anular", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_AnularTodo", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Conformidad", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Factura", False, "Beneficiario")
            
          Case Else
            If rs_grdPrincipal!modalidad_pago = "I" And rs_grdBeneficiario.RecordCount > 0 Then
                Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "Beneficiario")
              Else
                Call pl_HabilitaUnaOpcion("Tool_Nuevo", True, "Beneficiario")
            End If
            Call pl_HabilitaUnaOpcion("Tool_Monto", True, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Anular", True, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_AnularTodo", True, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Conformidad", True, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Factura", True, "Beneficiario")

        End Select
          
    End Select
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"


End Sub

Private Function fl_VerificaConformidad() As Boolean
    'TITULO:                Función fl_VerificaConformidad
    'PROPOSITO:             Verifica los datos para registrar la conformidad
    'EJEMPLO DE LLAMADA:    fl_VerificaConformidad
    
    On Error GoTo EtiqError

    fl_VerificaConformidad = True ' asuminos que se cuenta con los datos mínimos para grabar

    
    ' verificamos si los montos son correstos
    rs_grdBeneficiario.MoveFirst
    While Not rs_grdBeneficiario.EOF
        If rs_grdBeneficiario!monto_us <= 0 Then
            MsgBox "El monto [" & rs_grdBeneficiario!monto_us & "] a liquidar correspondiente a [" & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & " " & rs_grdBeneficiario!Nombre & "] no es vàlido." & Chr(13) & "Corrija el error e intente registrar conformidad nuevamente.", vbInformation, "Aviso"
'            Call grdBeneficiario_RowColChange(0, 0)
            cmdBenMonto.SetFocus ' se posiciona en el boton de editar
            fl_VerificaConformidad = False
            Exit Function
        End If
        rs_grdBeneficiario.MoveNext
    Wend

    rs_grdBeneficiario.MoveFirst
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Function

EtiqError:
    fl_VerificaConformidad = False
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Function

Private Function fl_VerificaAprobar() As Boolean
    'TITULO:                Función fl_VerificaAprobar
    'PROPOSITO:             Verifica los datos para aprobar la liquidación
    'EJEMPLO DE LLAMADA:    fl_VerificaAprobar
    
    On Error GoTo EtiqError
    
    fl_VerificaAprobar = True ' asuminos que se cuenta con los datos mnimos para grabar

    ' verificamos registro de conformidad
    If Not (fl_VerificaConformidad) Then
        fl_VerificaAprobar = False
            Exit Function
    End If
    
    ' verificamos si tiene registro de conformidad
    rs_grdBeneficiario.MoveFirst
    While Not rs_grdBeneficiario.EOF
        If rs_grdBeneficiario!estado_conformidad <> "APR" Then
            MsgBox "La conformidad correspondiente a [" & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & " " & rs_grdBeneficiario!Nombre & "] no esta registrada." & Chr(13) & "Corrija el error e intente registrar aprobar nuevamente.", vbInformation, "Aviso"
'            Call grdBeneficiario_RowColChange(0, 0)
            cmdBenConf.SetFocus ' se posiciona en el boton
            fl_VerificaAprobar = False
            Exit Function
        End If
        rs_grdBeneficiario.MoveNext
    Wend
    rs_grdBeneficiario.MoveFirst ' para no probar algun error de posicion

    ' verificamos si la modalidad de liquidación es coherente
    SQLs = "SELECT ges_gestion, unidad_codigo, codigo_grupo, numero_pago, count(codigo_beneficiario) as numbenf from ro_pagos_cronograma_detalle "
    SQLs = SQLs & "WHERE ges_gestion = '" & rs_grdPrincipal!ges_gestion & "' "
    SQLs = SQLs & " AND unidad_codigo = '" & rs_grdPrincipal!unidad_codigo & "' "
    SQLs = SQLs & " AND codigo_grupo = " & rs_grdPrincipal!CODIGO_GRUPO
    SQLs = SQLs & " AND estado_conformidad = 'APR' "
    SQLs = SQLs & " AND estado_devengado = 'APR' "
    SQLs = SQLs & " GROUP BY ges_gestion, unidad_codigo, codigo_grupo, numero_pago "
    SQLs = SQLs & " HAVING Count(codigo_beneficiario) > 1"
        
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 And rs_grdPrincipal!modalidad_pago = "I" Then
        MsgBox "La modalidad de liquidación [Planilla individual] correspondiente al grupo: [" & rs_grdPrincipal!CODIGO_GRUPO & "] [" & rs_grdPrincipal!descripcion_grupo & "], unidad: [" & rs_grdPrincipal!unidad_codigo & "] no es válida." & Chr(13) & "Corrija el error e intente aprobar nuevamente.", vbInformation, "Aviso"
        BtnModificar.SetFocus ' se posiciona en el boton de editar grupo
        fl_VerificaAprobar = False
        Exit Function
    End If
    
    ' verificamos si existe pendientes ordenes de liquidacion sin procesar aprobar
    SQLs = "SELECT 'MinPago' = MIN(numero_pago) FROM ro_pagos_cronograma "
    SQLs = SQLs & "WHERE (estado_codigo ='REG' or estado_devengado ='REG') and ges_gestion = '" & rs_grdPrincipal!ges_gestion & "' "
    SQLs = SQLs & "AND unidad_codigo = '" & rs_grdPrincipal!unidad_codigo & "' "
    SQLs = SQLs & "AND codigo_grupo = " & rs_grdPrincipal!CODIGO_GRUPO
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        If rstTemp!MinPago < rs_grdLiquida!NUMERO_PAGO Then
            MsgBox "La Orden de Liquidación Nro:[" & rstTemp!MinPago & "] no fue procesada. Debe ser procesada antes de procesar una Liquidación posterior." & Chr(13) & "Corrija el error e intente procesar nuevamente nuevamente.", vbInformation, "Aviso"
            rs_grdLiquida.MoveFirst
            rs_grdLiquida.Find " numero_pago =" & rstTemp!MinPago
            Call grdLiquida_RowColChange(0, 0)
            grdLiquida.SetFocus ' se posiciona en el boton de conmformidad
            fl_VerificaAprobar = False
            Exit Function
        End If
      
    End If
    
    ' verifica si los beneficiarios cuentan con registro de emite o no factura
    rs_grdBeneficiario.MoveFirst
    While Not rs_grdBeneficiario.EOF
        If Len(Trim(rs_grdBeneficiario!emite_factura)) = 0 Then
            MsgBox "El beneficiario [" & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & " " & rs_grdBeneficiario!Nombre & "] NO cuanta con registro de CON/SIN RETENCION." & Chr(13) & "Corrija el error e intente aprobar nuevamente.", vbInformation, "Aviso"
'            Call grdBeneficiario_RowColChange(0, 0)
            cmdBenFact.SetFocus ' se posiciona en el boton
            fl_VerificaAprobar = False
            Exit Function
        End If
        rs_grdBeneficiario.MoveNext
    Wend

    rs_grdBeneficiario.MoveFirst
    
    ' verificamos si existe registro de contrato si es asi verfiica si lasd fechas de inicio y fin estan registrados
    rs_grdBeneficiario.MoveFirst
    While Not rs_grdBeneficiario.EOF
        SQLs = "SELECT ao_contrato_c.fechas_confirmado FROM ao_adjudica_c LEFT OUTER JOIN ao_contrato_c ON ao_adjudica_c.ges_gestion = ao_contrato_c.ges_gestion AND ao_adjudica_c.unidad_codigo = ao_contrato_c.unidad_codigo AND "
        SQLs = SQLs & "ao_adjudica_c.codigo_solicitud = ao_contrato_c.codigo_solicitud AND ao_adjudica_c.numero_consultoria = ao_contrato_c.numero_consultoria AND ao_adjudica_c.codigo_beneficiario = ao_contrato_c.codigo_beneficiario "
        SQLs = SQLs & "WHERE ao_adjudica_c.gp_ges_gestion = '" & LblGestion.Caption & "' AND ao_adjudica_c.gp_unidad_codigo = '" & lblCodUniSol.Caption & "' AND ao_adjudica_c.gp_codigo_grupo = " & Val(lblCodGrupo.Caption) & " AND ao_adjudica_c.codigo_beneficiario = '" & rs_grdBeneficiario!codigo_beneficiario & "'"
        Set rstTemp = New ADODB.Recordset
        rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
        If rstTemp!fechas_confirmado <> "APR" Then
            MsgBox "La fechas de inicio y fin de contrato de [" & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & " " & rs_grdBeneficiario!Nombre & "] no se encuentra confirmado." & Chr(13) & "Corrija el error e intente registrar procesar nuevamente.", vbInformation, "Aviso"
'            Call grdBeneficiario_RowColChange(0, 0)
            cmdDatosContrato.SetFocus ' se posiciona en el boton
            fl_VerificaAprobar = False
            Exit Function
        End If
        rs_grdBeneficiario.MoveNext
    Wend
    rs_grdBeneficiario.MoveFirst ' para no probar algun error de posicion
  
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Function

EtiqError:
    fl_VerificaAprobar = False
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Function

Private Function fl_VerificaDevengar() As Boolean
    'TITULO:                Función fl_VerificaDevengar
    'PROPOSITO:             Verifica los datos para devengar la liquidación
    'EJEMPLO DE LLAMADA:    fl_VerificaDevengar
    
    On Error GoTo EtiqError
    
    fl_VerificaDevengar = True ' asuminos que se cuenta con los datos mnimos para grabar

    ' verificamos registro de conformidad
    If Not (fl_VerificaConformidad) Then
        fl_VerificaDevengar = False
        Exit Function
    End If
    
    ' verificamos aprobación de liquidación
    If rs_grdLiquida!estado_codigo <> "APR" Then
        MsgBox "La liquidación Nro.: [" & rs_grdLiquida!NUMERO_PAGO & "] [" & rs_grdLiquida!Concepto & "] no se encuentra aprobado." & Chr(13) & "Corrija el error e intente devengar nuevamente.", vbInformation, "Aviso"
        cmdPagoAprob.SetFocus ' se posiciona en el boton
        fl_VerificaDevengar = False
        Exit Function
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Function

EtiqError:
    fl_VerificaDevengar = False
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Function

Private Sub pl_GeneraDevengado()
    'TITULO:                Procedimiento pl_GeneraDevengado
    'PROPOSITO:             Genera el devengado por beneficiario pasando por la tabla temporal
    'EJEMPLO DE LLAMADA:    call pl_GeneraDevengado
    
Dim Respuesta As String
Dim montocontrol As Double
Dim sesion$
Dim Error As Integer
Dim SeAsigno As Boolean

Dim rstOrden As New ADODB.Recordset
Dim RsTmp As New ADODB.Recordset
Dim rsc As New ADODB.Recordset


    On Error GoTo EtiqError ' activamos el manejador de errores
    Screen.MousePointer = vbHourglass
Error = 0

sesion = Left("S" & CStr(Rnd()), 10)

Set RsTmp = New ADODB.Recordset
RsTmp.Open "select * from ac_Ben_Devengado_TMP where sesion='" & sesion & "'", db, adOpenDynamic, adLockOptimistic

Do While RsTmp.RecordCount > 0
    sesion = Left("S" & CStr(Rnd()), 10)
    Set RsTmp = New ADODB.Recordset
    RsTmp.Open "select * from ac_Ben_Devengado_TMP where sesion='" & sesion & "'", db, adOpenDynamic, adLockOptimistic
Loop

'recorre todos los beneficiarios para generar su devengado en tabla temporal
rs_grdBeneficiario.MoveFirst
Do While Not rs_grdBeneficiario.EOF
    If rs_grdBeneficiario!estado_conformidad = "APR" Then
        'JQ QR
        'DE.dbo_apGeneralSearching "update ac_ben_comprdeven set monto_dolares_acum = 0 where gp_ges_gestion = '" & lblGestion.Caption & "' and gp_unidad_codigo ='" & lblCodUniSol.Caption & "' and gp_codigo_grupo = " & Val(lblCodGrupo.Caption) & " and codigo_beneficiario = '" & rs_grdBeneficiario!codigo_beneficiario & "'"
        
        SQLs = "select MIN(ordencomprobante) AS ordencomprobante From AC_BEN_COMPRDEVEN WHERE MONTO_BOLIVIANOS >0 AND gp_ges_gestion = '" & LblGestion.Caption & "' and gp_unidad_codigo ='" & lblCodUniSol.Caption & "' and gp_codigo_grupo = " & Val(lblCodGrupo.Caption) & " AND codigo_beneficiario = '" & rs_grdBeneficiario!codigo_beneficiario & "' and aprobotesoreria   = 'APR' and tipocomprobante = 'COM' order by ordencomprobante"
''        SQLs = "select distinct ordencomprobante From AC_BEN_COMPRDEVEN WHERE gp_ges_gestion = '" & lblGestion.Caption & "' and gp_unidad_codigo ='" & lblCodUniSol.Caption & "' and gp_codigo_grupo = " & Val(lblCodGrupo.Caption) & " AND codigo_beneficiario = '" & rs_grdBeneficiario!codigo_beneficiario & "' and aprobotesoreria   = 'APR' and tipocomprobante = 'COM' order by ordencomprobante"
        Set rstOrden = New ADODB.Recordset
        rstOrden.Open SQLs, db, adOpenStatic, adLockReadOnly
        
        If rstOrden.RecordCount = 0 Then
            Error = 1
        Else
             SeAsigno = False
             Do While Not rstOrden.EOF
                 If fl_ExisteSaldoParaDevengar(rs_grdBeneficiario!codigo_beneficiario, rstOrden!ordenComprobante) Then
                     If fl_ExisteEspacioEnComprometido(sesion, rstOrden!ordenComprobante) Then
                         'devenga lo que se pueda del comprometido --> TMP
                         ' comprueba si es solo porcentaje 100%
                         'JQ QR
                         'DE.dbo_ap_ComSolo100y258o222 rs_grdPrincipal!ges_gestion, rs_grdPrincipal!unidad_codigo, rs_grdPrincipal!codigo_grupo, rs_grdBeneficiario!codigo_beneficiario, rstOrden!ordenComprobante, Respuesta
                         If Respuesta = "APR" Then
                            'JQ QR
                            'DE.dbo_ap_GeneraDevEnTmp100 sesion, rs_grdPrincipal!ges_gestion, rs_grdBeneficiario!codigo_beneficiario, rstOrden!ordenComprobante, rs_grdPrincipal!unidad_codigo, rs_grdPrincipal!codigo_grupo, rs_grdLiquida!NUMERO_PAGO, rs_grdBeneficiario!emite_factura
                         Else
                            'JQ QR
                            'DE.dbo_ap_GeneraDevengadoEnTmp sesion, rs_grdPrincipal!ges_gestion, rs_grdBeneficiario!codigo_beneficiario, rstOrden!ordenComprobante, rs_grdPrincipal!unidad_codigo, rs_grdPrincipal!codigo_grupo, rs_grdLiquida!NUMERO_PAGO, rs_grdBeneficiario!emite_factura
                         End If
                         
                         SeAsigno = True
                     End If
                 End If
                 rstOrden.MoveNext
             Loop
             
             If SeAsigno = False Then
                 Error = 2
             End If
         End If
    
    End If
    'toma siguiente beneficiario
    rs_grdBeneficiario.MoveNext
    If Error > 0 Then Exit Do
Loop

If Error = 0 Then ' sin errores
    'este SP genera el devengado en base a la tabla ac_ben_devengado_tmp teniendo como agrupador la sesión
    'JQ QR
    'DE.dbo_ap_GeneraDevengado sesion, rs_grdPrincipal!ges_gestion, rs_grdPrincipal!unidad_codigo, rs_grdPrincipal!codigo_grupo, rs_grdLiquida!NUMERO_PAGO, GlUsuario

ElseIf Error = 1 Then
    MsgBox "Existe conformidad de parte de la Unidad, pero el Compromiso de Pago no está aprobado", vbCritical, "saf2002"
ElseIf Error = 2 Then
    MsgBox "No se genero Devengado por que existe error en los saldos del Compromiso de Pago, revise por favor", vbCritical, "SAF"
End If

'elimina registros de la sesion en la tabla temporal
'JQ QR
'DE.dbo_apGeneralSearching "delete ac_ben_devengado_tmp where sesion='" & sesion & "'"
    
    Screen.MousePointer = vbDefault
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Function fl_ExisteSaldoParaDevengar(codBen As String, ordenComprobante As Integer) As Boolean

    'verifica si existe saldo del comprometido para devengar
    fl_ExisteSaldoParaDevengar = False
    
    SQLs = "select saldo_US = SUM(monto_dolares), saldo_BS = SUM(monto_bolivianos) from ac_ben_comprdeven where codigo_beneficiario = '" & codBen & "' and tipocomprobante = 'COM' and aprobotesoreria='APR' AND GP_GES_GESTION='" & rs_grdPrincipal!ges_gestion & "' AND gp_unidad_codigo='" & rs_grdPrincipal!unidad_codigo & "' and  GP_CODIGO_GRUPO=" & rs_grdPrincipal!CODIGO_GRUPO
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    If Not rstTemp.EOF Then
        Do While Not rstTemp.EOF
            If rs_grdBeneficiario!tipo_moneda = "$US" And rstTemp!saldo_US - rs_grdBeneficiario!monto_us >= 0 Then
                fl_ExisteSaldoParaDevengar = True
            End If
            If rs_grdBeneficiario!tipo_moneda = "Bs" And rstTemp!saldo_BS - rs_grdBeneficiario!monto_bs >= 0 Then
                fl_ExisteSaldoParaDevengar = True
            End If
            rstTemp.MoveNext
        Loop
    End If
    rstTemp.Close
    
    
''''    'verifica si existe saldo del comprometido para devengar
''''    fl_ExisteSaldoParaDevengar = False
''''
''''''    SQLs = "select saldo = monto_dolares - monto_dolares_acum from ac_ben_comprdeven where codigo_beneficiario = '" & codBen & "' and ordencomprobante=" & ordenComprobante & " and tipocomprobante = 'COM' and aprobotesoreria='APR' AND GP_GES_GESTION='" & rs_grdPrincipal!ges_gestion & "' AND gp_unidad_codigo='" & rs_grdPrincipal!unidad_codigo & "' and  GP_CODIGO_GRUPO=" & rs_grdPrincipal!codigo_grupo
''''    SQLs = "select saldo_US = SUM(monto_dolares), saldo_BS = SUM(monto_bolivianos) from ac_ben_comprdeven where codigo_beneficiario = '" & codBen & "' and tipocomprobante = 'COM' and aprobotesoreria='APR' AND GP_GES_GESTION='" & rs_grdPrincipal!ges_gestion & "' AND gp_unidad_codigo='" & rs_grdPrincipal!unidad_codigo & "' and  GP_CODIGO_GRUPO=" & rs_grdPrincipal!codigo_grupo
''''    Set rstTemp = New ADODB.Recordset
''''    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
''''
''''    If Not rstTemp.EOF Then
''''        Do While Not rstTemp.EOF
''''            If rs_grdBeneficiario!tipo_moneda = "$US" And rstTemp!saldo_US - rs_grdBeneficiario!monto_US > 0 Then
''''            If rstTemp!saldo > 0 Then fl_ExisteSaldoParaDevengar = True
''''            rstTemp.MoveNext
''''        Loop
''''    End If
''''    rstTemp.Close
    
End Function

Function fl_ExisteEspacioEnComprometido(sesion As String, ordenComprobante As Integer) As Boolean
    'determina si existe espacio en el comprometido considerando los devengados benerados en la tabla temporal
    
    Dim rsc As New ADODB.Recordset
    Dim rsR As New ADODB.Recordset
    Dim rsT As New ADODB.Recordset
    Dim montoTMP As Double
    
    On Error GoTo EtiqError ' activamos el manejador de errores

    fl_ExisteEspacioEnComprometido = True
    
    rsR.Open "SELECT ges_gestion, org_codigo, codigo_pago " & _
             "From AC_BEN_COMPRDEVEN WHERE   codigo_beneficiario = '" & rs_grdBeneficiario!codigo_beneficiario & "' and " & _
                                            "aprobotesoreria   = 'APR'         and " & _
                                            "tipocomprobante   = 'COM'       and " & _
                                            "ordencomprobante  = " & ordenComprobante & " and " & _
                                            "gp_ges_Gestion    = '" & rs_grdPrincipal!ges_gestion & "' and " & _
                                            "gp_unidad_codigo  = '" & rs_grdPrincipal!unidad_codigo & "' and " & _
                                            "gp_codigo_grupo   = " & rs_grdPrincipal!CODIGO_GRUPO & " " & _
             "order by ges_gestion, org_codigo desc, codigo_pago", db, adOpenStatic, adLockReadOnly
    
    If rsR.RecordCount > 0 Then
        rsc.Open "SELECT monto_Dolares " & _
             "From pagos  WHERE     ges_Gestion = '" & rsR!ges_gestion & "' and " & _
                                     "org_codigo = '" & rsR!org_codigo & "' and " & _
                                     "codigo_pago = " & rsR!codigo_pago & " and " & _
                                     "tipo_formulario = 'COM' ", db, adOpenStatic, adLockReadOnly
        If rsc.EOF Then
            fl_ExisteEspacioEnComprometido = False
        Else
            rsT.Open "SELECT monto_dolares From ac_ben_devengado_TMP " & _
                     "WHERE sesion       = '" & sesion & "' and " & _
                           "Cges_gestion = '" & rsR!ges_gestion & "' and " & _
                           "Corg_codigo  = '" & rsR!org_codigo & "' and " & _
                           "Ccodigo_pago = " & rsR!codigo_pago, db, adOpenStatic, adLockReadOnly
            If rsT.RecordCount > 0 Then
                montoTMP = IIf(IsNull(rsT!monto_dolares), 0, rsT!monto_dolares)
            Else
                montoTMP = 0
            End If
            If rsc!monto_dolares - montoTMP > 0 Then
                fl_ExisteEspacioEnComprometido = True
            Else
                fl_ExisteEspacioEnComprometido = False
            End If
        End If
    Else
        fl_ExisteEspacioEnComprometido = False
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Function

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
    
End Function

Private Function fl_VerificaAnulaOrdLiq() As Boolean
    'TITULO:                Función fl_VerificaAnulaOrdLiq
    'PROPOSITO:             Verifica los datos para procesar la elimnacion
    'EJEMPLO DE LLAMADA:    fl_VerificaAnulaOrdLiq
    Dim rstTemp As ADODB.Recordset ' usado para la carga de los combos de base
    
    fl_VerificaAnulaOrdLiq = True
    
    ' verificamos si tiene algun devengado generado
    SQLs = "select * from ro_pagos_cronograma where ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and unidad_codigo='" & rs_grdPrincipal!unidad_codigo & "' and codigo_grupo=" & rs_grdPrincipal!CODIGO_GRUPO & " and numero_pago = " & rs_grdLiquida!NUMERO_PAGO & " and estado_devengado ='APR'"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    If rstTemp.RecordCount = 0 Then
        MsgBox "No tiene Orden de Liquidación generada para el Nro. de liquidación [" & rs_grdLiquida!NUMERO_PAGO & "].", vbInformation, "Aviso"
        grdLiquida.SetFocus
        fl_VerificaAnulaOrdLiq = False
        Exit Function
    End If
    
    ' verificamos si existe un numero de liquidación mayor con orden de liquidación
    SQLs = "select * from ro_pagos_cronograma where ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and unidad_codigo='" & rs_grdPrincipal!unidad_codigo & "' and codigo_grupo=" & rs_grdPrincipal!CODIGO_GRUPO & " and numero_pago > " & rs_grdLiquida!NUMERO_PAGO & " and estado_codigo in('REG','APR')"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    If rstTemp.RecordCount > 0 Then
        MsgBox "No puede anular la Orden de Liquidación Nro.[" & rs_grdLiquida!NUMERO_PAGO & "] por existir registro de liquidaciones posteriores.", vbInformation, "Aviso"
        grdLiquida.SetFocus
        fl_VerificaAnulaOrdLiq = False
        Exit Function
    End If
    
    ' verifica si tiene devengado de pago aprobados
    SQLs = "select * from ac_ben_comprDeven where gp_ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and gp_unidad_codigo ='" & rs_grdPrincipal!unidad_codigo & "' and gp_codigo_grupo =" & rs_grdPrincipal!CODIGO_GRUPO & " and gp_numero_pago = " & rs_grdLiquida!NUMERO_PAGO & " and tipoComprobante ='DEV' and aprobotesoreria='APR'"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        MsgBox "No puede anular la Orden de Liquidación correspondiente a [" & lblCodGrupo.Caption & "][" & lblDesGrupo.Caption & "] Nro. Liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "] de liquidación por tener devengado APROBADO." & Chr(13) & "Comuniquese con el administrador del sistema.", vbInformation, "Aviso"
        grdLiquida.SetFocus
        fl_VerificaAnulaOrdLiq = False
        Exit Function
    End If
    
    Set rstTemp = Nothing
    
End Function


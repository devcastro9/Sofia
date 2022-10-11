VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Ac_Puesto_Org 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clasificadores - RR.HH. - Puestos Funcionales"
   ClientHeight    =   6780
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11940
   Icon            =   "Ac_Puesto_Org.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Ac_Puesto_Org.frx":0A02
   ScaleHeight     =   6780
   ScaleWidth      =   11940
   Begin MSDataGridLib.DataGrid DtG_Auxiliar 
      Height          =   4065
      Left            =   15
      TabIndex        =   0
      Top             =   1875
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   7170
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   14737632
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "codigo_puesto"
         Caption         =   "Código"
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
         DataField       =   "denominacion_puesto"
         Caption         =   "Denominacion Puesto"
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
         DataField       =   "codigo_unidad"
         Caption         =   "Area"
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
         DataField       =   "unidad_org"
         Caption         =   "Unidad_Org"
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
         DataField       =   "fecha_creacion"
         Caption         =   "Fecha_Creacion"
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
         DataField       =   "vacante"
         Caption         =   "Vacante"
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
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2940.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   689.953
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraOpciones 
      BackColor       =   &H80000018&
      Height          =   1140
      Left            =   15
      TabIndex        =   14
      Top             =   700
      Width           =   5730
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   240
         Picture         =   "Ac_Puesto_Org.frx":22644
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Nuevo Registro"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdMod 
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "Ac_Puesto_Org.frx":29132
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "Anular"
         Height          =   720
         Left            =   1680
         Picture         =   "Ac_Puesto_Org.frx":299FC
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Anula Registro Activo"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3120
         Picture         =   "Ac_Puesto_Org.frx":2A6C6
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Busca un Registro"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdSal 
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   4560
         Picture         =   "Ac_Puesto_Org.frx":2B0C8
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Salir de Personas"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   3840
         Picture         =   "Ac_Puesto_Org.frx":2BACA
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Imprime Lista de Personas"
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton cmdAprueba 
         BackColor       =   &H0080C0FF&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2400
         Picture         =   "Ac_Puesto_Org.frx":2D24C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Aprueba Registro"
         Top             =   240
         Width           =   740
      End
   End
   Begin VB.Frame FraGrabarCancelar 
      BackColor       =   &H80000018&
      Height          =   1140
      Left            =   20
      TabIndex        =   22
      Top             =   700
      Width           =   5730
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   680
         Left            =   1560
         Picture         =   "Ac_Puesto_Org.frx":2D456
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   740
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reno&var"
         Height          =   540
         Left            =   2400
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   680
         Left            =   3000
         Picture         =   "Ac_Puesto_Org.frx":2D660
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   740
      End
   End
   Begin VB.Frame Fra_ABM 
      Height          =   5655
      Left            =   5760
      TabIndex        =   1
      Top             =   700
      Width           =   6135
      Begin VB.TextBox TxtEstado2 
         BackColor       =   &H00E0E0E0&
         DataField       =   "estado_registro"
         DataSource      =   "Ado_Auxiliar"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         TabIndex        =   38
         Text            =   "N"
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox TxtVacante2 
         Height          =   315
         ItemData        =   "Ac_Puesto_Org.frx":2D86A
         Left            =   3960
         List            =   "Ac_Puesto_Org.frx":2D874
         TabIndex        =   37
         Text            =   "SI"
         Top             =   480
         Width           =   660
      End
      Begin VB.TextBox TxtVacante 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "vacante"
         DataSource      =   "Ado_Auxiliar"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   36
         Text            =   "S"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox TxtSueldo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "idfuncionario"
         DataSource      =   "Ado_Auxiliar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   34
         Text            =   "0"
         Top             =   5040
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DtpFechaCrea 
         DataField       =   "fecha_creacion"
         DataSource      =   "Ado_Auxiliar"
         Height          =   285
         Left            =   2040
         TabIndex        =   26
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   84475905
         CurrentDate     =   40471
      End
      Begin VB.TextBox txtPuesto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "denominacion_puesto"
         DataSource      =   "Ado_Auxiliar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1120
         Width           =   5775
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "codigo_puesto"
         DataSource      =   "Ado_Auxiliar"
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
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtFuncion 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "funcion_general"
         DataSource      =   "Ado_Auxiliar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1800
         Width           =   5775
      End
      Begin VB.TextBox Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "estado_registro"
         DataSource      =   "Ado_Auxiliar"
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   5
         Text            =   "N"
         Top             =   480
         Width           =   615
      End
      Begin MSDataListLib.DataCombo Dtc_codigo 
         Bindings        =   "Ac_Puesto_Org.frx":2D880
         DataField       =   "codigo_unidad"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "codigo_unidad"
         BoundColumn     =   "codigo_unidad"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo Dtc_descrip 
         Bindings        =   "Ac_Puesto_Org.frx":2D898
         DataField       =   "codigo_unidad"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "Uni_descripcion_larga"
         BoundColumn     =   "codigo_unidad"
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
      Begin MSDataListLib.DataCombo DtcNivelDes 
         Bindings        =   "Ac_Puesto_Org.frx":2D8B0
         DataField       =   "nivel_puesto"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   4560
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "descripcion_nivel_puesto"
         BoundColumn     =   "nivel_puesto"
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
      Begin MSDataListLib.DataCombo DtcNivel 
         Bindings        =   "Ac_Puesto_Org.frx":2D8CF
         DataField       =   "nivel_puesto"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1440
         TabIndex        =   28
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "nivel_puesto"
         BoundColumn     =   "nivel_puesto"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcUniOrg 
         Bindings        =   "Ac_Puesto_Org.frx":2D8EE
         DataField       =   "unidad_ORG"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1920
         TabIndex        =   29
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "unidad_ORG"
         BoundColumn     =   "unidad_ORG"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcUniOrgDes 
         Bindings        =   "Ac_Puesto_Org.frx":2D909
         DataField       =   "unidad_ORG"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   3720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "denominacion_unidad"
         BoundColumn     =   "unidad_ORG"
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
      Begin MSDataListLib.DataCombo DtcNivelSB 
         Bindings        =   "Ac_Puesto_Org.frx":2D924
         DataField       =   "nivel_puesto"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   4560
         TabIndex        =   33
         Top             =   4200
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "sueldo_basico"
         BoundColumn     =   "nivel_puesto"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Vacante"
         Height          =   195
         Index           =   8
         Left            =   3900
         TabIndex        =   35
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Sueldo Basico:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   32
         Top             =   5080
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Unidad Organizacional"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   31
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Estado:"
         Height          =   255
         Index           =   5
         Left            =   5280
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Denominación del Puesto:"
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   11
         Top             =   900
         Width           =   2295
      End
      Begin VB.Label lblLabels 
         Caption         =   "Función Principal"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Caption         =   "Código Puesto:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Fecha de Creación:"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "AREA (Unidad Ejecutora):"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   1845
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nivel del Puesto:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   4320
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Ado_Auxiliar 
      Height          =   330
      Left            =   0
      Top             =   6000
      Width           =   5745
      _ExtentX        =   10134
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
      Caption         =   "Navegar"
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
   Begin MSAdodcLib.Adodc Ado_Clasificador 
      Height          =   330
      Left            =   0
      Top             =   6360
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
      Caption         =   "Ado_Clasificador"
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
   Begin MSAdodcLib.Adodc AdoUnidadOrg 
      Height          =   330
      Left            =   2160
      Top             =   6360
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
      Caption         =   "Ado_Clasificador"
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
   Begin MSAdodcLib.Adodc AdoUnidad 
      Height          =   330
      Left            =   4200
      Top             =   6360
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
      Caption         =   "AdoUnidad"
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
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESTRUCTURA DE PUESTOS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   450
      Left            =   7080
      TabIndex        =   40
      Top             =   120
      Width           =   4740
   End
End
Attribute VB_Name = "Ac_Puesto_Org"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_Clasificador As New ADODB.Recordset
Dim rs_Auxiliar As New ADODB.Recordset
Attribute rs_Auxiliar.VB_VarHelpID = -1

Dim var_cod As Integer
Dim VAR_VAL As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub cmdAprueba_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_Auxiliar!estado_codigo = "N" Then
      If sino = vbYes Then
         rs_Auxiliar!estado_codigo = "S"
         rs_Auxiliar!fecha_registro = Date
         rs_Auxiliar!usr_codigo = GlUsuario
         rs_Auxiliar.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_Auxiliar.CancelUpdate
        If mvBookMark > 0 Then
          rs_Auxiliar.Bookmark = mvBookMark
        Else
          rs_Auxiliar.MoveFirst
        End If
        mbDataChanged = False
        Fra_ABM.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        DtG_Auxiliar.Enabled = True
    End If
End Sub

Private Sub CmdDel_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_Auxiliar!estado_codigo = "S" Then
      If sino = vbYes Then
         rs_Auxiliar!estado_codigo = "L"
         rs_Auxiliar!fecha_registro = Date
         rs_Auxiliar!usr_codigo = GlUsuario
         rs_Auxiliar.UpdateBatch adAffectAll
      End If
   Else
      MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDesaprueba_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_Auxiliar!estado_codigo = "S" Then
      If sino = vbYes Then
         rs_Auxiliar!estado_codigo = "N"
         rs_Auxiliar!fecha_registro = Date
         rs_Auxiliar!usr_codigo = GlUsuario
         rs_Auxiliar.UpdateBatch adAffectAll
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub


Private Sub CmdGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If GlSW = "ADD" Then
      TxtCodigo.Text = rs_Auxiliar.RecordCount
      rs_Auxiliar!codigo_puesto = TxtCodigo.Text
      rs_Auxiliar!codigo_unidad = Dtc_codigo.Text
      rs_Auxiliar!ges_gestion = glGestion
    End If
      rs_Auxiliar!denominacion_puesto = txtPuesto.Text
      rs_Auxiliar!funcion_general = TxtFuncion.Text
      rs_Auxiliar!nivel_puesto = DtcNivel.Text
      If TxtVacante2.Text = "NO" Then
          rs_Auxiliar!vacante = "N"
      Else
          rs_Auxiliar!vacante = "S"
      End If
      rs_Auxiliar!fecha_creacion = DtpFechaCrea.Value
      rs_Auxiliar!unidad_ORG = DtcUniOrg.Text
      rs_Auxiliar!nivel_puesto = DtcNivel.Text
      rs_Auxiliar!idfuncionario = TxtSueldo.Text
      rs_Auxiliar!estado_codigo = "N"
      rs_Auxiliar!fecha_registro = Date
      rs_Auxiliar!usr_usuario = "ADMIN" 'GlUsuario
      rs_Auxiliar.Update    'Batch adAffectAll
      
      mbDataChanged = False
    
      Fra_ABM.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      DtG_Auxiliar.Enabled = True
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  If Dtc_codigo.Text = "" Then
    MsgBox "Debe registrar el AREA correspondiente ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If

End Sub

Private Sub CmdMod_Click()
  On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    DtG_Auxiliar.Enabled = False
    GlSW = "MOD"
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub CmdSal_Click()
'  If glPersNew = "O" Then
'    frmmo_pacientes.Dtc_ocupac = rs_Auxiliar!ocup_codigo
'    frmmo_pacientes.Dtc_OcupacDes = rs_Auxiliar!ocup_descripcion
'  End If
'  glPersNew = "N"
  Unload Me
End Sub

Private Sub Dtc_codigo_Click(Area As Integer)
    Dtc_descrip.BoundText = Dtc_codigo.BoundText
End Sub

Private Sub Dtc_descrip_Click(Area As Integer)
    Dtc_codigo.BoundText = Dtc_descrip.BoundText
End Sub

Private Sub DtcNivel_Click(Area As Integer)
    DtcNivelDes.BoundText = DtcNivel.BoundText
    DtcNivelSB.BoundText = DtcNivel.BoundText
End Sub

Private Sub DtcNivelDes_Click(Area As Integer)
    DtcNivel.BoundText = DtcNivelDes.BoundText
    DtcNivelSB.BoundText = DtcNivelDes.BoundText
End Sub

Private Sub DtcNivelDes_LostFocus()
    TxtSueldo.Text = DtcNivelSB.Text
End Sub

Private Sub DtcNivelSB_Click(Area As Integer)
    DtcNivel.BoundText = DtcNivelSB.BoundText
    DtcNivelDes.BoundText = DtcNivelSB.BoundText
End Sub

Private Sub DtcUniOrg_Click(Area As Integer)
    DtcUniOrgDes.BoundText = DtcUniOrg.BoundText
End Sub

Private Sub DtcUniOrgDes_Click(Area As Integer)
    DtcUniOrg.BoundText = DtcUniOrgDes.BoundText
End Sub

Private Sub Form_Load()

  Call abrirtabla
  
  Set rs_Clasificador = New ADODB.Recordset
  rs_Clasificador.Open "select * from rc_nivel_puesto  ", DB, adOpenKeyset, adLockOptimistic
  Set Ado_Clasificador.Recordset = rs_Clasificador.DataSource
  DtcNivelDes.BoundText = DtcNivel.BoundText
  
  Set rs_Unidad_Org = New ADODB.Recordset
  rs_Unidad_Org.Open "select * from rc_unidad_organizacional  ", DB, adOpenKeyset, adLockOptimistic
  Set AdoUnidadOrg.Recordset = rs_Unidad_Org.DataSource
  DtcUniOrgDes.BoundText = DtcUniOrg.BoundText
  
  Set rs_UNIDAD = New ADODB.Recordset
  rs_UNIDAD.Open "select * from fc_unidad_ejecutora  ", DB, adOpenKeyset, adLockOptimistic
  Set AdoUnidad.Recordset = rs_UNIDAD.DataSource
  Dtc_descrip.BoundText = Dtc_codigo.BoundText
  
'   Set rs_Unidad_Org = New ADODB.Recordset
'  rs_Unidad_Org.Open "select * from rc_unidad_organizacional  ", DB, adOpenKeyset, adLockOptimistic
'  Set AdoUnidadOrg.Recordset = rs_Unidad_Org.DataSource
  
'  rs_Auxiliar.AddNew
'  txtParam.Text = GlParametro
'  TxtForm.Text = GlForm
'  TxtCorrel.Text = GlCorrel

  mbDataChanged = False
  Fra_ABM.Enabled = False
  DtG_Auxiliar.Enabled = True
  GlSW = "NADA"
End Sub

Private Sub abrirtabla()
  Set rs_Auxiliar = New Recordset
  If rs_Auxiliar.State = 1 Then rs_Auxiliar.Close
  'queryinicial = "select * from rc_puesto_organizacional where param_codigo = '" & GlParametro & "' "
  queryinicial = "select * from rc_puesto_organizacional "
  rs_Auxiliar.Open queryinicial, DB, adOpenKeyset, adLockOptimistic
  rs_Auxiliar.Sort = "codigo_puesto"
  Set Ado_Auxiliar.Recordset = rs_Auxiliar.DataSource
  Set DtG_Auxiliar.DataSource = Ado_Auxiliar.Recordset
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
'    frmeo_Larvas_mosquitos.Fra_detalle.Visible = False
End Sub

Private Sub Ado_Auxiliar_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Muestra la posición de registro actual para este Recordset
  If rs_Auxiliar!vacante = "N" Then
          TxtVacante2.Text = "NO"
  Else
          TxtVacante2.Text = "SI"
  End If
  If rs_Auxiliar!estado_codigo = "" Or IsNull(rs_Auxiliar!estado_codigo) Then
          TxtEstado2.Text = "No aprobado"
          Txt_estado.Text = "N"
  End If
  If rs_Auxiliar!estado_codigo = "N" Then
          TxtEstado2.Text = "No aprobado"
  End If
  If rs_Auxiliar!estado_codigo = "S" Then
          TxtEstado2.Text = "Si Aprobado"
  End If
  If rs_Auxiliar!estado_codigo = "L" Then
          TxtEstado2.Text = "anuLado"
  End If
      Ado_Auxiliar.Caption = Ado_Auxiliar.Recordset.AbsolutePosition & " / " & Ado_Auxiliar.Recordset.RecordCount
End Sub

'Private Sub Ado_Auxiliar_WillChangeRecord(ByVal adReason As adodb.EventReasonEnum, ByVal cRecords As Long, adStatus As adodb.EventStatusEnum, ByVal pRecordset As adodb.Recordset)
'  'Aquí se coloca el código de validación
'  'Se llama a este evento cuando ocurre la siguiente acción
'  Dim bCancel As Boolean
'
'  Select Case adReason
'  Case adRsnAddNew
'  Case adRsnClose
'  Case adRsnDelete
'  Case adRsnFirstChange
'  Case adRsnMove
'  Case adRsnRequery
'  Case adRsnResynch
'  Case adRsnUndoAddNew
'  Case adRsnUndoDelete
'  Case adRsnUndoUpdate
'  Case adRsnUpdate
'  End Select
'
'  If bCancel Then adStatus = adStatusCancel
'End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
    'rs_Auxiliar.MoveLast
    rs_Auxiliar.AddNew
    'lblStatus.Caption = "Agregar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    DtG_Auxiliar.Enabled = False
    GlSW = "ADD"
    TxtVacante.Text = "SI"
    TxtEstado2.Text = "No aprobado"
'    rs_Auxiliar.AddNew
'    txtParam.Text = GlParametro
'    TxtForm.Text = "E-1" 'GlForm
'    TxtCorrel.Text = 1  'GlCorrel
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  rs_Auxiliar.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub


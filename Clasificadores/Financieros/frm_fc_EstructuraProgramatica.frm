VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_fc_EstructuraProgramatica 
   BackColor       =   &H00000000&
   Caption         =   "Clasificadores - Financieros - Proyectos"
   ClientHeight    =   5745
   ClientLeft      =   1290
   ClientTop       =   1890
   ClientWidth     =   10755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   10755
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "frm_fc_EstructuraProgramatica.frx":0000
      ScaleHeight     =   960
      ScaleWidth      =   12000
      TabIndex        =   30
      Top             =   120
      Width           =   12060
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "frm_fc_EstructuraProgramatica.frx":6C032
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "frm_fc_EstructuraProgramatica.frx":6C23C
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton CmdBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "frm_fc_EstructuraProgramatica.frx":6C446
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdIMPRIMIR 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4320
         Picture         =   "frm_fc_EstructuraProgramatica.frx":6C9FE
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   5160
         Picture         =   "frm_fc_EstructuraProgramatica.frx":6CFBB
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdBorrar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "frm_fc_EstructuraProgramatica.frx":6D1C5
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "frm_fc_EstructuraProgramatica.frx":6DE8F
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdAdicionar 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "frm_fc_EstructuraProgramatica.frx":6E46F
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3720
         MaskColor       =   &H00000000&
         Picture         =   "frm_fc_EstructuraProgramatica.frx":6EA93
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdaceptar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1680
         Picture         =   "frm_fc_EstructuraProgramatica.frx":6EC9D
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROYECTOS"
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
         Left            =   8235
         TabIndex        =   41
         Top             =   300
         Width           =   1965
      End
   End
   Begin VB.PictureBox fradatos 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4995
      Left            =   5640
      ScaleHeight     =   4935
      ScaleWidth      =   6495
      TabIndex        =   10
      Top             =   1200
      Width           =   6555
      Begin VB.ComboBox Combo2 
         DataField       =   "estado_codigo"
         DataSource      =   "Adoestructura"
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
         Height          =   315
         Left            =   4920
         TabIndex        =   29
         Text            =   "Combo2"
         Top             =   360
         Width           =   1110
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "pro_nivel"
         DataSource      =   "Adoestructura"
         Height          =   315
         Left            =   4560
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox Txtusuario 
         Height          =   285
         Left            =   4560
         TabIndex        =   25
         Text            =   "Text13"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Txthora 
         Height          =   285
         Left            =   3240
         TabIndex        =   24
         Text            =   "Text12"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Txtfecha 
         DataField       =   "fecha_registro"
         DataSource      =   "Adoestructura"
         Height          =   285
         Left            =   1800
         TabIndex        =   23
         Text            =   "Text11"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         DataField       =   "pro_codigo_sisin"
         DataSource      =   "Adoestructura"
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Text            =   "Text11"
         Top             =   3720
         Width           =   2220
      End
      Begin VB.TextBox Text5 
         DataField       =   "pro_sigla"
         DataSource      =   "Adoestructura"
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Text            =   "-"
         Top             =   2985
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         DataField       =   "pro_descripcion"
         DataSource      =   "Adoestructura"
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1680
         Width           =   5895
      End
      Begin VB.TextBox Text4 
         DataField       =   "pro_actividad"
         DataSource      =   "Adoestructura"
         Height          =   285
         Left            =   5280
         TabIndex        =   4
         Text            =   "Text5"
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text3 
         DataField       =   "pro_codigo"
         DataSource      =   "Adoestructura"
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Text            =   "Text4"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text2 
         DataField       =   "pro_proyecto"
         DataSource      =   "Adoestructura"
         Height          =   285
         Left            =   3960
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "pro_programa"
         DataSource      =   "Adoestructura"
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text10 
         DataField       =   "ges_gestion"
         DataSource      =   "Adoestructura"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Text            =   "2011"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Label3"
         Height          =   225
         Left            =   3360
         TabIndex        =   28
         Top             =   4200
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label7 
         Caption         =   "Label3"
         Height          =   225
         Left            =   240
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO SISIN:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   3765
         Width           =   1185
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA REGISTRO"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   4360
         Width           =   1500
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GESTION:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblLabels 
         Caption         =   "ACTIVIDAD"
         Height          =   255
         Index           =   4
         Left            =   5160
         TabIndex        =   19
         Top             =   615
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO REGISTRO:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   3120
         TabIndex        =   18
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "DENOMINACION"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   1395
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL DEL CODIGO:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   4560
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         Caption         =   "PROGRAMA"
         Height          =   255
         Index           =   8
         Left            =   2400
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   14
         Top             =   870
         Width           =   885
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SIGLA PROYECTO:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   13
         Top             =   3030
         Width           =   1440
      End
      Begin VB.Label lblLabels 
         Caption         =   "SUBPROGRAMA"
         Height          =   255
         Index           =   11
         Left            =   3600
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label lblLabels 
         Caption         =   "USUARIO"
         Height          =   255
         Index           =   12
         Left            =   4680
         TabIndex        =   11
         Top             =   4320
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   4995
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5400
      TabIndex        =   9
      Top             =   1200
      Width           =   5460
      Begin MSAdodcLib.Adodc Adoestructura 
         Height          =   330
         Left            =   0
         Top             =   4620
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         Caption         =   "Adodc1"
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
      Begin MSDataGridLib.DataGrid Grdlista 
         Bindings        =   "frm_fc_EstructuraProgramatica.frx":6EEA7
         Height          =   4575
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   8070
         _Version        =   393216
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "Pro_programa"
            Caption         =   "Pro_programa"
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
            DataField       =   "pro_proyecto"
            Caption         =   "Pro_proyecto"
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
            DataField       =   "pro_codigo"
            Caption         =   "Codigo"
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
            DataField       =   "Pro_actividad"
            Caption         =   "Pro_actividad"
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
            DataField       =   "pro_descripcion"
            Caption         =   "Nombre del Proyecto"
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
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
            DataField       =   "Pro_sigla"
            Caption         =   "Sigla"
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
            DataField       =   "pro_nivel"
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
         BeginProperty Column08 
            DataField       =   "pro_codigo_sisin"
            Caption         =   "Codigo SISIN"
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
            DataField       =   "Ges_gestion"
            Caption         =   "Gestion"
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
            DataField       =   "fecha_registro"
            Caption         =   "fecha_registro"
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
            DataField       =   "hora_registro"
            Caption         =   "hora_registro"
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
         BeginProperty Column12 
            DataField       =   "usr_codigo"
            Caption         =   "usr_usuario"
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
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3734.929
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_fc_EstructuraProgramatica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstESTRUCTURA As New ADODB.Recordset
Dim rsauxiliar As New ADODB.Recordset
Dim CAMPOS As ADODB.Field
'Dim ClBuscaGrid As CompBusquedas.ClBuscaEnGridExterno
 Dim sql_estructura As String
 Dim sw2 As String

Private Sub Adoestructura_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If pRecordset.EOF Or pRecordset.BOF Then
      'cmdEditar.Enabled = False
      'cmdBorrar.Enabled = False
      Text1.Text = Empty
      Text2.Text = Empty
      Text3.Text = Empty
      Text4.Text = Empty
      Text5.Text = Empty
      Text6.Text = Empty
      Text9.Text = Empty
      Text10.Text = Empty
    Exit Sub
   End If
    cmdEditar.Enabled = True
    'cmdBorrar.Enabled = True
   
   Select Case pRecordset.EditMode
      Case adEditInProgress
      Case adEditNone
'         Text1.Text = IIf(IsNull(pRecordset("pro_programa")), "", pRecordset("pro_programa"))
'         Text3.Text = IIf(IsNull(pRecordset("pro_codigo")), "", pRecordset("pro_codigo"))
'         Text2.Text = IIf(IsNull(pRecordset("pro_proyecto")), "", pRecordset("pro_proyecto"))
'         Text4.Text = IIf(IsNull(pRecordset("pro_actividad")), "", pRecordset("pro_actividad"))
'         Text5.Text = IIf(IsNull(pRecordset("pro_sigla")), "", pRecordset("pro_sigla"))
'         Text6.Text = IIf(IsNull(pRecordset("pro_descripcion")), "", pRecordset("pro_descripcion"))
'         Combo1.Text = IIf(IsNull(pRecordset("pro_nivel")), "", pRecordset("pro_nivel"))
'         Combo2.Text = IIf(IsNull(pRecordset("estado_codigo")), "", pRecordset("estado_codigo"))
'         Text9.Text = IIf(IsNull(pRecordset("pro_codigo_sisin")), "", pRecordset("pro_codigo_sisin"))
'         Text10.Text = IIf(IsNull(pRecordset("ges_gestion")), "", pRecordset("ges_gestion"))
'         Txtfecha.Text = IIf(IsNull(pRecordset("fecha_registro")), "", pRecordset("fecha_registro"))
'         Txthora.Text = IIf(IsNull(pRecordset("hora_registro")), "", pRecordset("hora_registro"))
'         Txtusuario.Text = IIf(IsNull(pRecordset("usr_codigo")), "", pRecordset("usr_codigo"))
              
      Case adEditDelete
      Case adEditAdd
   End Select
   Adoestructura.Caption = CStr(Adoestructura.Recordset.AbsolutePosition) & " de " & CStr(Adoestructura.Recordset.RecordCount)
End Sub
Private Sub CmdAceptar_Click()
Dim SQL_FOR As String
Dim sw As String
Dim rstestructuraaux As New ADODB.Recordset
On Error GoTo errorAceptar
 With Adoestructura
          If Text1 = "" Then
            MsgBox "INTRODUZCA DATOS"
            Text1.SetFocus
            Exit Sub
          End If
'                     If Text2 = "" Then
'                      MsgBox "INTRODUZCA DATOS"
'                      Text2.SetFocus
'                      Exit Sub
'                     End If
          If Text3 = "" Then
               MsgBox "INTRODUZCA DATOS"
                Text3.SetFocus
                Exit Sub
          End If
          If Text4 = "" Then
               MsgBox "INTRODUZCA DATOS"
                Text4.SetFocus
                Exit Sub
          End If
          If Text5 = "" Then
               MsgBox "INTRODUZCA DATOS"
                Text5.SetFocus
                Exit Sub
          End If
          If Text6 = "" Then
               MsgBox "INTRODUZCA DATOS"
                Text6.SetFocus
                Exit Sub
          End If
          If Text9 = "" Then
               MsgBox "INTRODUZCA DATOS"
                Text9.SetFocus
                Exit Sub
          End If
          If Text10 = "" Then
               MsgBox "INTRODUZCA DATOS"
                Text10.SetFocus
                Exit Sub
          End If
        If sw2 = "A" Then
            Set rstestructuraaux = New ADODB.Recordset
            SQL_FOR = "select * from Fc_estructura_programatica where pro_proyecto='" & Text3.Text & "' "
              'pro_programa = '" & Text1.Text & "' and pro_subprograma='" & Text2.Text & "' and and pro_actividad ='" & Text4.Text & "'
            rstestructuraaux.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic, adCmdText
            If rstestructuraaux.RecordCount > 0 Then
                'And Text1.Enabled And Text2.Enabled And Text3.Enabled And Text4.Enabled Then
                sw = True
                MsgBox " CODIGO DUPLICADO"
                Text1.SetFocus
                Exit Sub
            Else
                'DB.BeginTrans
                sw = False
                'If Text1.Enabled And Text2.Enabled And Text3.Enabled And Text4.Enabled Then
                Text3.Text = .Recordset.RecordCount
                .Recordset.AddNew
                .Recordset("pro_programa").Value = "26"  'Text1.Text
                .Recordset("pro_subprograma").Value = "00"  'Text2.Text
                .Recordset("pro_proyecto").Value = Val(Text3.Text) + 1
                .Recordset("pro_actividad").Value = "00"  'Text4.Text
                     
            End If
        End If
          .Recordset("pro_sigla").Value = Text5.Text
          .Recordset("pro_descripcion_larga").Value = Text6.Text
          .Recordset("pro_patron").Value = Trim(Combo1.Text)
          .Recordset("pro_activo").Value = Trim(Combo2.Text)
          .Recordset("codigo_sisin").Value = Text9.Text
          .Recordset("ges_gestion").Value = Text10.Text
          .Recordset("Usr_usuario").Value = frmLogin.txtUserName.Text
          .Recordset("fecha_registro").Value = Date
          .Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
          .Recordset.Update
          .Recordset.Requery
        'DB.CommitTrans
                                
End With
   
  Call Cmdadicionar_Click
   
  Call CmdCancelar_Click
   
Exit Sub

errorAceptar:
   
   Call pErrorRst(db.Errors)
  Adoestructura.Recordset.CancelUpdate
  db.RollbackTrans
End Sub

Private Sub Cmdadicionar_Click()
   Text1.Enabled = True
   Text2.Enabled = True
   Text3.Enabled = True
   Text4.Enabled = True
   Adoestructura.Enabled = False
   grdlista.Enabled = False
   fraDatos.Enabled = True
   'cmdBorrar.Visible = False
   CmdBuscar.Visible = False
   Cmdimprimir.Visible = False
   cmdSalir.Visible = False
   Cmdadicionar.Visible = False
   cmdEditar.Visible = False
   Cmdaceptar.Visible = True
   CmdCancelar.Visible = True
   Text1.Text = Empty
   Text2.Text = Empty
   Text3.Text = Empty
   Text4.Text = Empty
   Text5.Text = Empty
   Text6.Text = Empty
   Combo1.Text = Combo1.List(0)
   Combo2.Text = Combo2.List(0)
   Text9.Text = Empty
   Text10.Text = Empty
   Text10.SetFocus
   sw2 = "A"
End Sub

'Private Sub CmdBorrar_Click()
'
'   Dim Mensaje As String
' On Error GoTo errorDelete
'
'   Mensaje = "¿Borrar: " & _
'               Text1.Text & " " & _
'               Trim(Text6.Text) & "?"
'   If MsgBox(Mensaje, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar:") = vbYes Then
'      db.BeginTrans
'      Adoestructura.Recordset.Delete
'      db.CommitTrans
'   End If
'
'   Exit Sub
'
'errorDelete:
'
'   Dim e As ADODB.Error
'
'   For Each e In db.Errors
'      MsgBox "Error No. " & e.Number & " " & e.Description
'   Next
'
'   db.RollbackTrans
'
'End Sub


Private Sub CmdBuscar_Click()
''Busqueda.Visible = True
''fradatos.Enabled = True
' Set ClBuscaGrid = New CompBusquedas.ClBuscaEnGridExterno
'    Set ClBuscaGrid.Conexión = DB
'    ClBuscaGrid.EsTdbGrid = False
'    Set ClBuscaGrid.GridTrabajo = Grdlista
'    ClBuscaGrid.QueryUtilizado = sql_estructura
'    Set ClBuscaGrid.RecordsetTrabajo = Adoestructura.Recordset
'    'ClBuscaGrid.CamposVisibles = "11010011"
'    ClBuscaGrid.Ejecutar
End Sub

Private Sub CmdCancelar_Click()
  On Error Resume Next
   Text1.Enabled = True
   fraDatos.Enabled = False
   Adoestructura.Recordset.Requery
   grdlista.ReBind
  ' cmdBorrar.Visible = True
   CmdBuscar.Visible = True
   Cmdimprimir.Visible = True
   cmdSalir.Visible = True
   cmdEditar.Visible = True
   Cmdaceptar.Visible = False
    Cmdadicionar.Visible = True
   CmdCancelar.Visible = False
   Adoestructura.Enabled = True
   grdlista.Enabled = True
   Adoestructura.Recordset.Requery
   grdlista.ReBind
End Sub

Private Sub cmdEditar_Click()
   If Adoestructura.Recordset!estado_codigo = "REG" Then
     Adoestructura.Enabled = False
     grdlista.Enabled = False
     fraDatos.Enabled = True
     'cmdBorrar.Visible = False
     CmdBuscar.Visible = False
     Cmdimprimir.Visible = False
     cmdSalir.Visible = False
     Cmdadicionar.Visible = False
     cmdEditar.Visible = False
     Cmdaceptar.Visible = True
     CmdCancelar.Visible = True
     Text1.Enabled = False
     Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = True
     Text6.Enabled = True
     Text9.Enabled = True
     Text10.Enabled = True
    Text10.SetFocus
    sw2 = "M"
   Else
       MsgBox "No se puede modificar un registro Aprobado ...", , "Atencion"
   End If
End Sub


Private Sub Cmdimprimir_Click()
  Dim iResult As Integer
    'CrystalReport1.ReportFileName = App.Path & "\clasificadores\bancos\crybancos.rpt"
     CrystalReport1.WindowShowPrintSetupBtn = True
     CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.ReportFileName = "\SAF-2000\Clasificadores\Presupuesto\estructura programatica\cryprog.rpt"
  iResult = CrystalReport1.PrintReport
  If iResult <> 0 Then
      MsgBox CrystalReport1.LastErrorNumber & " : " & CrystalReport1.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
CrystalReport1.WindowState = crptMaximized
  
'REPPROG.Show

'   Set rptModalidadSeleccion.DataSource = rstestructura
'   rptModalidadSeleccion.Show vbModal
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub
Private Sub Form_Load()
  
     Label7.Caption = frmLogin.txtUserName.Text
'   Label9.Caption = Format(Time, "HH:mm:ss")
   Label11.Caption = Date
   fraDatos.Enabled = False
'   With fraEditar
'      .Visible = False
'      .Left = fraOpcion.Left
'      .Top = fraOpcion.Top
'   End With
   'cmdBorrar.enbled = False
   CmdBuscar.Visible = True
   Cmdimprimir.Visible = True
   cmdSalir.Visible = True
   Cmdadicionar.Visible = True
   cmdEditar.Visible = True
   Cmdaceptar.Visible = False
   CmdCancelar.Visible = False

   Set rstESTRUCTURA = New ADODB.Recordset
   sql_estructura = "select * from fc_estructura_programatica" ' order by pro_programa"
   rstESTRUCTURA.Open sql_estructura, db, adOpenKeyset, adLockOptimistic, adCmdText
   rstESTRUCTURA.Sort = "pro_codigo"
   Set Adoestructura.Recordset = rstESTRUCTURA
   
   Set ClBuscaGrid = Nothing
   sw2 = ""
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Resize()
   '  Centrear titulo
  ' With lblTitulo
   '   .Left = (fraTitulo.Width - .Width) \ 2
   'End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (rstESTRUCTURA.State = adStateClosed) Then rstESTRUCTURA.Close
   Set rstESTRUCTURA = Nothing
End Sub



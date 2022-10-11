VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frm_Puesto_Org 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ESTRUCTURA DE PUESTOS"
   ClientHeight    =   8085
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   10350
   Begin VB.Frame fraOpciones 
      BackColor       =   &H80000018&
      Height          =   5700
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   1050
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Adicionar"
         Height          =   720
         Left            =   120
         Picture         =   "Frm_Puesto_Org.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Nuevo Registro"
         Top             =   240
         Width           =   765
      End
      Begin VB.CommandButton CmdMod 
         Caption         =   "Modificar"
         Height          =   720
         Left            =   120
         Picture         =   "Frm_Puesto_Org.frx":6AEE
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   960
         Width           =   765
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "Anular"
         Enabled         =   0   'False
         Height          =   720
         Left            =   120
         Picture         =   "Frm_Puesto_Org.frx":73B8
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Anula Registro Activo"
         Top             =   1680
         Width           =   765
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   120
         Picture         =   "Frm_Puesto_Org.frx":8082
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Busca un Registro"
         Top             =   3120
         Width           =   765
      End
      Begin VB.CommandButton CmdSal 
         Caption         =   "Salir"
         Height          =   720
         Left            =   120
         Picture         =   "Frm_Puesto_Org.frx":894C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Salir de Personas"
         Top             =   4800
         Width           =   765
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   120
         Picture         =   "Frm_Puesto_Org.frx":8B56
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Imprime Lista de Personas"
         Top             =   3840
         Width           =   765
      End
      Begin VB.CommandButton cmdAprueba 
         BackColor       =   &H0080C0FF&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   120
         Picture         =   "Frm_Puesto_Org.frx":A2D8
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Aprueba Registro"
         Top             =   2400
         Width           =   770
      End
   End
   Begin VB.Frame FraGrabarCancelar 
      BackColor       =   &H80000018&
      Height          =   5700
      Left            =   0
      TabIndex        =   23
      Top             =   600
      Width           =   1050
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reno&var"
         Height          =   540
         Left            =   120
         TabIndex        =   26
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   120
         Picture         =   "Frm_Puesto_Org.frx":A4E2
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2520
         Width           =   765
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   720
         Left            =   120
         Picture         =   "Frm_Puesto_Org.frx":A6EC
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1320
         Width           =   780
      End
   End
   Begin VB.Frame Fra_ABM 
      Height          =   5535
      Left            =   4680
      TabIndex        =   2
      Top             =   600
      Width           =   5655
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "fecha_creacion"
         DataSource      =   "Ado_Auxiliar"
         Height          =   375
         Left            =   1680
         TabIndex        =   28
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51445761
         CurrentDate     =   40471
      End
      Begin VB.TextBox txtParam 
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
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1200
         Width           =   5415
      End
      Begin VB.TextBox TxtForm 
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
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtCorrel 
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
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2160
         Width           =   5415
      End
      Begin VB.TextBox Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "estado_registro"
         DataSource      =   "Ado_Auxiliar"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   6
         Text            =   "N"
         Top             =   360
         Width           =   615
      End
      Begin MSDataListLib.DataCombo Dtc_codigo 
         Bindings        =   "Frm_Puesto_Org.frx":AB2E
         DataField       =   "unidad_ORG"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1920
         TabIndex        =   13
         Top             =   4200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "unidad_ORG"
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
      Begin MSDataListLib.DataCombo Dtc_descrip 
         Bindings        =   "Frm_Puesto_Org.frx":AB49
         DataField       =   "unidad_ORG"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   4560
         Width           =   5295
         _ExtentX        =   9340
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
      Begin MSDataListLib.DataCombo DtcNivelDes 
         Bindings        =   "Frm_Puesto_Org.frx":AB64
         DataField       =   "nivel_puesto"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   1680
         TabIndex        =   29
         Top             =   3480
         Width           =   3855
         _ExtentX        =   6800
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
         Bindings        =   "Frm_Puesto_Org.frx":AB83
         DataField       =   "nivel_puesto"
         DataSource      =   "Ado_Auxiliar"
         Height          =   315
         Left            =   3240
         TabIndex        =   30
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "nivel_puesto"
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
      Begin VB.Label lblLabels 
         Caption         =   "Estado:"
         Height          =   255
         Index           =   5
         Left            =   4320
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Denominación del Puesto:"
         Height          =   255
         Index           =   20
         Left            =   180
         TabIndex        =   12
         Top             =   900
         Width           =   2295
      End
      Begin VB.Label lblLabels 
         Caption         =   "Función Principal"
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   11
         Top             =   1920
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
         Left            =   180
         TabIndex        =   10
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fecha de Creación:"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Unidad Organizacional"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nivel del Puesto:"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   3600
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid DtG_Auxiliar 
      Height          =   5160
      Left            =   1080
      TabIndex        =   0
      Top             =   615
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   9102
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   14737632
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
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
         DataField       =   "nivel_puesto"
         Caption         =   "Nivel"
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
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   599.811
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Ado_Auxiliar 
      Height          =   330
      Left            =   1080
      Top             =   5880
      Width           =   3585
      _ExtentX        =   6324
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
      BackColor       =   14737632
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
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESTRUCTURA DE PUESTOS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   405
      Left            =   5400
      TabIndex        =   1
      Top             =   0
      Width           =   4680
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   525
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10335
   End
   Begin VB.Image Image3 
      Height          =   1440
      Left            =   0
      Picture         =   "Frm_Puesto_Org.frx":ABA2
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "Frm_Puesto_Org"
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
   If rs_Auxiliar!estado_registro = "N" Then
      If sino = vbYes Then
         rs_Auxiliar!estado_registro = "S"
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
   If rs_Auxiliar!estado_registro = "S" Then
      If sino = vbYes Then
         rs_Auxiliar!estado_registro = "L"
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
   If rs_Auxiliar!estado_registro = "S" Then
      If sino = vbYes Then
         rs_Auxiliar!estado_registro = "N"
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
      rs_Auxiliar!codigo_puesto = TxtForm.Text
      rs_Auxiliar!unidad_org = Dtc_codigo.Text
      rs_Auxiliar!ges_gestion = "2010"
    End If
      rs_Auxiliar!denominacion_puesto = txtParam.Text
      rs_Auxiliar!funcion_general = txtCorrel.Text
      rs_Auxiliar!nivel_puesto = DtcNivel.Text
      rs_Auxiliar!vacante = "N"
      rs_Auxiliar!fecha_creacion = DTPicker1.Value
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
    MsgBox "Debe registrar la Actividad de la Gente ...", vbCritical + vbExclamation, "Validación de datos"
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
End Sub

Private Sub DtcNivelDes_Click(Area As Integer)
    DtcNivel.BoundText = DtcNivelDes.BoundText
End Sub

Private Sub Form_Load()

    Call abrirtabla
  
  Set rs_Clasificador = New ADODB.Recordset
  rs_Clasificador.Open "select * from rc_nivel_puesto  ", db, adOpenKeyset, adLockOptimistic
  Set Ado_Clasificador.Recordset = rs_Clasificador.DataSource
  Dtc_descrip.BoundText = Dtc_codigo.BoundText
  
  Set rs_Unidad_Org = New ADODB.Recordset
  rs_Unidad_Org.Open "select * from rc_unidad_organizacional  ", db, adOpenKeyset, adLockOptimistic
  Set AdoUnidadOrg.Recordset = rs_Unidad_Org.DataSource
  'Dtc_descrip.BoundText = Dtc_codigo.BoundText
  
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
  queryinicial = "select * from rc_puesto_organizacional  "
  rs_Auxiliar.Open queryinicial, db, adOpenKeyset, adLockOptimistic
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
'    rs_Auxiliar.AddNew
    txtParam.Text = GlParametro
    TxtForm.Text = "E-1" 'GlForm
    txtCorrel.Text = 1  'GlCorrel
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


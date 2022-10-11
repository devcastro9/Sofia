VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRc_Calendario2 
   Caption         =   "Clasificadores - Administrativos - Horarios Laborales"
   ClientHeight    =   8685
   ClientLeft      =   1065
   ClientTop       =   2415
   ClientWidth     =   16200
   Icon            =   "FrmRc_Calendario2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   16200
   WindowState     =   2  'Maximized
   Begin VB.Frame FraEmpresa 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   4260
      Left            =   5880
      TabIndex        =   10
      Top             =   620
      Width           =   10335
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   240
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   3120
         Width           =   3975
      End
      Begin VB.ComboBox Txt01 
         DataField       =   "hora_extra"
         DataSource      =   "adoLista"
         Height          =   315
         ItemData        =   "FrmRc_Calendario2.frx":31F48
         Left            =   240
         List            =   "FrmRc_Calendario2.frx":31F52
         TabIndex        =   13
         Text            =   "NO"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox txt02 
         DataField       =   "minutos_tolerancia"
         DataSource      =   "adoLista"
         Height          =   315
         ItemData        =   "FrmRc_Calendario2.frx":31F5E
         Left            =   240
         List            =   "FrmRc_Calendario2.frx":31F83
         TabIndex        =   12
         Text            =   "10"
         Top             =   2280
         Width           =   735
      End
      Begin VB.ComboBox TxtGestion 
         DataField       =   "ges_gestion"
         DataSource      =   "adoLista"
         Height          =   315
         ItemData        =   "FrmRc_Calendario2.frx":31FB3
         Left            =   240
         List            =   "FrmRc_Calendario2.frx":31FC3
         TabIndex        =   11
         Text            =   "2011"
         Top             =   600
         Width           =   900
      End
      Begin MSComCtl2.DTPicker DTPFec_Inicio 
         DataField       =   "fecha_desde"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   2520
         TabIndex        =   14
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   102236161
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Día:                                          Laboral/Feriado"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Width           =   3330
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Día:                                          Laboral/Feriado"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   2040
         Width           =   3330
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Gestion:                                      Fecha:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   45
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2790
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Día:                                          Laboral/Feriado"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   40
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   3330
      End
   End
   Begin VB.PictureBox picButtons 
      BackColor       =   &H00C0FFC0&
      Height          =   660
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   480
         Left            =   9120
         Picture         =   "FrmRc_Calendario2.frx":31FDF
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir de Personas"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton CmdIMPRIMIR 
         Caption         =   "Imprimir"
         Height          =   480
         Left            =   5040
         Picture         =   "FrmRc_Calendario2.frx":329E1
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd_busqueda 
         Caption         =   "&Buscar"
         Height          =   480
         Left            =   4080
         Picture         =   "FrmRc_Calendario2.frx":333E3
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdAdicionar 
         Caption         =   "Nuevo"
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
         Left            =   240
         Picture         =   "FrmRc_Calendario2.frx":33DE5
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Nuevo Registro"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Modif."
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
         Left            =   1200
         Picture         =   "FrmRc_Calendario2.frx":3436F
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Aprobar"
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
         Left            =   3120
         Picture         =   "FrmRc_Calendario2.frx":348F9
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aprueba Registro"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "AnuLar"
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
         Left            =   2160
         Picture         =   "FrmRc_Calendario2.frx":34E83
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Anula Registro Activo"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdaceptar 
         Caption         =   "Grabar"
         Height          =   480
         Left            =   6000
         Picture         =   "FrmRc_Calendario2.frx":35885
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   480
         Left            =   6960
         Picture         =   "FrmRc_Calendario2.frx":35E0F
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   855
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   7920
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
   End
   Begin MSAdodcLib.Adodc adoLista 
      Height          =   330
      Left            =   0
      Top             =   5040
      Width           =   10335
      _ExtentX        =   18230
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
      BackColor       =   12648384
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
   Begin MSAdodcLib.Adodc AdoTurno 
      Height          =   375
      Left            =   120
      Top             =   6120
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   661
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
      Caption         =   "AdoTurno"
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
   Begin MSDataGridLib.DataGrid grdlista 
      Bindings        =   "FrmRc_Calendario2.frx":36399
      Height          =   4035
      Left            =   0
      TabIndex        =   17
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7117
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648384
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
      ColumnCount     =   8
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "fecha"
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
      BeginProperty Column02 
         DataField       =   "dia"
         Caption         =   "Día"
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
         DataField       =   "tipo"
         Caption         =   "Lab/Fer"
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
         DataField       =   "descripcion"
         Caption         =   "Descripción"
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
      BeginProperty Column06 
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
      BeginProperty Column07 
         DataField       =   "usr_usuario"
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
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2475.213
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmRc_Calendario2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstorg As New ADODB.Recordset
Dim rs_turnos As New ADODB.Recordset
Dim rs_PARAMETRO As New ADODB.Recordset

Dim CAMPOS As ADODB.Field
'Dim ClBuscaGrid As CompBusquedas.ClBuscaEnGridExterno
Dim sql_financiador As String
Dim SW2 As String

Private Sub Adolista_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'   If pRecordset.EOF Or pRecordset.BOF Then
'      cmdEditar.Enabled = False
'      cmdBorrar.Enabled = False
'      Text1.Text = Empty
'      Text2.Text = Empty
'      Text3.Text = Empty
'      Text4.Text = Empty
'      dtcfu.Text = ""
'      dtcfue.Text = ""
'      Exit Sub
'   End If
   
'   cmdEditar.Enabled = True
'   cmdBorrar.Enabled = True
'
'   Select Case pRecordset.EditMode
'      Case adEditInProgress
'      Case adEditNone
'         Txt01.Text = IIf(IsNull(pRecordset("Dia_control")), "", pRecordset("Dia_control"))
'         TxtGestion.Text = IIf(IsNull(pRecordset("ges_gestion")), "", pRecordset("ges_gestion"))
'         txt02.Text = IIf(IsNull(pRecordset("minutos_tolerancia")), "", pRecordset("minutos_tolerancia"))
'         txt03.Value = IIf(IsNull(pRecordset("primer_ingreso")), "", pRecordset("primer_ingreso"))
'         txt04.Value = IIf(IsNull(pRecordset("primera_salida")), "", pRecordset("primera_salida"))
'         txt05.Value = IIf(IsNull(pRecordset("tope_primer_ingreso")), "", pRecordset("tope_primer_ingreso"))
'         Txt06.Text = IIf(IsNull(pRecordset("Hora_extra_AM")), "", pRecordset("Hora_extra_AM"))
'         txt07.Value = IIf(IsNull(pRecordset("segundo_ingreso")), "", pRecordset("segundo_ingreso"))
'         txt08.Value = IIf(IsNull(pRecordset("segunda_salida")), "", pRecordset("segunda_salida"))
'         txt09.Value = IIf(IsNull(pRecordset("tope_segundo_ingreso")), "", pRecordset("tope_segundo_ingreso"))
'         txt10.Text = IIf(IsNull(pRecordset("Hora_extra_PM")), "", pRecordset("Hora_extra_PM"))
'         DTPFec_Inicio.Value = IIf(IsNull(pRecordset("fecha_registro")), "", pRecordset("fecha_registro"))
'         TxtHora.Text = IIf(IsNull(pRecordset("hora_registro")), "", pRecordset("hora_registro"))
''         TxtUsuario.Text = IIf(IsNull(pRecordset("usr_usuario")), "", pRecordset("usr_usuario"))
'      Case adEditDelete
'      Case adEditAdd
'   End Select
   adoLista.Caption = CStr(adoLista.Recordset.AbsolutePosition) & " de " & CStr(adoLista.Recordset.RecordCount)
End Sub
   
Private Sub cmdAceptar_Click()
On Error GoTo errorAceptar
Dim sw As Boolean
Dim SQL_FOR As String
Dim RSTORAUX As New ADODB.Recordset

   With adoLista
        If TxtGestion = "" Then
              MsgBox "INTRODUZCA DATOS"
              TxtGestion.SetFocus
              Exit Sub
        End If
        If txt01 = "" Then
              MsgBox "INTRODUZCA DATOS"
              txt01.SetFocus
              Exit Sub
        End If
        If Txt02 = "" Then
              MsgBox "INTRODUZCA DATOS"
              Txt02.SetFocus
              Exit Sub
        End If
        If txt03 = "" Then
              MsgBox "INTRODUZCA DATOS"
              txt03.SetFocus
              Exit Sub
        End If
        If txt03 > txt04 Then
              MsgBox "La Hora de INGRESO, NO puede ser mayor a la fecha de SALIDA.."
              txt03.SetFocus
              Exit Sub
        End If
        If DTPFec_Inicio > DTPFec_Fin Then
              MsgBox "La fecha inicial de Control, NO puede ser mayor a la fecha Final de Control.."
              DTPFec_Inicio.SetFocus
              Exit Sub
        End If
        If (txt05 < txt03) Or (txt05 > txt04) Then
              MsgBox "La fecha tope debe ser mayor a la de INGRESO y menor a la de SALIDA.."
              txt05.SetFocus
              Exit Sub
        End If
'        If (txt09 < txt07) Or (txt09 > txt08) Then
'              MsgBox "La fecha tope debe ser mayor a la de INGRESO y menor a la de SALIDA.."
'              txt09.SetFocus
'              Exit Sub
'        End If
'
'    Set RSTORAUX = New ADODB.Recordset
'    SQL_FOR = "select * from Fc_ORGANISMO_FINANCIAMIENTO where ORG_CODIGO = '" & Text1.Text & "'"
'    RSTORAUX.Open SQL_FOR, DB, adOpenKeyset, adLockOptimistic, adCmdText
'    If RSTORAUX.RecordCount > 0 And Text1.Enabled Then
'      sw = True
'      MsgBox " CODIGO DUPLICADO"
'      Text1.SetFocus
'      Exit Sub
'    End If
    '
    'DB.BeginTrans
    sw = False
    If SW2 = "ADD" Then
'        .Recordset.AddNew
        .Recordset("ges_gestion").Value = TxtGestion.Text
        .Recordset("turno") = Dtc_Par.Text
    End If
    .Recordset("descripcion").Value = Dtc_ParDes.Text
    .Recordset("fecha_desde").Value = DTPFec_Inicio.Value
    .Recordset("fecha_hasta").Value = DTPFec_Fin.Value
    .Recordset("hora_extra").Value = (txt01.Text)
    .Recordset("hora_ingreso").Value = Format(txt03.Value, "HH:mm:ss")
    .Recordset("hora_salida").Value = Format(txt04.Value, "HH:mm:ss")
    .Recordset("tope_hora_ingreso").Value = Format(txt05.Value, "HH:mm:ss")
    .Recordset("minutos_tolerancia").Value = Txt02.Text
    .Recordset("usr_usuario").Value = GlUsuario
    .Recordset("fecha_registro").Value = Date   'Format(Date, "dd/mm/aaaa")
    .Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
    .Recordset.Update
    '.Recordset.Requery
    'DB.CommitTrans
            
   End With
   
   SW2 = "XX"
   Cmdadicionar.Visible = True
   Cmdeditar.Visible = True
   Cmdborrar.Visible = True
   cmdRefresh.Visible = True
   Cmd_busqueda.Visible = True
   'CmdIMPRIMIR.Visible = True
   Cmdsalir.Visible = True

   Cmdaceptar.Visible = False
   CmdCancelar.Visible = False
   adoLista.Enabled = True
   Grdlista.Enabled = True
   Call abrir_tabla
   TxtGestion.Enabled = True
   Exit Sub

errorAceptar:
   
   Call pErrorRst(DB.Errors)
   
   adoLista.Recordset.CancelUpdate
   
   'DB.RollbackTrans
End Sub
 Private Sub Cmdadicionar_Click()
   SW2 = "ADD"
   Cmdadicionar.Visible = False
   Cmdeditar.Visible = False
   Cmdborrar.Visible = False
   cmdRefresh.Visible = False
   Cmd_busqueda.Visible = False
   'CmdIMPRIMIR.Visible = False
   Cmdsalir.Visible = False

   Cmdaceptar.Visible = True
   CmdCancelar.Visible = True
   adoLista.Recordset.AddNew
   adoLista.Enabled = False
   Grdlista.Enabled = False
   'TxtGestion.Text = Empty
   'Txt01.Text = Empty
'   Txt02.Text = Empty
'   Txt06.Text = Empty
'   txt10.Text = Empty
End Sub

'Private Sub Cmdborrar_Click()
'   Dim Mensaje As String
'
'On Error GoTo errorDelete
'
'   Mensaje = "¿Borrar: " & _
'               Text1.Text & " " & _
'               Trim(Text3.Text) & "?"
'   If MsgBox(Mensaje, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar:") = vbYes Then
'      db.BeginTrans
'      adoLista.Recordset.Delete
'      db.CommitTrans
'   End If
'
'   Exit Sub
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

Private Sub Cmd_Busqueda_Click()
''BUSQUEDA.Visible = True
''fradatos.Enabled = True
' Set ClBuscaGrid = New CompBusquedas.ClBuscaEnGridExterno
'    Set ClBuscaGrid.Conexión = DB
'    ClBuscaGrid.EsTdbGrid = False
'    Set ClBuscaGrid.GridTrabajo = grdlista
'    ClBuscaGrid.QueryUtilizado = sql_financiador
'    Set ClBuscaGrid.RecordsetTrabajo = adoLista.Recordset
'    'ClBuscaGrid.CamposVisibles = "11010011"
'    ClBuscaGrid.Ejecutar

End Sub

Private Sub cmdCancelar_Click()
  On Error Resume Next
   SW2 = "XX"
   Cmdadicionar.Visible = True
   Cmdeditar.Visible = True
   Cmdborrar.Visible = True
   cmdRefresh.Visible = True
   Cmd_busqueda.Visible = True
   'CmdIMPRIMIR.Visible = True
   Cmdsalir.Visible = True

   Cmdaceptar.Visible = False
   CmdCancelar.Visible = False
   adoLista.Enabled = True
   Grdlista.Enabled = True
   Call abrir_tabla
   TxtGestion.Enabled = True
End Sub

Private Sub cmdEditar_Click()
   SW2 = "MOD"
   Cmdadicionar.Visible = False
   Cmdeditar.Visible = False
   Cmdborrar.Visible = False
   cmdRefresh.Visible = False
   Cmd_busqueda.Visible = False
   'CmdIMPRIMIR.Visible = False
   Cmdsalir.Visible = False

   Cmdaceptar.Visible = True
   CmdCancelar.Visible = True
   adoLista.Enabled = False
   Grdlista.Enabled = False

   TxtGestion.Enabled = False
'   txt01.Enabled = False
   txt01.SetFocus
End Sub

Private Sub CmdImprimir_Click()
  Dim IResult As Integer
    'CrystalReport1.ReportFileName = App.Path & "\clasificadores\bancos\crybancos.rpt"
     CrystalReport1.WindowShowPrintSetupBtn = True
     CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.ReportFileName = "\SAF-2000\Clasificadores\presupuesto\organismo financiador\cryorgfin.rpt"
  IResult = CrystalReport1.PrintReport
  If IResult <> 0 Then
      MsgBox CrystalReport1.LastErrorNumber & " : " & CrystalReport1.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If

CrystalReport1.WindowState = crptMaximized

'REPORGFIN.Show

'   rptModalidadSeleccion.Show vbModal
End Sub

Private Sub cmdRefresh_Click()
'  If adoLista.Recordset!turno <> "AM" And adoLista.Recordset!turno <> "PM" Then
'    If adoLista.Recordset!turno = "AM" Then
'        GlHora1 = adoLista.Recordset!tope_hora_ingreso
'    End If
'    If adoLista.Recordset!turno = "PM" Then
'        GlHora2 = adoLista.Recordset!tope_hora_ingreso
'    End If
'    Set rs_PARAMETRO = New ADODB.Recordset
'    rs_PARAMETRO.Open "select * from gc_parametros_sistema where estado_registro = 'S' ", cnn, adOpenDynamic, adLockReadOnly
'    If rs_PARAMETRO.RecordCount > 0 Then
'        'rs_PARAMETRO.MoveFirst
'        rs_PARAMETRO!Hora_Ingreso1 = GlHora1
'        rs_PARAMETRO!Hora_Ingreso2 = GlHora2
'    End If
'  End If
'  adoLista.Recordset!estado_registro = "SI"

'Dim dateValue As Date = #6/11/2008#
'Console.WriteLine(dateValue.ToString("dddd", New CultureInfo("es-ES"))     ' Displays miércoles.

'Weekday(date, [firstdayofweek])
'•vbUseSystemDayOfWeek = 0 (el del sistema)
'•vbSunday = 1
'•vbMonday = 2
'•vbTuesday = 3
'•vbWednesday = 4
'•vbThursday = 5
'•vbFriday = 6
'•vbSaturday = 7

Dim dia As String
Dim fechita As Date
'dia = WeekdayName(Weekday(Date))
'MsgBox dia

Dim rsCalendar As New ADODB.Recordset
   Set rsCalendar = New ADODB.Recordset
   rsCalendar.Open "select * from gc_calendario ", DB, adOpenKeyset, adLockOptimistic, adCmdText
'   Set AdoPais.Recordset = rsCalendar
Dim i As Integer
  i = 1
  fechita = CDate("01/01/2011")
  While fechita >= "01/01/2011" And fechita <= "31/12/2011"
     rsCalendar.AddNew
     dia = WeekdayName(Weekday(fechita))
     rsCalendar!fecha = fechita 'Format(fechita, "dd/mm/aaaa")
     rsCalendar!ges_gestion = Year(fechita)
     If dia = "sábado" Or dia = "domingo" Then
        rsCalendar!tipo = "F"
     Else
        rsCalendar!tipo = "H"
     End If
     fechita = fechita + 1
  Wend


End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

'Private Sub Dtc_Par_Click(Area As Integer)
''    Dtc_ParDes.BoundText = Dtc_Par.BoundText
'End Sub
'
'Private Sub Dtc_ParDes_Click(Area As Integer)
''    Dtc_Par.BoundText = Dtc_ParDes.BoundText
'End Sub

Private Sub Form_Load()
   'Dim sql_fuente As String
   SW2 = "XX"
   Cmdadicionar.Visible = True
   Cmdeditar.Visible = True
   Cmdborrar.Visible = True
   cmdRefresh.Visible = True
   Cmd_busqueda.Visible = True
   'CmdIMPRIMIR.Visible = True
   Cmdsalir.Visible = True

   Cmdaceptar.Visible = False
   CmdCancelar.Visible = False
   adoLista.Enabled = True
   Grdlista.Enabled = True
   
   Call abrir_tabla
   
   Set rs_turnos = New ADODB.Recordset
   'sql_fuente = "select * from rc_turnos" ' order by fte_codigo"
   rs_turnos.Open "select * from rc_turnos", DB, adOpenKeyset, adLockOptimistic, adCmdText
   rs_turnos.Sort = "correl"
'  ' MsgBox rstfue.RecordCount
   Set AdoTurno.Recordset = rs_turnos
'
   
'   Set rstorg = New ADODB.Recordset
'   sql_financiador = "select * from rc_horarios" 'order by org_codigo"
'   rstorg.Open sql_financiador, DB, adOpenKeyset, adLockOptimistic, adCmdText
''   rstorg.Sort = "Dia_control"
'   Set adoLista.Recordset = rstorg
'   'Set ClBuscaGrid = Nothing
  
End Sub

Private Sub abrir_tabla()
    Set rstorg = New ADODB.Recordset
    If rstorg.State = 1 Then rstorg.Close
    sql_financiador = "select * from rc_horarios" 'order by org_codigo"
    rstorg.Open sql_financiador, DB, adOpenKeyset, adLockOptimistic, adCmdText
    rstorg.Sort = "correl"
    Set adoLista.Recordset = rstorg
    Set Grdlista.DataSource = adoLista.Recordset
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (rstorg.State = adStateClosed) Then rstorg.Close
   'Set rstorg = Nothing

End Sub

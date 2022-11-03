VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmOrigenDestino 
   BackColor       =   &H8000000C&
   Caption         =   "TRASPASOS PRESUPUESTARIOS . . ."
   ClientHeight    =   5640
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   ControlBox      =   0   'False
   ForeColor       =   &H80000004&
   Icon            =   "FrmOrigenDestino.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdBuscaDes 
      Caption         =   "&Busca Destino"
      DownPicture     =   "FrmOrigenDestino.frx":0442
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Picture         =   "FrmOrigenDestino.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   900
   End
   Begin VB.CommandButton CmdBuscaOri 
      Caption         =   "&Busca Origen"
      DownPicture     =   "FrmOrigenDestino.frx":0CC6
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Picture         =   "FrmOrigenDestino.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   900
   End
   Begin MSDataGridLib.DataGrid DtgOrigenF 
      Bindings        =   "FrmOrigenDestino.frx":154A
      Height          =   2175
      Left            =   885
      TabIndex        =   0
      Top             =   120
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648447
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   19
      RowDividerStyle =   3
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
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Elija el Registro ""ORIGEN""  .  .  . "
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "fte_codigo"
         Caption         =   "Fte"
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
         DataField       =   "org_codigo"
         Caption         =   "Org"
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
         DataField       =   "pro_programa"
         Caption         =   "Pro"
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
         DataField       =   "pro_proyecto"
         Caption         =   "Pry"
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
         DataField       =   "pro_actividad"
         Caption         =   "Act"
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
         DataField       =   "par_codigo"
         Caption         =   "Partida"
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
         DataField       =   "fgs_formulado"
         Caption         =   "Formulado Bs."
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
         DataField       =   "fgs_adiciones"
         Caption         =   "Add/Red.Bs."
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
         DataField       =   "fgs_modificaciones"
         Caption         =   "Traspasos Bs."
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
         DataField       =   "fgs_vigente"
         Caption         =   "Vigente Bs."
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
         DataField       =   "par_descripcion_larga"
         Caption         =   "      Descripción"
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
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
         EndProperty
         BeginProperty Column10 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DtgDestinoF 
      Bindings        =   "FrmOrigenDestino.frx":1567
      Height          =   2175
      Left            =   885
      TabIndex        =   1
      Top             =   3015
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648447
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   19
      RowDividerStyle =   3
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
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Elija el Registro ""DESTINO"" . . ."
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "fte_codigo"
         Caption         =   "Fte"
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
         DataField       =   "org_codigo"
         Caption         =   "Org"
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
         DataField       =   "pro_programa"
         Caption         =   "Pro"
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
         DataField       =   "pro_proyecto"
         Caption         =   "Pry"
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
         DataField       =   "pro_actividad"
         Caption         =   "Act"
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
         DataField       =   "par_codigo"
         Caption         =   "Partida"
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
         DataField       =   "fgs_formulado"
         Caption         =   "Formulado Bs."
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
         DataField       =   "fgs_adiciones"
         Caption         =   "Add/Red.Bs."
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
         DataField       =   "fgs_modificaciones"
         Caption         =   "Traspasos Bs."
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
         DataField       =   "fgs_vigente"
         Caption         =   "Vigente Bs."
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
         DataField       =   "par_descripcion_larga"
         Caption         =   "      Descripción"
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
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
         EndProperty
         BeginProperty Column10 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoOrigenF 
      Height          =   375
      Left            =   885
      Top             =   2280
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   661
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
      Caption         =   "Digite ""Doble Click"" para Elegir un Registro Origen ..."
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
   Begin MSAdodcLib.Adodc AdoDestinoF 
      Height          =   375
      Left            =   885
      Top             =   5160
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   661
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
      Caption         =   "Digite ""Doble Click"" para Elegir un Registro Destino ..."
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
Attribute VB_Name = "FrmOrigenDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBuscaDes_Click()
On Error GoTo Error:
    OriDes = "D"
    varbusca = "FOR"
    For Each CAMPOS In rsformulacion.Fields
        FrmBusqueda.CmbCampo.AddItem CAMPOS.Name
    Next CAMPOS
    FrmBusqueda.Show
Exit Sub
Error:
    MsgBox "Existe error de sintaxis", vbDefaultButton2, "ERROR"

End Sub

Private Sub CmdBuscaOri_Click()
On Error GoTo Error:
    OriDes = "O"
    varbusca = "FOR"
    For Each CAMPOS In rsformulacion.Fields
        FrmBusqueda.CmbCampo.AddItem CAMPOS.Name
    Next CAMPOS
    FrmBusqueda.Show
Exit Sub
Error:
    MsgBox "Existe error de sintaxis", vbDefaultButton2, "ERROR"

End Sub

Private Sub DtgDestinoF_DblClick()
   FrmFormulacion.dtcFteT_des.Text = DtgDestinoF.Columns(0)
   FrmFormulacion.DtcOrgT_des.Text = DtgDestinoF.Columns(1)
   FrmFormulacion.dtcProT_des.Text = DtgDestinoF.Columns(2)
   FrmFormulacion.dtcPryT_des.Text = DtgDestinoF.Columns(3)
   FrmFormulacion.dtcActT_des.Text = DtgDestinoF.Columns(4)
   FrmFormulacion.dtcParT_des.Text = DtgDestinoF.Columns(5)
   
   Call define_origen
   
   Unload Me
   'FrmOrigenDestino.Visible = False
   FrmFormulacion.Text5.SetFocus

End Sub

Private Sub DtgOrigenF_DblClick()
   FrmFormulacion.dtcFteT.Text = DtgOrigenF.Columns(0)
   FrmFormulacion.DtcOrgT.Text = DtgOrigenF.Columns(1)
   FrmFormulacion.dtcProT.Text = DtgOrigenF.Columns(2)
   FrmFormulacion.dtcPryT.Text = DtgOrigenF.Columns(3)
   FrmFormulacion.dtcActT.Text = DtgOrigenF.Columns(4)
   FrmFormulacion.dtcParT.Text = DtgOrigenF.Columns(5)
   
   parametro = "fv_formulacion_gasto.fgs_modificaciones" + " >= " + "'0'"
   Call abrir_formulacionD
   
   DtgOrigenF.Enabled = False
   DtgDestinoF.Enabled = True
   
   CmdBuscaOri.Visible = False
   CmdBuscaDes.Visible = True
End Sub

Private Sub Form_Load()
    Call define_origen
	Call SeguridadSet(Me)
End Sub

Public Sub abrir_formulacionO()
  Set rsformulacion = New ADODB.Recordset       'Abrir fv_formulacion_gasto
    If rsformulacion.State = 1 Then rsformulacion.Close
    rsformulacion.Open "select * from fv_formulacion_gasto where " & parametro & " order by org_codigo, pro_proyecto, par_codigo ", db, adOpenDynamic, adLockOptimistic
    If rsformulacion.RecordCount > 0 Then
        Set adoOrigenF.Recordset = rsformulacion
        Set DtgOrigenF.DataSource = adoOrigenF.Recordset
    Else
        Set RSNADA = New ADODB.Recordset
        Set adoOrigenF.Recordset = rsformulacion
        Set DtgOrigenF.DataSource = RSNADA
    End If
End Sub

Public Sub abrir_formulacionD()
  Set rsformulacion = New ADODB.Recordset       'Abrir fv_formulacion_gasto
    If rsformulacion.State = 1 Then rsformulacion.Close
    rsformulacion.Open "select * from fv_formulacion_gasto where " & parametro & " order by org_codigo, pro_proyecto, par_codigo ", db, adOpenDynamic, adLockOptimistic
    If rsformulacion.RecordCount > 0 Then
        Set AdoDestinoF.Recordset = rsformulacion
        Set DtgDestinoF.DataSource = AdoDestinoF.Recordset
    Else
        Set RSNADA = New ADODB.Recordset
        Set AdoDestinoF.Recordset = rsformulacion
        Set DtgDestinoF.DataSource = RSNADA
    End If
End Sub

Private Sub define_origen()
    parametro = "fv_formulacion_gasto.fgs_modificaciones" + " <= " + "'0'" + " and " + "left(fv_formulacion_gasto.par_codigo,1)" + " <> " + "'1'"
    Call abrir_formulacionO
    DtgOrigenF.Enabled = True
    DtgDestinoF.Enabled = False
    
    CmdBuscaOri.Visible = True
    CmdBuscaDes.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call define_origen
End Sub

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmDesplegado 
   Caption         =   "Desplegando Cheques"
   ClientHeight    =   8535
   ClientLeft      =   180
   ClientTop       =   1815
   ClientWidth     =   11400
   Icon            =   "FrmDesplegado.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8535
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "CRITERIO PARA FILTRAR"
      Height          =   1545
      Left            =   1845
      TabIndex        =   3
      Top             =   1050
      Width           =   9480
      Begin VB.CommandButton cmdOrdenarPor 
         Caption         =   "&Ordenar por"
         Height          =   330
         Left            =   7800
         TabIndex        =   19
         Top             =   600
         Width           =   1605
      End
      Begin VB.CommandButton cmdBorrarFiltro 
         Height          =   375
         Left            =   7815
         Picture         =   "FrmDesplegado.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Borra todo el criterio de filtrado"
         Top             =   195
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txtSubtitulo 
         Height          =   315
         Left            =   1740
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   1080
         Width           =   5970
      End
      Begin VB.CommandButton cmdNuevoFiltro 
         Height          =   375
         Left            =   7230
         Picture         =   "FrmDesplegado.frx":1794
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Agrega un nuevo criterio"
         Top             =   195
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.ComboBox CmbOperador 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmDesplegado.frx":1A9E
         Left            =   3090
         List            =   "FrmDesplegado.frx":1AB7
         TabIndex        =   8
         Top             =   630
         Width           =   1065
      End
      Begin VB.ComboBox CmbCampo 
         Height          =   315
         Left            =   135
         TabIndex        =   7
         Top             =   630
         Width           =   2865
      End
      Begin VB.TextBox TxtValor 
         Height          =   315
         Left            =   4200
         TabIndex        =   9
         Top             =   630
         Width           =   3510
      End
      Begin VB.Label lblOrdenarPor 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7800
         TabIndex        =   20
         Top             =   1050
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "SubTitulo del Reporte:"
         Height          =   270
         Left            =   150
         TabIndex        =   18
         Top             =   1140
         Width           =   1650
      End
      Begin VB.Label LblValor 
         Caption         =   "Valor"
         Height          =   285
         Left            =   4335
         TabIndex        =   6
         Top             =   345
         Width           =   675
      End
      Begin VB.Label LblOperador 
         Caption         =   "Operador"
         Height          =   255
         Left            =   3105
         TabIndex        =   5
         Top             =   345
         Width           =   885
      End
      Begin VB.Label LblCampo 
         Caption         =   "Campo"
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   345
         Width           =   615
      End
   End
   Begin Crystal.CrystalReport CryHis 
      Left            =   1020
      Top             =   7230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame Frame3 
      Height          =   6000
      Left            =   345
      TabIndex        =   10
      Top             =   1050
      Width           =   1320
      Begin VB.CommandButton CmdFiltCriterio 
         Caption         =   "Filtra CRITERIO"
         Height          =   975
         Left            =   165
         Picture         =   "FrmDesplegado.frx":1AE4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1275
         Width           =   1005
      End
      Begin VB.CommandButton CmdRestaurar 
         Caption         =   "Restaurar"
         Height          =   975
         Left            =   150
         Picture         =   "FrmDesplegado.frx":1CEE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2250
         Width           =   1020
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprime Listado"
         Height          =   975
         Left            =   165
         Picture         =   "FrmDesplegado.frx":2458
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   300
         Width           =   1005
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   975
         Left            =   165
         Picture         =   "FrmDesplegado.frx":2AC2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4545
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   690
      Left            =   365
      ScaleHeight     =   630
      ScaleWidth      =   10905
      TabIndex        =   1
      Top             =   195
      Width           =   10965
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "IMPRESION HISTORICA DE CHEQUES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   420
         Left            =   1425
         TabIndex        =   2
         Top             =   135
         Width           =   8130
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   0
         Picture         =   "FrmDesplegado.frx":2F04
         Top             =   0
         Width           =   11640
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4305
      Left            =   1860
      TabIndex        =   0
      Top             =   2730
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7594
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1890
      Top             =   6870
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\MVB5\Labs\Neptuno.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\MVB5\Labs\Neptuno.mdb"
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
Attribute VB_Name = "FrmDesplegado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsbusca As New ADODB.Recordset
Dim CAMPOS As ADODB.Field

Dim Vquery As String
Dim Cadena As String
Dim vCampos(1, 30) As String
Dim auxCadenaFiltro As String
Dim vCadenaFiltro As String

Private Sub cmdNuevoFiltro_Click()
Dim i As Byte
Dim SW As Byte

    'Validacion de la entrada del filtro
    If CmbCampo.Text = "" Then
      MsgBox "Debe elegir un campo", vbInformation + vbOKOnly, "Atencion"
      Exit Sub
    End If
    If CmbOperador.Text = "" Then
      MsgBox "Debe elegir un operador", vbInformation + vbOKOnly, "Atencion"
      Exit Sub
    End If
    If TxtValor.Text = "" Then
      MsgBox "Debe ingresar un valor para completar la expresion", vbInformation + vbOKOnly, "Atencion"
      Exit Sub
    End If
    
    'Depuracion del tipo de operador
    If CmbOperador.Text = " LIKE " Then
       TxtValor = "%" & TxtValor & "%"
    End If
    
    'Depuracion de los tipos de valores
    i = 0
    SW = 0
    While vCampos(0, i) <> "" And SW = 0
      If vCampos(0, i) = CmbCampo.Text Then
          SW = 1
      Else
          i = i + 1
      End If
    Wend
    Select Case vCampos(1, i)
        Case 135, 200
          TxtValor = "'" & TxtValor & "'"
        Case 3, 5
          ' no se hace nada
    End Select
    
    'Armado de la cadena de filtro
    If vCadenaFiltro = "" Then
      vCadenaFiltro = " And " & CmbCampo.Text & CmbOperador.Text & TxtValor
    Else
      vCadenaFiltro = vCadenaFiltro & " And " & CmbCampo.Text & CmbOperador.Text & TxtValor
    End If
    auxCadenaFiltro = vCadenaFiltro
    
    'Limpiamos los combos para aramr el criterio
    CmbCampo = ""
    CmbOperador = ""
    TxtValor = ""
End Sub

Private Sub cmdBorrarFiltro_Click()
    'Inicializa combos de filtro
    CmbCampo = ""
    CmbOperador = ""
    TxtValor = ""
    'Limpia variables de criterio para filtro
    auxCadenaFiltro = ""
    vCadenaFiltro = ""
End Sub

Private Sub cmdOrdenarPor_Click()
On Error GoTo QError
    lblOrdenarPor = CmbCampo.Text
    'Limpiamos los combos para armar el criterio
    CmbCampo = ""
    CmbOperador = ""
    TxtValor = ""
    
    'Agrega el campo por el cual se va ordenar la informacion
    If lblOrdenarPor <> "" Then
        If rsbusca.State = 1 Then rsbusca.Close
        rsbusca.Open Vquery & vCadenaFiltro & " Order By " & lblOrdenarPor, db, adOpenStatic, adLockReadOnly
        Set Adodc1.Recordset = rsbusca
        Set DataGrid1.DataSource = rsbusca
    Else
        MsgBox "No existe ningun campo para ordenar la informacion!", vbCritical, "Atencion"
    End If
    Exit Sub
QError:
MsgBox "Error Inesperado. Intente de nuevo haciendo click en el boton Restaurar.", vbCritical + vbOKOnly, "Atencion"
End Sub

Private Sub CmdRestaurar_Click()
Dim i As Byte
    If FrmActivacionCheques.OptCheques = True Then
        Vquery = "Select * From cns_cheques Where cheque_o_trf='C'"
    Else
        Vquery = "Select * From cns_cheques Where cheque_o_trf='T'"
    End If
    
    If rsbusca.State = 1 Then rsbusca.Close
    rsbusca.Open Vquery, db, adOpenStatic, adLockReadOnly
    
    Set Adodc1.Recordset = rsbusca
    Set DataGrid1.DataSource = Adodc1
    
    'Inicializa combos de filtro
    CmbCampo = ""
    CmbOperador = ""
    TxtValor = ""
    
    lblOrdenarPor = ""
    
    'Limpia variables de criterio para filtro
    auxCadenaFiltro = ""
    vCadenaFiltro = ""
End Sub

Private Sub CmdFiltCriterio_Click()
Dim i As Byte
Dim SW As Byte

On Error GoTo QError

    'Validacion de la entrada del filtro
    If CmbCampo.Text = "" Then
      MsgBox "Debe elegir un campo", vbInformation + vbOKOnly, "Atencion"
      Exit Sub
    End If
    If CmbOperador.Text = "" Then
      MsgBox "Debe elegir un operador", vbInformation + vbOKOnly, "Atencion"
      Exit Sub
    End If
    If TxtValor.Text = "" Then
      MsgBox "Debe ingresar un valor para completar la expresion", vbInformation + vbOKOnly, "Atencion"
      Exit Sub
    End If
    
    'Depuracion del tipo de operador
    If CmbOperador.Text = " LIKE " Then
       TxtValor = "%" & TxtValor & "%"
    End If
    
    'Depuracion de los tipos de valores
    i = 0
    SW = 0
    While vCampos(0, i) <> "" And SW = 0
      If vCampos(0, i) = CmbCampo.Text Then
          SW = 1
      Else
          i = i + 1
      End If
    Wend
    Select Case vCampos(1, i)
        Case 135, 200
          TxtValor = "'" & TxtValor & "'"
        Case 3, 5
          ' no se hace nada
    End Select
    
    'Armado de la cadena de filtro
    If vCadenaFiltro = "" Then
      vCadenaFiltro = " And " & CmbCampo.Text & CmbOperador.Text & TxtValor
    Else
      vCadenaFiltro = vCadenaFiltro & " And " & CmbCampo.Text & CmbOperador.Text & TxtValor
    End If
    auxCadenaFiltro = vCadenaFiltro
    
    'Limpiamos los combos para armar el criterio
    CmbCampo = ""
    CmbOperador = ""
    TxtValor = ""

    If vCadenaFiltro <> "" Then
        If rsbusca.State = 1 Then rsbusca.Close
        rsbusca.Open Vquery & vCadenaFiltro, db, adOpenStatic, adLockReadOnly
        Set Adodc1.Recordset = rsbusca
        Set DataGrid1.DataSource = rsbusca
    Else
        MsgBox "No existe ningun criterio para filtrar la informacion!", vbCritical, "Atencion"
    End If
    Exit Sub
QError:
  MsgBox "Error en la construccion del criterio de filtrado. Revise", vbCritical + vbOKOnly, "Atencion"
End Sub

Private Sub cmdImprimir_Click()
Dim iResult As String
Dim Prueba As String
  If rsbusca.RecordCount <= 0 Then
    MsgBox "No existen registros para imprimir ", vbCritical + vbDefaultButton1, "Validación de datos"
    Exit Sub
  End If
      
  'Agrega el campo por el cual se va ordenar la informacion
  If lblOrdenarPor <> "" Then
    vCadenaFiltro = vCadenaFiltro & " Order By " & lblOrdenarPor
  End If

  Prueba = Vquery & vCadenaFiltro
  Screen.MousePointer = vbHourglass
  CryHis.Reset
  CryHis.WindowShowSearchBtn = True
  CryHis.Formulas(1) = "PTitulo='" & Cadena & "'"
  CryHis.Formulas(2) = "PSubtitulo='" & txtSubtitulo & "'"
  CryHis.StoredProcParam(0) = Prueba
  CryHis.ReportFileName = App.Path & "\FormsTesoreria\Operacion de Cheques\Rpt_EstadoCheques.rpt"
  iResult = CryHis.PrintReport
  Screen.MousePointer = vbDefault
  If iResult <> 0 Then
     MsgBox CryHis.LastErrorNumber & " : " & CryHis.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub cmdSalir_Click()
    FrmActivacionCheques.Show
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Byte
    If FrmActivacionCheques.OptCheques = True Then
        Vquery = "Select DISTINCT * From cns_cheques Where cheque_o_trf='C'"
        Cadena = "REPORTE HISTORICO DE CHEQUES"
    Else
        Vquery = "Select DISTINCT * From cns_cheques Where cheque_o_trf='T'"
        Cadena = "REPORTE HISTORICO DE TRANSFERENCIAS"
     End If
    auxCadenaFiltro = ""
    
    rsbusca.Open Vquery, db, adOpenStatic, adLockReadOnly
    
    Set Adodc1.Recordset = rsbusca
    Set DataGrid1.DataSource = Adodc1
    
    'Inicializa cadena de filtro
    vCadenaFiltro = ""
    auxCadenaFiltro = ""
    
    'Inicializa el combo de valores
    i = 0
    For Each CAMPOS In rsbusca.Fields
        CmbCampo.AddItem CAMPOS.Name
        vCampos(0, i) = CAMPOS.Name
        vCampos(1, i) = CAMPOS.Type
        i = i + 1
    Next CAMPOS
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If rsbusca.State = 1 Then rsbusca.Close
End Sub



VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmBuscaEnTodo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Registros"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "FrmBuscaEnTodo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraAvanzadas 
      Height          =   1560
      Left            =   15
      TabIndex        =   0
      Top             =   4335
      Width           =   5970
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Cerrar"
         Height          =   360
         Left            =   5085
         TabIndex        =   24
         Top             =   240
         Width           =   765
      End
      Begin VB.CommandButton CmdFiltCriterio 
         Caption         =   "Filtrar"
         Height          =   270
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1185
         Width           =   1110
      End
      Begin VB.CommandButton CmdRestaurar 
         Caption         =   "Restaurar"
         Height          =   270
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1185
         Width           =   1110
      End
      Begin VB.CommandButton cmdOrdenar 
         Caption         =   "Ordenar"
         Height          =   270
         Left            =   3210
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1185
         Width           =   1110
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   270
         Left            =   4755
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1185
         Width           =   1110
      End
      Begin VB.CommandButton cmdVer 
         Caption         =   "..."
         Height          =   360
         Left            =   4635
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox TxtValor 
         Height          =   315
         Left            =   3315
         TabIndex        =   5
         Top             =   630
         Width           =   2535
      End
      Begin VB.ComboBox CmbCampo 
         Height          =   315
         Left            =   135
         TabIndex        =   4
         Top             =   630
         Width           =   2250
      End
      Begin VB.ComboBox CmbOperador 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmBuscaEnTodo.frx":0ECA
         Left            =   2385
         List            =   "FrmBuscaEnTodo.frx":0EE3
         TabIndex        =   3
         Top             =   630
         Width           =   900
      End
      Begin VB.CommandButton cmdNuevoFiltro 
         Caption         =   "V"
         Height          =   360
         Left            =   3735
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Agrega un nuevo criterio"
         Top             =   240
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton cmdBorrarFiltro 
         Caption         =   "X"
         Height          =   360
         Left            =   4185
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Borra todo el criterio de filtrado"
         Top             =   240
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label LblCampo 
         BackStyle       =   0  'Transparent
         Caption         =   "Campo"
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   345
         Width           =   615
      End
      Begin VB.Label LblOperador 
         BackStyle       =   0  'Transparent
         Caption         =   "Operador"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   345
         Width           =   885
      End
      Begin VB.Label LblValor 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         Height          =   285
         Left            =   3330
         TabIndex        =   6
         Top             =   345
         Width           =   675
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   0
         Picture         =   "FrmBuscaEnTodo.frx":0F10
         Top             =   120
         Width           =   11640
      End
   End
   Begin VB.Frame FraGridPropio 
      Height          =   5895
      Left            =   15
      TabIndex        =   18
      Top             =   0
      Width           =   5955
      Begin VB.CommandButton cmdAvanzadas 
         Caption         =   "&Avanzadas"
         Height          =   270
         Left            =   105
         TabIndex        =   22
         Top             =   255
         Width           =   1110
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   270
         Left            =   1645
         TabIndex        =   21
         Top             =   255
         Width           =   1110
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   270
         Left            =   3185
         TabIndex        =   20
         Top             =   255
         Width           =   1110
      End
      Begin VB.CommandButton cmdElegir 
         Caption         =   "Elegir"
         Height          =   270
         Left            =   4725
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   255
         Width           =   1110
      End
      Begin MSDataGridLib.DataGrid GridPropio 
         Height          =   5145
         Left            =   60
         TabIndex        =   23
         Top             =   705
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   9075
         _Version        =   393216
         BackColor       =   14351870
         ForeColor       =   -2147483627
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
               LCID            =   1034
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
               LCID            =   1034
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
   End
   Begin VB.Frame fraCriterio 
      Height          =   1560
      Left            =   15
      TabIndex        =   9
      Top             =   4335
      Visible         =   0   'False
      Width           =   5955
      Begin VB.CommandButton cmdOcultar 
         Caption         =   "..."
         Height          =   360
         Left            =   5415
         TabIndex        =   13
         Top             =   210
         Width           =   435
      End
      Begin VB.TextBox txtCadenaOrdenacion 
         BackColor       =   &H00404040&
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
         Height          =   585
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   855
         Width           =   5220
      End
      Begin VB.TextBox txtCadenaFiltro 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   570
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   210
         Width           =   5220
      End
   End
End
Attribute VB_Name = "FrmBuscaEnTodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public QueryUtilizado As String
Public RecordsetTrabajo As ADODB.Recordset
Public GridTrabajo As DataGrid
Public CamposVisibles As String
Public EnGridPropio As Boolean
Public Elegido As String
Public AuxSQL As String

Dim rsRecordsetPropio As New ADODB.Recordset
Dim CAMPOS As ADODB.Field
Dim Cadena As String
Dim vCampos(2, 50) As String
Dim vCadenaFiltro As String
Dim vCadFilAux As String
Dim vCadenaOrdenacion As String
Dim vCadOrdAux As String
Dim vCadAux As String
Dim vCadenaSelectFrom As String
Dim SW As Byte
Dim vEsOrdenacion As Boolean

Private Sub cmdAvanzadas_Click()
If SW = 0 Then
  FraAvanzadas.Visible = True
  FraGridPropio.Height = FraAvanzadas.Top
  GridPropio.Height = GridPropio.Height - FraAvanzadas.Height
  SW = 1
Else
  FraAvanzadas.Visible = False
  FraGridPropio.Height = 5895
  GridPropio.Height = GridPropio.Height + FraAvanzadas.Height
  SW = 0
End If
End Sub

Private Sub cmdNuevoFiltro_Click()
Dim i As Byte
Dim SW As Byte
Dim vField As String

    'Validacion de la entrada del filtro
    If CmbCAMPO.Text = "" Then
      MsgBox "Debe elegir un campo", vbInformation + vbOKOnly, "Atencion"
      Exit Sub
    End If
    If CmbOPERADOR.Text = "" Then
      MsgBox "Debe elegir un operador", vbInformation + vbOKOnly, "Atencion"
      Exit Sub
    End If
    If TxtValor.Text = "" Then
      MsgBox "Debe ingresar un valor para completar la expresion", vbInformation + vbOKOnly, "Atencion"
      Exit Sub
    End If
    
    'Depuracion del tipo de operador
    If CmbOPERADOR.Text = " LIKE " Then
       TxtValor = "%" & TxtValor & "%"
    End If
    
    'Depuracion de los tipos de valores
    i = 0
    SW = 0
    While vCampos(0, i) <> "" And SW = 0
      If vCampos(0, i) = CmbCAMPO.Text Then
          vField = vCampos(2, i)
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
    If vCadenaFiltro = " Where " Then
      vCadenaFiltro = vCadenaFiltro & vField & CmbOPERADOR.Text & TxtValor
    Else
      vCadenaFiltro = vCadenaFiltro & " And " & vField & CmbOPERADOR.Text & TxtValor
    End If
    txtCadenaFiltro = vCadenaFiltro
    
    'Limpiamos los combos para armar el criterio
    CmbCAMPO = ""
    CmbOPERADOR = ""
    TxtValor = ""
End Sub

Private Sub cmdBorrarFiltro_Click()
    'Inicializa combos de filtro
    CmbCAMPO = ""
    CmbOPERADOR = ""
    TxtValor = ""
    'Limpia variables de criterio para filtro
    txtCadenaFiltro = ""
    txtCadenaOrdenacion = ""
    vCadenaFiltro = vCadFilAux
    vCadenaOrdenacion = vCadOrdAux
End Sub

Private Sub cmdOcultar_Click()
    fraCriterio.Visible = False
    FraAvanzadas.Visible = True
End Sub

Private Sub CmdOrdenar_Click()
Dim i As Byte
Dim SW As Byte
Dim vField As String
    vEsOrdenacion = True
    'Validacion de la entrada de la ordenacion
    If CmbCAMPO.Text = "" Then
      MsgBox "Debe elegir un campo para la ordenacion", vbInformation + vbOKOnly, "Atencion"
      Exit Sub
    End If
    
    'Depuracion de los tipos de valores
    i = 0
    SW = 0
    While vCampos(0, i) <> "" And SW = 0
      If vCampos(0, i) = CmbCAMPO.Text Then
          vField = vCampos(2, i)
          SW = 1
      Else
          i = i + 1
      End If
    Wend
    
    'Armado de la cadena de ordenacion
    If vCadenaOrdenacion = " Order By " Then
      vCadenaOrdenacion = vCadenaOrdenacion & vField
    Else
      vCadenaOrdenacion = vCadenaOrdenacion & ", " & vField
    End If
    txtCadenaOrdenacion = vCadenaOrdenacion
    
    'Realiza la ordenacion
    CmdFiltCriterio_Click
    vEsOrdenacion = False
    
    'Limpiamos los combos
    CmbCAMPO = ""
    CmbOPERADOR = ""
    TxtValor = ""
End Sub

Private Sub CmdRestaurar_Click()
    'Inicializa combos de filtro
    CmbCAMPO = ""
    CmbOPERADOR = ""
    TxtValor = ""
    'Limpia variables de criterio para filtro
    txtCadenaFiltro = ""
    txtCadenaOrdenacion = ""
    vCadenaFiltro = vCadFilAux
    vCadenaOrdenacion = vCadOrdAux
    
    If EnGridPropio = True Then
        If rsRecordsetPropio.State = 1 Then rsRecordsetPropio.Close
        rsRecordsetPropio.Open QueryUtilizado, db, adOpenStatic, adLockReadOnly
        Set GridPropio.DataSource = rsRecordsetPropio
    Else
        If RecordsetTrabajo.State = 1 Then RecordsetTrabajo.Close
        RecordsetTrabajo.Open QueryUtilizado, db, adOpenStatic, adLockReadOnly
        Set GridTrabajo.DataSource = RecordsetTrabajo
    End If
    'cmdBorrarFiltro_Click
End Sub

Private Sub CmdFiltCriterio_Click()
Dim i As Byte
Dim SW As Byte
Dim vField As String
On Error GoTo QError

    If Not vEsOrdenacion Then
        'Validacion de la entrada del filtro
        If CmbCAMPO.Text = "" Then
          MsgBox "Debe elegir un campo", vbInformation + vbOKOnly, "Atencion"
          Exit Sub
        End If
        If CmbOPERADOR.Text = "" Then
          MsgBox "Debe elegir un operador", vbInformation + vbOKOnly, "Atencion"
          Exit Sub
        End If
        If TxtValor.Text = "" Then
          MsgBox "Debe ingresar un valor para completar la expresion", vbInformation + vbOKOnly, "Atencion"
          Exit Sub
        End If
        
        'Depuracion del tipo de operador
        If CmbOPERADOR.Text = " LIKE " Then
           TxtValor = "%" & TxtValor & "%"
        End If
        
        'Depuracion de los tipos de valores
        i = 0
        SW = 0
        While vCampos(0, i) <> "" And SW = 0
          If vCampos(0, i) = CmbCAMPO.Text Then
              vField = vCampos(2, i)
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
        If vCadenaFiltro = " Where " Then
          vCadenaFiltro = vCadenaFiltro & vField & CmbOPERADOR.Text & TxtValor
        Else
          vCadenaFiltro = vCadenaFiltro & " And " & vField & CmbOPERADOR.Text & TxtValor
        End If
        txtCadenaFiltro = vCadenaFiltro
        
        'Limpiamos los combos para armar un nuevo criterio
        CmbCAMPO = ""
        CmbOPERADOR = ""
        TxtValor = ""
    End If
    
    'Se realiza el filtrado
    If Trim(UCase(vCadenaFiltro)) = "WHERE" And Trim(UCase(vCadenaOrdenacion)) = "ORDER BY" Then
        AuxSQL = vCadenaSelectFrom
    End If
    If Trim(UCase(vCadenaFiltro)) <> "WHERE" And Trim(UCase(vCadenaOrdenacion)) = "ORDER BY" Then
        AuxSQL = vCadenaSelectFrom & " " & vCadenaFiltro
    End If
    If Trim(UCase(vCadenaFiltro)) = "WHERE" And Trim(UCase(vCadenaOrdenacion)) <> "ORDER BY" Then
        AuxSQL = vCadenaSelectFrom & " " & vCadenaOrdenacion
    End If
    If Trim(UCase(vCadenaFiltro)) <> "WHERE" And Trim(UCase(vCadenaOrdenacion)) <> "ORDER BY" Then
        AuxSQL = vCadenaSelectFrom & " " & vCadenaFiltro & " " & vCadenaOrdenacion
    End If
    
    If AuxSQL <> vCadenaSelectFrom Then
        If EnGridPropio = True Then
            If rsRecordsetPropio.State = 1 Then rsRecordsetPropio.Close
            rsRecordsetPropio.Open AuxSQL, db, adOpenStatic, adLockReadOnly
            Set GridPropio.DataSource = rsRecordsetPropio
        Else
            If RecordsetTrabajo.State = 1 Then RecordsetTrabajo.Close
            RecordsetTrabajo.Open AuxSQL, db, adOpenStatic, adLockReadOnly
            Set GridTrabajo.DataSource = RecordsetTrabajo
        End If
    Else
        MsgBox "No existe ningun criterio para filtrar la informacion!", vbCritical, "Atencion"
    End If
    Exit Sub
QError:
  MsgBox "Error en la construccion del criterio de filtrado. Revise", vbCritical + vbOKOnly, "Atencion"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdElegir_Click()
    Elegido = GridPropio.Columns(0).Value
    CmdSalir_Click
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdVer_Click()
    fraCriterio.Visible = True
    FraAvanzadas.Visible = False
End Sub

Private Sub Form_Load()
Dim i As Byte
    If EnGridPropio = True Then
'        If rsRecordsetPropio.State = 1 Then rsRecordsetPropio.Close
'        rsRecordsetPropio.Open QueryUtilizado, db, adOpenStatic, adLockReadOnly
'        Set GridPropio.DataSource = rsRecordsetPropio
        'Inicializa el frame de avanzadas
        SW = 0
        FraAvanzadas.Visible = False
    Else
'        If RecordsetTrabajo.State = 1 Then RecordsetTrabajo.Close
'        RecordsetTrabajo.Open QueryUtilizado, db, adOpenStatic, adLockReadOnly
'        Set GridTrabajo.DataSource = RecordsetTrabajo
        'Inicializa el aspecto del formulario
        FraAvanzadas.Left = 15
        FraAvanzadas.Top = 0
        fraCriterio.Left = 15
        fraCriterio.Top = 0
        FraGridPropio.Visible = False
        FrmBuscaEnTodo.Height = FraAvanzadas.Height + 400
    End If
    
    'Inicializa cadena de ordenacion
    If SearchSubStr(UCase(QueryUtilizado), "ORDER BY", 0) = "" Then
      vCadenaOrdenacion = " Order By "
      vCadenaSelectFrom = QueryUtilizado
      vCadOrdAux = vCadenaOrdenacion
    Else
      vCadenaOrdenacion = SearchSubStr(UCase(QueryUtilizado), "ORDER BY", 2)
      vCadenaSelectFrom = Mid(QueryUtilizado, 1, Len(QueryUtilizado) - Len(vCadenaOrdenacion) - 1)
      vCadOrdAux = vCadenaOrdenacion
    End If
    txtCadenaOrdenacion = ""
    
    'Inicializa cadena de filtro
    If SearchSubStr(UCase(vCadenaSelectFrom), "WHERE", 0) = "" Then
      vCadenaFiltro = " Where "
      vCadFilAux = vCadenaFiltro
    Else
      vCadenaFiltro = SearchSubStr(UCase(vCadenaSelectFrom), "WHERE", 2)
      vCadenaSelectFrom = Mid(vCadenaSelectFrom, 1, Len(vCadenaSelectFrom) - Len(vCadenaFiltro) - 1)
      vCadFilAux = vCadenaFiltro
    End If
    txtCadenaFiltro = ""
    
    'Inicializa el combo de valores
    i = 0
    If EnGridPropio = True Then
        For Each CAMPOS In rsRecordsetPropio.Fields
        CmbCAMPO.AddItem CAMPOS.Name
        vCampos(0, i) = CAMPOS.Name
        vCampos(1, i) = CAMPOS.Type
        vCampos(2, i) = SearchSubStr(QueryUtilizado, CAMPOS.Name, 1)
        i = i + 1
        Next CAMPOS
    Else
        For Each CAMPOS In RecordsetTrabajo.Fields
        CmbCAMPO.AddItem CAMPOS.Name
        vCampos(0, i) = CAMPOS.Name
        vCampos(1, i) = CAMPOS.Type
        vCampos(2, i) = SearchSubStr(QueryUtilizado, CAMPOS.Name, 1)
        i = i + 1
        Next CAMPOS
    End If
    vEsOrdenacion = False
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Set QueryUtilizado = Nothing
  If EnGridPropio = False Then
    Set RecordsetTrabajo = Nothing
    Set GridTrabajo = Nothing
  End If
End Sub

Private Function SearchSubStr(SourceString As String, FindString As String, Action As Byte) As String
'Action 2= Busca y Arma Cadena Hacia Atras
'       1= Busca y Arma Cadena Hacia Atras
'       0= Solo Busca

Dim l As Integer
Dim k As Integer
Dim c As Integer
Dim p As Integer
Dim s As String
Dim SW As Integer

  SearchSubStr = ""
  l = Len(SourceString)
  k = Len(FindString)
  c = 1  'Contador de reccorrido
  SW = 0 'No Encontrado
  While c <= l - k And SW = 0
    If Mid(SourceString, c, k) = FindString Then
        If Action = 1 Then
            p = c - 1
            s = ""
            While Mid(SourceString, p, 1) <> "," And Mid(SourceString, p, 1) <> " " And UCase(Mid(SourceString, p - 3, 4)) <> " AS "
              s = Mid(SourceString, p, 1) & s
              p = p - 1
            Wend
        End If
        SW = 1
    Else
      c = c + 1
    End If
  Wend
  Select Case Action
      Case 2
        SearchSubStr = Mid(SourceString, c, Len(SourceString) - c + 1)
      Case 1
        SearchSubStr = s & FindString
      Case 0
        If SW = 1 Then
            SearchSubStr = FindString
        Else
            SearchSubStr = ""
        End If
  End Select
End Function

Private Sub GridPropio_DblClick()
    cmdElegir_Click
End Sub

VERSION 5.00
Begin VB.Form GrFrmBuscaEnGridExterno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elige el criterio de Busqueda"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   DrawStyle       =   1  'Dash
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GrFrmBuscaEnGridExterno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "GrFrmBuscaEnGridExterno.frx":0A02
   ScaleHeight     =   1635
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicCampos 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11580
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   11520
      ScaleWidth      =   5100
      TabIndex        =   6
      Top             =   0
      Width           =   5160
      Begin VB.ComboBox CmbCompara 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "GrFrmBuscaEnGridExterno.frx":CA44
         Left            =   2160
         List            =   "GrFrmBuscaEnGridExterno.frx":CA5A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtCompara 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   3360
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox CmbCampo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "GrFrmBuscaEnGridExterno.frx":CA75
         Left            =   120
         List            =   "GrFrmBuscaEnGridExterno.frx":CA77
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.PictureBox ImlImagenesAv 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5280
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   8
         Top             =   240
         Width           =   1200
      End
      Begin VB.PictureBox TbrAvanzadas 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   360
         Negotiate       =   -1  'True
         ScaleHeight     =   585
         ScaleWidth      =   4290
         TabIndex        =   7
         Top             =   840
         Width           =   4350
         Begin VB.CommandButton BtnBuscaSig 
            Caption         =   "Buscar Siguiente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   9360
            TabIndex        =   15
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton BtnBuscaAnterior 
            Caption         =   "Buscar Anterior"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   7920
            TabIndex        =   14
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton BtnBusca1ro 
            Caption         =   "Buscar Primero"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6480
            TabIndex        =   13
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton BtnOrdenar 
            Caption         =   "Ordenar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   5040
            TabIndex        =   12
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton BtnCerrar 
            BackColor       =   &H00808080&
            Height          =   615
            Left            =   2880
            Picture         =   "GrFrmBuscaEnGridExterno.frx":CA79
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Width           =   1455
         End
         Begin VB.CommandButton BtnRefrecar 
            BackColor       =   &H00808080&
            Caption         =   "Refrescar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1440
            Picture         =   "GrFrmBuscaEnGridExterno.frx":3668B
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   0
            Width           =   1455
         End
         Begin VB.CommandButton BtnFiltrar 
            BackColor       =   &H00808080&
            Caption         =   "Filtrar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            Picture         =   "GrFrmBuscaEnGridExterno.frx":36C15
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
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
         Height          =   240
         Left            =   3645
         TabIndex        =   0
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operador:"
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
         Height          =   240
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Columna:"
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
         Height          =   195
         Left            =   345
         TabIndex        =   2
         Top             =   120
         Width           =   1305
      End
   End
End
Attribute VB_Name = "GrFrmBuscaEnGridExterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************************
'**     PROPOSITO GENERAL: Formulario de Búsqueda y Filtrado     **
'**     AUTOR: Dulfredo Rojas Valencia                           **
'**     FECHA CREACION: 21/05/00                                 **
'**     FECHA ULTIMA REVISION: 13/07/00                          **
'**     CONVERSION A COMPONENTE: 1/06/2000                       **
'******************************************************************
Option Explicit
'*** Ventana Top Most
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
       ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOP = 0
Const HWND_TOPMOST = -1
Const HWND_NOTTOPMOST = -2
Const SWP_NOSIZE = 1
Const SWP_NOMOVE = 2
Const SWP_NOZORDER = &H4
Const SWP_NOOWNERZORDER = &H200      '  No usar el orden Z del propietario
Dim flags As Integer
Dim Resultado As Long
'*** Manejo de filtrado y busqueda
Dim ListaCampos() As String
Dim Encontro As Boolean
'Dim PQConexion As ADODB.Connection
'Dim rsTablaAux As ADODB.Recordset
Dim OrdenarAsc As Boolean
Dim ElQueryOriginal As String
Dim NuevoQuery As String
Dim EsTrueDBGridAux As Boolean
'Dim ElTDBGridAux As TrueOleDBGrid60.TDBGrid
'Dim ElGridAux As MSDataGridLib.DataGrid
Dim PQTipoRs As Integer
Dim PQTipoBloqueo As Integer

Private Sub OpcionRefrescar()
  On Error GoTo RefErr
  Screen.MousePointer = vbHourglass
  NuevoQuery = ElQueryOriginal
  If rsTablaAux.State = 1 Then rsTablaAux.Close
  'rsTablaAux.Close
  rsTablaAux.Open NuevoQuery, PQConexion, PQTipoRs, PQTipoBloqueo
  If EsTrueDBGridAux Then
    Set ElTDBGridAux.DataSource = rsTablaAux
  Else
    Set ElGridAux.DataSource = rsTablaAux
  End If
  Screen.MousePointer = vbDefault
  Exit Sub
RefErr:
  MsgBox "Error:" & Err & " " & Err.Description, vbInformation + vbOKOnly, "Atención"
End Sub

Private Sub OpcionFiltrar()
On Error GoTo Que_Error
Dim NuevoFiltro As String
  If (CmbCampo.Text = "") Then 'Or (TxtCompara.Text = "") Then
    MsgBox "Debe elegir el campo que quiere filtrar y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
  Else
    Screen.MousePointer = vbHourglass
    If CmbCompara.Text = "" Then CmbCompara.Text = "="
    If TxtCompara.Text = "" Then CmbCompara.Text = "="
    'Si es Cadena
    If EsTipoCadena(rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Type) Then
      If CmbCompara.Text = "Como" Then
        NuevoFiltro = "(" & NombreCampo(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex))) & " LIKE '%" & TxtCompara & "%')"
      Else
        NuevoFiltro = "(" & NombreCampo(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex))) & " " & CmbCompara.Text & " '" & TxtCompara & "')"
      End If
    Else 'Si es Número
      If EsTipoNumerico(rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Type) Then
        If TxtCompara.Text <> "" Then If Not IsNumeric(TxtCompara.Text) Then MsgBox "Debe ingresar una Valor válido (Numérico)...", vbExclamation + vbOKOnly, "Atención": Screen.MousePointer = vbDefault: Exit Sub
        If CmbCompara.Text = "Como" Then CmbCompara.Text = "="
        NuevoFiltro = "(" & NombreCampo(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex))) & " " & CmbCompara.Text & " " & TxtCompara & ")"
      Else 'Si es fecha
        If EsTipoFechaHora(rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Type) Then
          If TxtCompara.Text <> "" Then If Not IsDate(TxtCompara.Text) Then MsgBox "Debe ingresar una Valor válido (Fecha)...", vbExclamation + vbOKOnly, "Atención": Screen.MousePointer = vbDefault: Exit Sub
          If CmbCompara.Text = "Como" Then CmbCompara.Text = "="
          NuevoFiltro = "(" & NombreCampo(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex))) & " " & CmbCompara.Text & " '" & TxtCompara & "')"
        Else   'Otro Tipo
          NuevoFiltro = "(" & NombreCampo(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex))) & " = " & TxtCompara & ")"
        End If
      End If
    End If
    If (TxtCompara.Text = "") Then NuevoFiltro = "(" & NuevoFiltro & "or(" & NombreCampo(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex))) & " is Null))"
    'Ve si tiene WHERE
    If InStr(1, UCase(NuevoQuery), "WHERE") <> 0 Then
      NuevoQuery = NuevoQuery & " and " & NuevoFiltro
    Else
      NuevoQuery = NuevoQuery & " WHERE " & NuevoFiltro
    End If
    rsTablaAux.Close
    rsTablaAux.Open NuevoQuery, PQConexion, PQTipoRs, PQTipoBloqueo
    If EsTrueDBGridAux Then
      Set ElTDBGridAux.DataSource = rsTablaAux
    Else
      Set ElGridAux.DataSource = rsTablaAux
    End If
    rsTablaAux.Requery
    Screen.MousePointer = vbDefault
  End If
  Exit Sub
Que_Error:
  Screen.MousePointer = vbDefault
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Atención"
End Sub

Private Sub OpcionOrdenar(Ascendente As Boolean)
Dim AuxCampo As String
On Error GoTo Que_Error
  If (CmbCampo.Text = "") Then
    MsgBox "Debe elegir el campo por el que quiere ordenar", vbInformation + vbOKOnly, "Error de procedimiento"
  Else
    Screen.MousePointer = vbHourglass
    'check for the use of the ctrl key for descending sort
    AuxCampo = NombreCampo(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))   ' NombreCampo(CmbCampo.Text)
    If Not Ascendente Then
      rsTablaAux.Sort = AuxCampo & " DESC"
    Else
      rsTablaAux.Sort = AuxCampo & " ASC"
    End If
    Screen.MousePointer = vbDefault
  End If
  Exit Sub
Que_Error:
  Screen.MousePointer = vbDefault
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Atención"
End Sub

Private Sub OpcionBuscarPrimero()
On Error GoTo QueError
Dim Marca As Variant
  Screen.MousePointer = vbHourglass
  Marca = rsTablaAux.Bookmark
  rsTablaAux.MoveFirst
  If EsTipoCadena(rsTablaAux.Fields(NombreCampo(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ItemData(CmbCampo.ListIndex)))))).Type) Then
    Select Case CmbCompara.Text
      Case "=", ">", ">=", "<", "<="
        rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " " & CmbCompara.Text & " '" & TxtCompara.Text & "'", 0, adSearchForward
      Case "Como"
        rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " LIKE '%" & TxtCompara.Text & "%'", 0, adSearchForward
    End Select
  Else
    If EsTipoNumerico(rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ItemData(CmbCampo.ListIndex))))).Type) Then
      If TxtCompara.Text <> "" Then If Not IsNumeric(TxtCompara.Text) Then MsgBox "Debe ingresar una Valor válido (Numérico)...", vbExclamation + vbOKOnly, "Atención": Screen.MousePointer = vbDefault: Exit Sub
      Select Case CmbCompara.Text
        Case "=", ">", ">=", "<", "<="
          rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " " & CmbCompara.Text & " " & TxtCompara.Text, 0, adSearchForward
        Case "Como"
          rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ItemData(CmbCampo.ListIndex))))).Name & " LIKE " & TxtCompara.Text, 0, adSearchForward
      End Select
    Else
      If CmbCompara.Text = "Como" Then CmbCompara.Text = "="
      rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " " & CmbCompara.Text & " " & TxtCompara.Text, 0, adSearchForward
    End If
  End If
  If rsTablaAux.EOF Then Encontro = False Else Encontro = True
  Screen.MousePointer = vbDefault
  If Not Encontro Then
    MsgBox "No existe ninguna coincidencia", vbInformation + vbOKOnly, "No encontrado"
    rsTablaAux.Bookmark = Marca
  End If
  Exit Sub
QueError:
  Screen.MousePointer = vbDefault
  If Err.Number = 3021 Then
    MsgBox "No existe ningún registro activo!", vbInformation + vbOKOnly, "Atención"
  Else
    MsgBox Err.Number & ", " & Err.Description
  End If
End Sub

Private Sub OpcionBuscarAnterior()
On Error GoTo QueError
Dim Marca As Variant
  Screen.MousePointer = vbHourglass
  Marca = rsTablaAux.Bookmark
  rsTablaAux.MovePrevious
  If rsTablaAux.BOF Then rsTablaAux.MoveFirst
  If EsTipoCadena(rsTablaAux.Fields(NombreCampo(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex))))).Type) Then
    Select Case CmbCompara.Text
      Case "=", ">", ">=", "<", "<="
          rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " " & CmbCompara.Text & " '" & TxtCompara.Text & "'", 0, adSearchBackward
      Case "Como"
          rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " LIKE '%" & TxtCompara.Text & "%'", 0, adSearchBackward
    End Select
  Else
    If EsTipoNumerico(rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Type) Then
      If TxtCompara.Text <> "" Then If Not IsNumeric(TxtCompara.Text) Then MsgBox "Debe ingresar una Valor válido (Numérico)...", vbExclamation + vbOKOnly, "Atención": Screen.MousePointer = vbDefault: Exit Sub
      Select Case CmbCompara.Text
        Case "=", ">", ">=", "<", "<="
            rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " " & CmbCompara.Text & " " & TxtCompara.Text, 0, adSearchBackward
        Case "Como"
            rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " LIKE " & TxtCompara.Text, 0, adSearchBackward
      End Select
    Else
      If CmbCompara.Text = "Como" Then CmbCompara.Text = "="
      rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " " & CmbCompara.Text & " " & TxtCompara.Text, 0, adSearchBackward
    End If
  End If
  If rsTablaAux.BOF Then Encontro = False Else Encontro = True
  Screen.MousePointer = vbDefault
  If Not Encontro Then
    MsgBox "No existe ninguna coincidencia", vbInformation + vbOKOnly, "No encontrado"
    rsTablaAux.Bookmark = Marca
  End If
  Exit Sub
QueError:
  Screen.MousePointer = vbDefault
  If Err.Number = 3021 Then
    MsgBox "No existe ningún registro activo!", vbInformation + vbOKOnly, "Atención"
  Else
    MsgBox Err.Number & ", " & Err.Description
  End If
End Sub

Private Sub OpcionBuscarSiguiente()
On Error GoTo QueError
Dim Marca As Variant
  Screen.MousePointer = vbHourglass
  Marca = rsTablaAux.Bookmark
  rsTablaAux.MoveNext
  If rsTablaAux.EOF Then rsTablaAux.MoveLast
  If EsTipoCadena(rsTablaAux.Fields(NombreCampo(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex))))).Type) Then
    Select Case CmbCompara.Text
      Case "=", ">", ">=", "<", "<="
          rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " " & CmbCompara.Text & " '" & TxtCompara.Text & "'", 0, adSearchForward
      Case "Como"
          rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " LIKE '%" & TxtCompara.Text & "%'", 0, adSearchForward
    End Select
  Else
    If EsTipoNumerico(rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Type) Then
      If TxtCompara.Text <> "" Then If Not IsNumeric(TxtCompara.Text) Then MsgBox "Debe ingresar una Valor válido (Numérico)...", vbExclamation + vbOKOnly, "Atención": Screen.MousePointer = vbDefault: Exit Sub
      Select Case CmbCompara.Text
        Case "=", ">", ">=", "<", "<="
          rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " " & CmbCompara.Text & " " & TxtCompara.Text, 0, adSearchForward
        Case "Como"
            rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " LIKE " & TxtCompara.Text, 0, adSearchForward
      End Select
    Else
      If CmbCompara.Text = "Como" Then CmbCompara.Text = "="
      rsTablaAux.Find rsTablaAux.Fields(QuitaAntesPunto(ListaCampos(CmbCampo.ItemData(CmbCampo.ListIndex)))).Name & " " & CmbCompara.Text & " " & TxtCompara.Text, 0, adSearchForward
    End If
  End If
  If rsTablaAux.EOF Then Encontro = False Else Encontro = True
  Screen.MousePointer = vbDefault
  If Not Encontro Then
    MsgBox "No existe ninguna coincidencia", vbInformation + vbOKOnly, "No encontrado"
    rsTablaAux.Bookmark = Marca
  End If
  Exit Sub
QueError:
  Screen.MousePointer = vbDefault
  If Err.Number = 3021 Then
    MsgBox "No existe ningún registro activo!", vbInformation + vbOKOnly, "Atención"
  Else
    MsgBox Err.Number & ", " & Err.Description
  End If
End Sub

'**************** CONTROLES **********************
Private Sub OpcionSalir()

    buscados = 0
  Unload Me
End Sub

Private Sub BtnBusca1ro_Click()
    If (CmbCampo.Text = "") Or (CmbCompara.Text = "") Or (TxtCompara.Text = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe elegir la columna por la que quiere buscar, el operador y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
        CmbCampo.SetFocus
        Exit Sub
    End If
    OpcionBuscarPrimero
End Sub

Private Sub BtnBuscaAnterior_Click()
    If (CmbCampo.Text = "") Or (CmbCompara.Text = "") Or (TxtCompara.Text = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe elegir la columna por la que quiere buscar, el operador y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
        CmbCampo.SetFocus
        Exit Sub
    End If
    OpcionBuscarAnterior
End Sub

Private Sub BtnBuscaSig_Click()
    If (CmbCampo.Text = "") Or (CmbCompara.Text = "") Or (TxtCompara.Text = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe elegir la columna por la que quiere buscar, el operador y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
        CmbCampo.SetFocus
        Exit Sub
    End If
    OpcionBuscarSiguiente
End Sub

Private Sub BtnCerrar_Click()
    OpcionSalir
End Sub

Private Sub BtnFiltrar_Click()
    'OpcionRefrescar
    OpcionFiltrar
    queryinicial99 = NuevoQuery
    GlSqlAux = NuevoQuery
End Sub

Private Sub BtnOrdenar_Click()
    If (CmbCampo.Text = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe elegir la columna por la que quiere ordenar.", vbInformation + vbOKOnly, "Error de Procedimiento"
        CmbCampo.SetFocus
        Exit Sub
      End If
      OpcionOrdenar OrdenarAsc
End Sub

Private Sub BtnRefrecar_Click()
    'OpcionFiltrar
    OpcionRefrescar
End Sub

Private Sub Form_Load()
  OrdenarAsc = True
  flags = SWP_NOSIZE Or SWP_NOMOVE 'Or SWP_NOZORDER  'SWP_NOOWNERZORDER
'
  Resultado = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Resultado = SetWindowPos(Me.hwnd, HWND_NOTTOPMOST, 0, 0, 0, 0, flags)
  OpcionSalir
End Sub

Private Sub TxtCompara_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
   OpcionBuscarPrimero
 End If
End Sub

'Private Sub TbrAvanzadas_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'  If ButtonMenu.Parent.Key = "Ordenar" Then
'    If (CmbCampo.Text = "") Then
'      Screen.MousePointer = vbDefault
'      MsgBox "Debe elegir la columna por la que quiere ordenar.", vbInformation + vbOKOnly, "Error de Procedimiento"
'      CmbCampo.SetFocus
'      Exit Sub
'    End If
'    Select Case ButtonMenu.Key
'      Case "Asc"
'        TbrAvanzadas.Buttons(5).Image = 3
'        OrdenarAsc = True
'      Case "Desc"
'        TbrAvanzadas.Buttons(5).Image = 4
'        OrdenarAsc = False
'    End Select
'    OpcionOrdenar OrdenarAsc
'  End If
'  If ButtonMenu.Parent.Key = "Buscar" Then
'    If (CmbCampo.Text = "") Or (CmbCompara.Text = "") Or (TxtCompara.Text = "") Then
'      Screen.MousePointer = vbDefault
'      MsgBox "Debe elegir la columna por la que quiere buscar, el operador y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
'      CmbCampo.SetFocus
'      Exit Sub
'    End If
'    Select Case ButtonMenu.Key
'      Case "Primero"
'        TbrAvanzadas.Buttons(7).Image = 5
'      Case "Anterior"
'        TbrAvanzadas.Buttons(7).Image = 7
'      Case "Siguiente"
'        TbrAvanzadas.Buttons(7).Image = 6
'    End Select
'    TbrAvanzadas.Buttons(7).Caption = ButtonMenu.Text
'    TbrAvanzadas_ButtonClick TbrAvanzadas.Buttons(7)
'  End If
'End Sub

'Private Sub TbrAvanzadas_ButtonClick(ByVal Button As MSComctlLib.Button)
'  Select Case Button.Key
'    Case "Refrescar"
'      OpcionRefrescar
'    Case "Filtrar"
'      OpcionFiltrar
'    Case "Ordenar"
'      If (CmbCampo.Text = "") Then
'        Screen.MousePointer = vbDefault
'        MsgBox "Debe elegir la columna por la que quiere ordenar.", vbInformation + vbOKOnly, "Error de Procedimiento"
'        CmbCampo.SetFocus
'        Exit Sub
'      End If
'      OpcionOrdenar OrdenarAsc
'    Case "Buscar"
'      If (CmbCampo.Text = "") Or (CmbCompara.Text = "") Or (TxtCompara.Text = "") Then
'        Screen.MousePointer = vbDefault
'        MsgBox "Debe elegir la columna por la que quiere buscar, el operador y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
'        CmbCampo.SetFocus
'        Exit Sub
'      End If
'      Select Case Button.Caption
'        Case "&Primero"
'          OpcionBuscarPrimero
'        Case "&Anterior"
'          OpcionBuscarAnterior
'        Case "&Siguiente"
'          OpcionBuscarSiguiente
'      End Select
'    Case "Salir"
'      OpcionSalir
'  End Select
'End Sub

''********************** Procedimientos ********************
Public Sub GrPrincipal(QConexion As Object, _
                       rsTabla As ADODB.Recordset, _
                       QTipoRs As Integer, _
                       QTipoBloqueo As Integer, _
                       ElQuery As String, _
                       ElGrid As Object, _
                       Optional EsTdbGrid As Boolean = False, _
                       Optional Titulo As String = "Realice su Elección", _
                       Optional CamposVisibles As String = "", _
                       Optional CDefecto As String)
'ElQuery,        El Query inicial que tiene el recordset
'ElGrid          El Grid que se está utilizando
'rsTablaAux,     Tabla con la que se tiene que trabajar
'EsTDBGrid,      Si se está utilizando un TrueDBGrid, o el DataGrid normal
'Título,         Titulo de Venta, "" para titulo por defecto
'CamposVisibles, Cadena de 1 y 0 que representan cada columna del recordset
'                1: Visible; 0 no Visible
'CDefecto,       Campo de busqueda que se activa por defecto
Dim i As Integer
Dim Linea As Integer
On Error GoTo LoadErr
  Screen.MousePointer = vbHourglass
  Linea = 0
  'Cargar a variables locales, toda la información
  ElQueryOriginal = ElQuery
  Linea = 1
  NuevoQuery = ElQuery
  Linea = 2
  Set PQConexion = New ADODB.Connection
  Linea = 3
  Set PQConexion = QConexion
  Linea = 4
  EsTrueDBGridAux = EsTdbGrid
  Linea = 5
  Set rsTablaAux = New ADODB.Recordset
  Linea = 6
  Set rsTablaAux = rsTabla
  Linea = 7
  PQTipoRs = QTipoRs
  Linea = 8
  PQTipoBloqueo = QTipoBloqueo
  Linea = 9
  If EsTrueDBGridAux Then
    Set ElTDBGridAux = ElGrid
  Else
    Set ElGridAux = ElGrid
  End If
  Linea = 10
  If EsTrueDBGridAux Then
    Linea = 11
    ReDim ListaCampos(ElTDBGridAux.Columns.Count)
    Linea = 12
    For i = 0 To ElTDBGridAux.Columns.Count - 1
      If Len(CamposVisibles) >= i + 1 Then
        If Mid(CamposVisibles, i + 1, 1) = "1" Then
          CmbCampo.AddItem ElTDBGridAux.Columns(i).Caption
          CmbCampo.ItemData(CmbCampo.NewIndex) = i
          ListaCampos(i) = NombreCampoQuery(ElTDBGridAux.Columns(i).DataField)
        End If
      Else
        CmbCampo.AddItem ElTDBGridAux.Columns(i).Caption
        CmbCampo.ItemData(CmbCampo.NewIndex) = i
        ListaCampos(i) = NombreCampoQuery(ElTDBGridAux.Columns(i).DataField)
      End If
    Next i
    Linea = 13
  Else
    Linea = 14
    ReDim ListaCampos(ElGridAux.Columns.Count)
    Linea = 15
    For i = 0 To ElGridAux.Columns.Count - 1
      If Len(CamposVisibles) >= i + 1 Then
        If Mid(CamposVisibles, i + 1, 1) = "1" Then
          CmbCampo.AddItem ElGridAux.Columns(i).Caption
          CmbCampo.ItemData(CmbCampo.NewIndex) = i
          ListaCampos(i) = NombreCampoQuery(ElGridAux.Columns(i).DataField)
        End If
      Else
        CmbCampo.AddItem ElGridAux.Columns(i).Caption
        CmbCampo.ItemData(CmbCampo.NewIndex) = i
        ListaCampos(i) = NombreCampoQuery(ElGridAux.Columns(i).DataField)
      End If
    Next i
    Linea = 15
  End If
  Linea = 16
  If Titulo <> "" Then Me.Caption = Titulo
  Linea = 17
  On Error Resume Next
  Linea = 18
  CmbCompara.Text = "Como"
  Linea = 19
  If CDefecto <> "" Then CmbCampo.Text = CDefecto
  Linea = 20
  Screen.MousePointer = vbDefault
  Linea = 21
  Me.Show  ' Muestra el formulario
  Exit Sub
LoadErr:
  Screen.MousePointer = vbDefault
  MsgBox "Error al mostrar Formulario:" & Err & " " & Err.Description & vbCrLf & "En Linea: " & Linea, vbInformation + vbOKOnly, "Atención"
  Select Case MsgBox("¿ Anula, Reintenta o Ignora ?", vbCritical + vbAbortRetryIgnore, "Error!")
    Case vbRetry
      Resume
    Case vbIgnore
      Resume Next
  End Select
  Unload Me
End Sub

Private Function NombreCampo(CampoBusca As String) As String
  If InStr(1, CampoBusca, " ") Then
    NombreCampo = "[" & CampoBusca & "]"
'    NombreCampo = rsTablaAux.Fields(DtgElige.Col).Name
  Else
    NombreCampo = CampoBusca
  End If
End Function

Private Function EsTipoCadena(QueTipo As Byte) As Boolean
  If (QueTipo = adChar) Or (QueTipo = adVarChar) Or _
       (QueTipo = adLongVarChar) Or (QueTipo = adLongVarWChar) Or _
       (QueTipo = adVarWChar) Or (QueTipo = adWChar) Then
    EsTipoCadena = True
  Else
    EsTipoCadena = False
  End If
End Function

Private Function EsTipoNumerico(QueTipo As Byte) As Boolean
  If (QueTipo = adBigInt) Or (QueTipo = adCurrency) Or _
     (QueTipo = adDecimal) Or (QueTipo = adDouble) Or _
     (QueTipo = adInteger) Or (QueTipo = adNumeric) Or _
     (QueTipo = adSingle) Or (QueTipo = adSmallInt) Or _
     (QueTipo = adTinyInt) Or (QueTipo = adUnsignedBigInt) Or _
     (QueTipo = adUnsignedInt) Or (QueTipo = adUnsignedSmallInt) Or _
     (QueTipo = adUnsignedTinyInt) Or (QueTipo = adVarNumeric) Then
    EsTipoNumerico = True
  Else
    EsTipoNumerico = False
  End If
End Function

Private Function EsTipoFechaHora(QueTipo As Byte) As Boolean
  If (QueTipo = adDate) Or (QueTipo = adDBDate) Or _
     (QueTipo = adDBTime) Or _
     (QueTipo = adDBTimeStamp) Then
    EsTipoFechaHora = True
  Else
    EsTipoFechaHora = False
  End If
End Function

Private Function NombreCampoQuery(NombreDelCampo As String) As String
Dim AuxQuery As String
Dim PosCampo As Integer
Dim i As Integer
Dim Encontro As Boolean
  AuxQuery = Mid(ElQueryOriginal, 1, InStr(1, ElQueryOriginal, "FROM", vbTextCompare) - 1)
  PosCampo = InStr(1, AuxQuery, NombreDelCampo, vbTextCompare)
  If PosCampo = 0 Then
    NombreCampoQuery = NombreDelCampo
  Else
    If Mid(AuxQuery, PosCampo - 1, 1) = "." Then
      i = PosCampo - 1
      Encontro = False
      While Not Encontro
        If (Mid(AuxQuery, i, 1) = ",") Or (Mid(AuxQuery, i, 1) = " ") Then
          Encontro = True
        Else
          i = i - 1
        End If
      Wend
      NombreCampoQuery = Mid(AuxQuery, i + 1, PosCampo - i - 1) & NombreDelCampo
    Else
      NombreCampoQuery = NombreDelCampo
    End If
  End If
End Function

Private Function QuitaAntesPunto(NombreDelCampo As String) As String
Dim AuxPos As Integer
  AuxPos = InStr(1, NombreDelCampo, ".", vbTextCompare)
  If AuxPos <> 0 Then
    QuitaAntesPunto = Mid(NombreDelCampo, AuxPos + 1, Len(NombreDelCampo) - AuxPos)
  Else
    QuitaAntesPunto = NombreDelCampo
  End If
End Function


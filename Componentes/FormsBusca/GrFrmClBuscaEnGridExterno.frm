VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form GrFrmClBuscaEnGridExterno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Realice su Elección"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   DrawStyle       =   1  'Dash
   Icon            =   "GrFrmClBuscaEnGridExterno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar TbrAvanzadas 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   735
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   635
      ButtonWidth     =   2143
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImlImagenesAv"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refrescar"
            Key             =   "Refrescar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Filtrar  "
            Key             =   "Filtrar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ordenar"
            Key             =   "Ordenar"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Asc"
                  Text            =   "&Ascendentemente"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Desc"
                  Text            =   "&Descendentemente"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Primero"
            Key             =   "Buscar"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Primero"
                  Text            =   "&Primero"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Anterior"
                  Text            =   "&Anterior"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Siguiente"
                  Text            =   "&Siguiente"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicCampos 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      Picture         =   "GrFrmClBuscaEnGridExterno.frx":27A2
      ScaleHeight     =   675
      ScaleWidth      =   5565
      TabIndex        =   6
      Top             =   0
      Width           =   5625
      Begin VB.CommandButton CmdListo 
         Height          =   680
         Left            =   4800
         Picture         =   "GrFrmClBuscaEnGridExterno.frx":5A3C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
      Begin MSComctlLib.ImageList ImlImagenesAv 
         Left            =   3120
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridExterno.frx":5D46
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridExterno.frx":5EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridExterno.frx":5FFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridExterno.frx":6154
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridExterno.frx":62AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridExterno.frx":6408
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridExterno.frx":6562
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox CmbCompara 
         Height          =   315
         ItemData        =   "GrFrmClBuscaEnGridExterno.frx":66BC
         Left            =   1800
         List            =   "GrFrmClBuscaEnGridExterno.frx":66D2
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtCompara 
         Height          =   324
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox CmbCampo 
         Height          =   315
         ItemData        =   "GrFrmClBuscaEnGridExterno.frx":66ED
         Left            =   120
         List            =   "GrFrmClBuscaEnGridExterno.frx":66EF
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   2760
         TabIndex        =   0
         Top             =   0
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operador:"
         Height          =   195
         Left            =   1800
         TabIndex        =   4
         Top             =   0
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Columna:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   660
      End
   End
End
Attribute VB_Name = "GrFrmClBuscaEnGridExterno"
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
       ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
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
Dim PQConexion As ADODB.Connection
Dim rsTablaAux As ADODB.Recordset
Dim OrdenarAsc As Boolean
Dim ElQueryOriginal As String
Dim NuevoQuery As String
Dim EsTrueDBGridAux As Boolean
Dim ElTDBGridAux As TrueOleDBGrid60.TDBGrid
Dim ElGridAux As MSDataGridLib.DataGrid
Dim PQTipoRs As Integer
Dim PQTipoBloqueo As Integer

Private Sub OpcionRefrescar()
  On Error GoTo RefErr
  Screen.MousePointer = vbHourglass
  NuevoQuery = ElQueryOriginal
  rsTablaAux.Close
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

Private Sub CmdListo_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  OrdenarAsc = True
  flags = SWP_NOSIZE Or SWP_NOMOVE 'Or SWP_NOZORDER  'SWP_NOOWNERZORDER
'
  Resultado = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Resultado = SetWindowPos(Me.hwnd, HWND_NOTTOPMOST, 0, 0, 0, 0, flags)
End Sub

Private Sub TxtCompara_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
   OpcionBuscarPrimero
 End If
End Sub

Private Sub TbrAvanzadas_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  If ButtonMenu.Parent.Key = "Ordenar" Then
    If (CmbCampo.Text = "") Then
      Screen.MousePointer = vbDefault
      MsgBox "Debe elegir la columna por la que quiere ordenar.", vbInformation + vbOKOnly, "Error de Procedimiento"
      CmbCampo.SetFocus
      Exit Sub
    End If
    Select Case ButtonMenu.Key
      Case "Asc"
        TbrAvanzadas.Buttons(5).Image = 3
        OrdenarAsc = True
      Case "Desc"
        TbrAvanzadas.Buttons(5).Image = 4
        OrdenarAsc = False
    End Select
    OpcionOrdenar OrdenarAsc
  End If
  If ButtonMenu.Parent.Key = "Buscar" Then
    If (CmbCampo.Text = "") Or (CmbCompara.Text = "") Or (TxtCompara.Text = "") Then
      Screen.MousePointer = vbDefault
      MsgBox "Debe elegir la columna por la que quiere buscar, el operador y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
      CmbCampo.SetFocus
      Exit Sub
    End If
    Select Case ButtonMenu.Key
      Case "Primero"
        TbrAvanzadas.Buttons(7).Image = 5
      Case "Anterior"
        TbrAvanzadas.Buttons(7).Image = 7
      Case "Siguiente"
        TbrAvanzadas.Buttons(7).Image = 6
    End Select
    TbrAvanzadas.Buttons(7).Caption = ButtonMenu.Text
    TbrAvanzadas_ButtonClick TbrAvanzadas.Buttons(7)
  End If
End Sub

Private Sub TbrAvanzadas_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "Refrescar"
      OpcionRefrescar
    Case "Filtrar"
      OpcionFiltrar
    Case "Ordenar"
      If (CmbCampo.Text = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe elegir la columna por la que quiere ordenar.", vbInformation + vbOKOnly, "Error de Procedimiento"
        CmbCampo.SetFocus
        Exit Sub
      End If
      OpcionOrdenar OrdenarAsc
    Case "Buscar"
      If (CmbCampo.Text = "") Or (CmbCompara.Text = "") Or (TxtCompara.Text = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe elegir la columna por la que quiere buscar, el operador y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
        CmbCampo.SetFocus
        Exit Sub
      End If
      Select Case Button.Caption
        Case "&Primero"
          OpcionBuscarPrimero
        Case "&Anterior"
          OpcionBuscarAnterior
        Case "&Siguiente"
          OpcionBuscarSiguiente
      End Select
  End Select
End Sub

'********************** Procedimientos ********************
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
'rsTablaAux,     Tabla con la que se tiene que trabajar
'ElQuery,        El Query inicial que tiene el recordset
'ElGrid          El Grid que se está utilizando
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


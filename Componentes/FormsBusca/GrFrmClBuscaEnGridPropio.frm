VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form GrFrmBuscaEnGridPropio 
   Caption         =   "Realice su Elección"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   DrawStyle       =   1  'Dash
   Icon            =   "GrFrmClBuscaEnGridPropio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar TbrOpciones 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   582
      ButtonWidth     =   2275
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImlImagenes"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Avanzadas"
            Key             =   "Avanzadas"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir   "
            Key             =   "Imprimir"
            ImageIndex      =   3
            Object.Width           =   1800
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            Key             =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Elegir"
            Key             =   "Elegir"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TbrAvanzadas 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   5040
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
   Begin TrueOleDBGrid60.TDBGrid DtgElige 
      Align           =   3  'Align Left
      Height          =   4035
      Left            =   0
      OleObjectBlob   =   "GrFrmClBuscaEnGridPropio.frx":27A2
      TabIndex        =   7
      Top             =   330
      Width           =   6015
   End
   Begin VB.PictureBox PicCampos 
      Align           =   2  'Align Bottom
      Height          =   675
      Left            =   0
      Picture         =   "GrFrmClBuscaEnGridPropio.frx":2C11B
      ScaleHeight     =   615
      ScaleWidth      =   5565
      TabIndex        =   6
      Top             =   4365
      Width           =   5625
      Begin MSComctlLib.ImageList ImlImagenesAv 
         Left            =   0
         Top             =   0
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
               Picture         =   "GrFrmClBuscaEnGridPropio.frx":2F3B5
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridPropio.frx":2F50F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridPropio.frx":2F669
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridPropio.frx":2F7C3
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridPropio.frx":2F91D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridPropio.frx":2FA77
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridPropio.frx":2FBD1
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox CmbCompara 
         Height          =   315
         ItemData        =   "GrFrmClBuscaEnGridPropio.frx":2FD2B
         Left            =   2160
         List            =   "GrFrmClBuscaEnGridPropio.frx":2FD41
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtCompara 
         Height          =   324
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox CmbCampo 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin MSComctlLib.ImageList ImlImagenes 
         Left            =   4920
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridPropio.frx":2FD5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridPropio.frx":301AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridPropio.frx":30600
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridPropio.frx":3091A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GrFrmClBuscaEnGridPropio.frx":30C34
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   3120
         TabIndex        =   0
         Top             =   0
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operador:"
         Height          =   195
         Left            =   2160
         TabIndex        =   4
         Top             =   0
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Columna:"
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   0
         Width           =   660
      End
   End
End
Attribute VB_Name = "GrFrmBuscaEnGridPropio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************
'**     PROPOSITO GENERAL: Formulario de Búsqueda     **
'**     AUTOR: Dulfredo Rojas Valencia                **
'**     FECHA CREACION: 01/01/98                      **
'**     ULTIMA REVISION: 01/10/99                     **
'**     CONVERSION A COMPONENTE: 31/05/2000           **
'*******************************************************
Option Explicit
Dim rsElige As ADODB.Recordset  'Contiene los datos
Dim Encontro As Boolean
Dim PQConexion As ADODB.Connection
Dim sqlOriginal As String
Dim NuevoQuery As String
Dim mbCtrlKey As Integer
Dim POcultarPrimero As Boolean
Dim PFiltrosMultiples As Boolean
Dim OrdenarAsc As Boolean
Dim AuxAliasGrid As String
Dim PTamañoCampos As String
Dim ListaCampos() As String
Public CodBuscado  As String
Public CodBuscado1  As String
Public CodBuscado2 As String
Public CodBuscado3 As String

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdAceptar_Click()
On Error Resume Next
  If rsElige.RecordCount > 0 Then
    CodBuscado = rsElige.Fields(0).Value   'DtgElige.Columns(0).Text
    If rsElige.Fields.Count > 1 Then CodBuscado1 = rsElige.Fields(1).Value  'DtgElige.Columns(1).Text
    If rsElige.Fields.Count > 2 Then CodBuscado2 = rsElige.Fields(2).Value  'DtgElige.Columns(2).Text
    If rsElige.Fields.Count > 3 Then CodBuscado3 = rsElige.Fields(3).Value  'DtgElige.Columns(3).Text
  End If
  Unload Me
End Sub

Private Sub CmdActualizar_Click()
  On Error GoTo RefErr
  Screen.MousePointer = vbHourglass
  If PFiltrosMultiples Then
    NuevoQuery = sqlOriginal
    rsElige.Close
    rsElige.Open NuevoQuery, PQConexion, adOpenStatic
    Set DtgElige.DataSource = rsElige
  Else
    rsElige.Filter = adFilterNone
    rsElige.Sort = ""
  End If
  If POcultarPrimero Then DtgElige.Columns(0).Visible = False
  NombresColGrid AuxAliasGrid, False
  If PTamañoCampos <> "" Then TamañosColGrid PTamañoCampos
  Screen.MousePointer = vbDefault
  Exit Sub
RefErr:
    MsgBox "Error:" & Err & " " & Err.Description
End Sub

Private Sub OpcionFiltroSimple()
On Error GoTo Que_Error
'**********************
'OJO: Para SQL Server el comodin es %
'OJO: Para Access el comodin es *
'**********************
  If (CmbCampo.Text = "") Or (TxtCompara.Text = "") Then
    MsgBox "Debe elegir el campo que quiere filtrar y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
  Else
    Screen.MousePointer = vbHourglass
    If CmbCompara.Text = "" Then CmbCompara.Text = "="
    'Si es Cadena
    If EsTipoCadena(rsElige.Fields(rsElige.Fields(CmbCampo.ListIndex).Name).Type) Then
      If CmbCompara.Text = "Como" Then
        rsElige.Filter = NombreCampo(CmbCampo.Text) & " LIKE '%" & TxtCompara & "%'"
      Else
        rsElige.Filter = NombreCampo(CmbCampo.Text) & " " & CmbCompara.Text & " '" & TxtCompara & "'"
      End If
    Else 'Si es Número
      If EsTipoNumerico(rsElige.Fields(rsElige.Fields(CmbCampo.ListIndex).Name).Type) Then
        If CmbCompara.Text = "Como" Then CmbCompara.Text = "="
        rsElige.Filter = NombreCampo(CmbCampo.Text) & " " & CmbCompara.Text & " " & TxtCompara
      Else 'Si es fecha
        If EsTipoFechaHora(rsElige.Fields(rsElige.Fields(CmbCampo.ListIndex).Name).Type) Then
          If CmbCompara.Text = "Como" Then CmbCompara.Text = "="
          rsElige.Filter = NombreCampo(CmbCampo.Text) & " " & CmbCompara.Text & " '" & TxtCompara & "'"
        Else   'Otro Tipo
          rsElige.Filter = NombreCampo(CmbCampo.Text) & " = " & TxtCompara
        End If
      End If
    End If
    Screen.MousePointer = vbDefault
  End If
  Exit Sub
Que_Error:
  Screen.MousePointer = vbDefault
  If Err.Number <> 0 Then
    MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error al Filtrar"
  End If
End Sub

Private Sub OpcionFiltrosMultiples()
Dim NuevoFiltro As String
On Error GoTo Que_Error
'**********************
'OJO: Para SQL Server el comodin es %
'OJO: Para Access el comodin es *
'**********************
  If (CmbCampo.Text = "") Then
    MsgBox "Debe elegir el campo que quiere filtrar y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
    Exit Sub
  End If
  If (ListaCampos(CmbCampo.ListIndex) = "") Then 'Or (TxtCompara.Text = "") Then
    MsgBox "Debe elegir el campo que quiere filtrar y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
  Else
    Screen.MousePointer = vbHourglass
    If CmbCompara.Text = "" Then CmbCompara.Text = "="
    If TxtCompara.Text = "" Then CmbCompara.Text = "="
    'Si es Cadena
    If EsTipoCadena(rsElige.Fields(rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name).Type) Then
      If CmbCompara.Text = "Como" Then
        NuevoFiltro = "(" & NombreCampoQuery(ListaCampos(CmbCampo.ListIndex)) & " LIKE '%" & TxtCompara & "%')"
      Else
        NuevoFiltro = "(" & NombreCampoQuery(ListaCampos(CmbCampo.ListIndex)) & " " & CmbCompara.Text & " '" & TxtCompara & "')"
      End If
    Else 'Si es Número
      If EsTipoNumerico(rsElige.Fields(rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name).Type) Then
        If TxtCompara.Text <> "" Then If Not IsNumeric(TxtCompara.Text) Then MsgBox "Debe ingresar una Valor válido (Numérico)...", vbExclamation + vbOKOnly, "Atención": Screen.MousePointer = vbDefault: Exit Sub
        If CmbCompara.Text = "Como" Then CmbCompara.Text = "="
        NuevoFiltro = "(" & NombreCampoQuery(ListaCampos(CmbCampo.ListIndex)) & " " & CmbCompara.Text & " " & TxtCompara & ")"
      Else 'Si es fecha
        If EsTipoFechaHora(rsElige.Fields(rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name).Type) Then
          If TxtCompara.Text <> "" Then If Not IsDate(TxtCompara.Text) Then MsgBox "Debe ingresar una Valor válido (Fecha)...", vbExclamation + vbOKOnly, "Atención": Screen.MousePointer = vbDefault: Exit Sub
          If CmbCompara.Text = "Como" Then CmbCompara.Text = "="
          NuevoFiltro = "(" & NombreCampoQuery(ListaCampos(CmbCampo.ListIndex)) & " " & CmbCompara.Text & " '" & TxtCompara & "')"
        Else   'Otro Tipo
          NuevoFiltro = "(" & NombreCampoQuery(ListaCampos(CmbCampo.ListIndex)) & " = " & TxtCompara & ")"
        End If
      End If
    End If
    If (TxtCompara.Text = "") Then NuevoFiltro = "(" & NuevoFiltro & "or(" & NombreCampoQuery(ListaCampos(CmbCampo.ListIndex)) & " is Null))"
    'Ve si tiene WHERE
    If InStr(1, UCase(NuevoQuery), "WHERE") <> 0 Then
      NuevoQuery = NuevoQuery & " and " & NuevoFiltro
    Else
      NuevoQuery = NuevoQuery & " WHERE " & NuevoFiltro
    End If
    rsElige.Close
    rsElige.Open NuevoQuery, PQConexion, adOpenStatic
    Set DtgElige.DataSource = rsElige
    If POcultarPrimero Then DtgElige.Columns(0).Visible = False
    NombresColGrid AuxAliasGrid, False
    If PTamañoCampos <> "" Then TamañosColGrid PTamañoCampos
    Screen.MousePointer = vbDefault
  End If
  Exit Sub
Que_Error:
  Screen.MousePointer = vbDefault
  If Err.Number <> 0 Then
    MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error al Filtrar"
  End If
End Sub

Private Sub OpcionOrdenar(EsDeBarra As Boolean)
Dim AuxCampo As String
On Error GoTo Que_Error
  If (CmbCampo.Text = "") Then
    MsgBox "Debe elegir el campo por el que quiere ordenar", vbInformation + vbOKOnly, "Error de procedimiento"
  Else
    Screen.MousePointer = vbHourglass
    'check for the use of the ctrl key for descending sort
'    AuxCampo = NombreCampo(CmbCampo.Text)
    AuxCampo = ListaCampos(CmbCampo.ListIndex)
    If EsDeBarra Then
      If OrdenarAsc Then
        rsElige.Sort = AuxCampo & " ASC"
      Else
        rsElige.Sort = AuxCampo & " DESC"
      End If
    Else
      If mbCtrlKey Then
        rsElige.Sort = AuxCampo & " DESC"
        mbCtrlKey = 0
      Else
        rsElige.Sort = AuxCampo & " ASC"
      End If
    End If
    Set DtgElige.DataSource = rsElige
    If POcultarPrimero Then DtgElige.Columns(0).Visible = False
    NombresColGrid AuxAliasGrid, False
    If PTamañoCampos <> "" Then TamañosColGrid PTamañoCampos
    Screen.MousePointer = vbDefault
  End If
  Exit Sub
Que_Error:
  Screen.MousePointer = vbDefault
  If Err.Number <> 0 Then
    MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error al Ordenar"
  End If
End Sub

Private Sub CmdPrimero_Click()
On Error GoTo QueError
Dim Marca As Variant
  Screen.MousePointer = vbHourglass
  Marca = rsElige.Bookmark
  rsElige.MoveFirst
  If EsTipoCadena(rsElige.Fields(NombreCampo(ListaCampos(CmbCampo.ListIndex))).Type) Then
    Select Case CmbCompara.Text
      Case "=", ">", ">=", "<", "<="
        rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " " & CmbCompara.Text & " '" & TxtCompara.Text & "'", 0, adSearchForward
      Case "Como"
        rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " LIKE '%" & TxtCompara.Text & "%'", 0, adSearchForward
    End Select
  Else
    If EsTipoNumerico(rsElige.Fields(rsElige.Fields(CmbCampo.ListIndex).Name).Type) Then
      If TxtCompara.Text <> "" Then If Not IsNumeric(TxtCompara.Text) Then MsgBox "Debe ingresar una Valor válido (Numérico)...", vbExclamation + vbOKOnly, "Atención": Screen.MousePointer = vbDefault: Exit Sub
      Select Case CmbCompara.Text
        Case "=", ">", ">=", "<", "<="
          rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " " & CmbCompara.Text & " " & TxtCompara.Text, 0, adSearchForward
        Case "Como"
          rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " LIKE " & TxtCompara.Text, 0, adSearchForward
      End Select
    Else
      If CmbCompara.Text = "Como" Then CmbCompara.Text = "="
      rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " " & CmbCompara.Text & " " & TxtCompara.Text, 0, adSearchForward
    End If
  End If
  If rsElige.EOF Then Encontro = False Else Encontro = True
  Screen.MousePointer = vbDefault
  If Not Encontro Then
    MsgBox "No existe ninguna coincidencia", vbInformation + vbOKOnly, "No encontrado"
    rsElige.Bookmark = Marca
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

Private Sub CmdAnterior_Click()
On Error GoTo QueError
Dim Marca As Variant
  Screen.MousePointer = vbHourglass
  Marca = rsElige.Bookmark
  rsElige.MovePrevious
  If rsElige.BOF Then rsElige.MoveFirst
  If EsTipoCadena(rsElige.Fields(NombreCampo(ListaCampos(CmbCampo.ListIndex))).Type) Then
    Select Case CmbCompara.Text
      Case "=", ">", ">=", "<", "<="
          rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " " & CmbCompara.Text & " '" & TxtCompara.Text & "'", 0, adSearchBackward
      Case "Como"
          rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " LIKE '%" & TxtCompara.Text & "%'", 0, adSearchBackward
    End Select
  Else
    If EsTipoNumerico(rsElige.Fields(rsElige.Fields(CmbCampo.ListIndex).Name).Type) Then
      If TxtCompara.Text <> "" Then If Not IsNumeric(TxtCompara.Text) Then MsgBox "Debe ingresar una Valor válido (Numérico)...", vbExclamation + vbOKOnly, "Atención": Screen.MousePointer = vbDefault: Exit Sub
      Select Case CmbCompara.Text
        Case "=", ">", ">=", "<", "<="
            rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " " & CmbCompara.Text & " " & TxtCompara.Text, 0, adSearchBackward
        Case "Como"
            rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " LIKE " & TxtCompara.Text, 0, adSearchBackward
      End Select
    Else
      If CmbCompara.Text = "Como" Then CmbCompara.Text = "="
      rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " " & CmbCompara.Text & " " & TxtCompara.Text, 0, adSearchBackward
    End If
  End If
  If rsElige.BOF Then Encontro = False Else Encontro = True
  Screen.MousePointer = vbDefault
  If Not Encontro Then
    MsgBox "No existe ninguna coincidencia", vbInformation + vbOKOnly, "No encontrado"
    rsElige.Bookmark = Marca
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

Private Sub CmdSiguiente_Click()
On Error GoTo QueError
Dim Marca As Variant
  Screen.MousePointer = vbHourglass
  Marca = rsElige.Bookmark
  rsElige.MoveNext
  If rsElige.EOF Then rsElige.MoveLast
  If EsTipoCadena(rsElige.Fields(NombreCampo(ListaCampos(CmbCampo.ListIndex))).Type) Then
    Select Case CmbCompara.Text
      Case "=", ">", ">=", "<", "<="
          rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " " & CmbCompara.Text & " '" & TxtCompara.Text & "'", 0, adSearchForward
      Case "Como"
          rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " LIKE '%" & TxtCompara.Text & "%'", 0, adSearchForward
    End Select
  Else
    If EsTipoNumerico(rsElige.Fields(rsElige.Fields(CmbCampo.ListIndex).Name).Type) Then
      If TxtCompara.Text <> "" Then If Not IsNumeric(TxtCompara.Text) Then MsgBox "Debe ingresar una Valor válido (Numérico)...", vbExclamation + vbOKOnly, "Atención": Screen.MousePointer = vbDefault: Exit Sub
      Select Case CmbCompara.Text
        Case "=", ">", ">=", "<", "<="
          rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " " & CmbCompara.Text & " " & TxtCompara.Text, 0, adSearchForward
        Case "Como"
            rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " LIKE " & TxtCompara.Text, 0, adSearchForward
      End Select
    Else
      If CmbCompara.Text = "Como" Then CmbCompara.Text = "="
      rsElige.Find rsElige.Fields(ListaCampos(CmbCampo.ListIndex)).Name & " " & CmbCompara.Text & " " & TxtCompara.Text, 0, adSearchForward
    End If
  End If
  If rsElige.EOF Then Encontro = False Else Encontro = True
  Screen.MousePointer = vbDefault
  If Not Encontro Then
    MsgBox "No existe ninguna coincidencia", vbInformation + vbOKOnly, "No encontrado"
    rsElige.Bookmark = Marca
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
Private Sub DtgElige_DblClick()
  CmdAceptar_Click
End Sub

Private Sub DtgElige_GroupColMove(ByVal Position As Integer, ByVal ColIndex As Integer, Cancel As Integer)
    Dim strSort As String
    Dim Col As TrueOleDBGrid60.Column
' Loop through GroupColumns collection and construct
' the sort string for the Sort property of the Recordset
    For Each Col In DtgElige.GroupColumns
        If strSort <> vbNullString Then
            strSort = strSort & ", "
        End If
        strSort = strSort & "[" & Col.DataField & "]"
    Next Col
    DtgElige.HoldFields
    rsElige.Sort = strSort
End Sub

Private Sub DtgElige_HeadClick(ByVal ColIndex As Integer)
Dim AuxCampo As String
    Screen.MousePointer = vbHourglass
    'check for the use of the ctrl key for descending sort
    AuxCampo = rsElige.Fields(ColIndex).Name
    If mbCtrlKey Then
      rsElige.Sort = AuxCampo & " DESC"
      mbCtrlKey = 0
    Else
      rsElige.Sort = AuxCampo & " ASC"
    End If
    Set DtgElige.DataSource = rsElige
    If POcultarPrimero Then DtgElige.Columns(0).Visible = False
    NombresColGrid AuxAliasGrid, False
    If PTamañoCampos <> "" Then TamañosColGrid PTamañoCampos
    Screen.MousePointer = vbDefault
End Sub

Private Sub DtgElige_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mbCtrlKey = Shift
End Sub

Private Sub Form_Load()
  OrdenarAsc = True
End Sub

Private Sub TxtCompara_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
   CmdPrimero_Click
 End If
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  DtgElige.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  rsElige.Close
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
    OpcionOrdenar True
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
      CmdActualizar_Click
    Case "Filtrar"
      If PFiltrosMultiples Then
        OpcionFiltrosMultiples
      Else
        OpcionFiltroSimple
      End If
    Case "Ordenar"
      OpcionOrdenar True
    Case "Buscar"
      If (CmbCampo.Text = "") Or (CmbCompara.Text = "") Or (TxtCompara.Text = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Debe elegir la columna por la que quiere buscar, el operador y escribir un valor", vbInformation + vbOKOnly, "Error de Procedimiento"
        CmbCampo.SetFocus
        Exit Sub
      End If
      Select Case Button.Caption
        Case "&Primero"
          CmdPrimero_Click
        Case "&Anterior"
          CmdAnterior_Click
        Case "&Siguiente"
          CmdSiguiente_Click
      End Select
  End Select
End Sub

Private Sub TbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "Avanzadas"
      If TbrOpciones.Buttons(1).Value Then
        PicCampos.Visible = True
        TbrAvanzadas.Visible = True
        TbrOpciones.Buttons(1).Image = ImlImagenes.ListImages(2).Index
      Else
        PicCampos.Visible = False
        TbrAvanzadas.Visible = False
        TbrOpciones.Buttons(1).Image = ImlImagenes.ListImages(1).Index
      End If
    Case "Imprimir"
      DtgElige.PrintInfo.PageHeader = Me.Caption
      DtgElige.PrintInfo.SettingsOrientation = 2
      DtgElige.PrintInfo.PrintPreview
    Case "Cancelar"
      cmdCancelar_Click
    Case "Elegir"
      CmdAceptar_Click
  End Select
End Sub

'********************** Procedimientos ********************
Public Sub Elige(QConexion As Object, sqlEnviado As String, Titulo As String, _
                 Optional CDefecto As String, Optional OcultarPrimero As Boolean = False, _
                 Optional TamañoCampos As String, Optional FiltrosMultiples As Boolean = False, _
                 Optional AliasGrid As String)
'sqlEnviado,     Query para buscar
'Título,         Titulo de Venta, "" para titulo por defecto
'CDefecto,       Campo de busqueda que se activa por defecto
'OcultarPrimero, Si quiere ocultar la primera columna
'TamañoCampos,   Ejem: 1-1.5-2; que representa 3 columnas donde:
'                la primera debe estar en su tamaño por defecto,
'                la segunda 0.5 mas de su tamaño y
'                la tercera 2 veces su tamaño por defecto
'AliasGrid,      NombreCampo1-NombreCampo2,...
'Dim i As Byte
On Error GoTo LoadErr
  Screen.MousePointer = vbHourglass
  PFiltrosMultiples = FiltrosMultiples
  AuxAliasGrid = AliasGrid
  PTamañoCampos = TamañoCampos
  POcultarPrimero = OcultarPrimero
  PicCampos.Visible = False
  TbrAvanzadas.Visible = False
  CodBuscado = "":    CodBuscado1 = "":    CodBuscado2 = ""
  mbCtrlKey = 0
  sqlOriginal = sqlEnviado
  NuevoQuery = sqlEnviado
  Set PQConexion = New ADODB.Connection
  Set PQConexion = QConexion
  Set rsElige = New ADODB.Recordset
  rsElige.CursorLocation = adUseClient
  rsElige.Open sqlOriginal, QConexion, adOpenStatic
  Set DtgElige.DataSource = rsElige
  If Titulo <> "" Then Me.Caption = Titulo
  On Error Resume Next
  CmbCompara.Text = "Como"
  If CDefecto <> "" Then CmbCampo.Text = CDefecto
  'Ocultar primera columna de codigo
  If OcultarPrimero Then DtgElige.Columns(0).Visible = False
  'Nombres de Columnas
  NombresColGrid AuxAliasGrid, True
  'Tamaños de columnas
  If PTamañoCampos <> "" Then TamañosColGrid PTamañoCampos
  Screen.MousePointer = vbDefault
  Me.Show vbModal ' Muestra el formulario
  Exit Sub
LoadErr:
  Screen.MousePointer = vbDefault
  MsgBox "Error al mostrar Proyecto:" & Err & " " & Err.Description
  Select Case MsgBox("¿ Anula, Reintenta o Ignora ?", vbCritical + vbAbortRetryIgnore, "Error!")
    Case vbRetry
      Resume
    Case vbIgnore
      Resume Next
  End Select
  Unload Me
End Sub

Private Function NombreCampo(CampoBusca As String) As String
  If InStr(1, CampoBusca, " ", vbTextCompare) Then
    NombreCampo = "[" & CampoBusca & "]"
'    NombreCampo = rsElige.Fields(DtgElige.Col).Name
  Else
    NombreCampo = CampoBusca
  End If
'Dim AuxCad1 As String
'Dim QuerySelect As String, QueryBusca As String
'Dim i As Integer, j As Integer
'Dim Encontro As Boolean, EncontroVacio As Boolean
'  '#1. Si es *, entonces el nombre del campo es tal cual
'  If InStr(1, sqlOriginal, "*", 1) <> 0 Then
'    NombreCampo = CampoBusca
'    Exit Function
'  End If
'  '#2. Recorta el Query hasta antes del FROM
'  Encontro = False
'  QuerySelect = Mid(sqlOriginal, 1, InStr(1, sqlOriginal, " From ", vbTextCompare) - 1)
'  '#3. Busca el campo de atras hacia adelante
'  i = Len(QuerySelect)
'  While (i >= 1) And Not Encontro
'    QueryBusca = Mid(QuerySelect, i, Len(QuerySelect) - 1)
'    If InStr(1, QueryBusca, CampoBusca, 1) <> 0 Then
'      Encontro = True
'    Else
'      i = i - 1
'    End If
'  Wend
'  '#4. Pone en QuerySelect, desde SELECT hasta el campo encontrado
'  QuerySelect = Mid(QuerySelect, 1, i + Len(CampoBusca) - 1)
'  '#5. Verifica que hay antes del campo requerido ('.', ',' o ' ') y en base a eso
'  '    obtiene el nombre real
'  AuxCad1 = Mid(QuerySelect, i - 1, 1)
'  Select Case Mid(QuerySelect, i - 1, 1)
'    Case "."
'      '** Si es Punto, busca el prefijo para ponerlo en NombreCampo
'      EncontroVacio = False
'      j = i - 2
'      While j >= 1 And Not EncontroVacio
'        If (Mid(QuerySelect, j, 1) = " ") Or (Mid(QuerySelect, j, 1) = ",") Then
'          EncontroVacio = True
'        Else
'          j = j - 1
'        End If
'      Wend
'      NombreCampo = Mid(QuerySelect, j + 1, i - j + Len(CampoBusca) + 1)
'    Case ","
'      '** Si es coma el nombre del campo es tal cual
'      NombreCampo = CampoBusca
'    Case " "
'      '** Si es espacio, llegar a primer caracter distinto de vacio
'      j = i - 2
'      Encontro = False
'      While Not Encontro
'        If Mid(QuerySelect, j, 1) <> " " Then
'          Encontro = True
'        Else
'          j = j - 1
'        End If
'      Wend
'      '** Si el caracter encontrado es ',', entonces el nombre es tal cual
'      If Mid(QuerySelect, j, 1) = "," Then
'        NombreCampo = CampoBusca
'      Else
'        '** Si encuentra una 'S' de 'AS', se trata de un alias
'        If UCase(Mid(QuerySelect, j, 1)) = "S" Then
'          j = j - 2
'          i = j
'          'llegar a primer caracter distinto de vacio
'          Encontro = False
'          While Not Encontro
'            If Mid(QuerySelect, j, 1) <> " " Then
'              Encontro = True
'            Else
'              j = j - 1
'            End If
'          Wend
'          'llegar a primer caracter vacio o ,
'          Encontro = False
'          While Not Encontro
'            If (Mid(QuerySelect, j, 1) = " ") Or (Mid(QuerySelect, j, 1) = ",") Then
'              Encontro = True
'            Else
'              j = j - 1
'            End If
'          Wend
'          NombreCampo = Trim(Mid(QuerySelect, j + 1, i - j))
'        Else
'          '** Si es otra cosa, entonces el nombre es tal cual
'          NombreCampo = CampoBusca
'        End If
'      End If
'  End Select
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

Private Sub TamañosColGrid(CadTamañoCampos As String)
Dim AuxTamaños As String
Dim AuxTamCampo As String
Dim AuxPos As Integer
Dim i As Integer
    AuxTamaños = CadTamañoCampos
    For i = 0 To DtgElige.Columns.Count - 1
      AuxPos = InStr(1, AuxTamaños, "-", vbTextCompare)
      If (AuxPos <> 0) Then
        AuxTamCampo = Mid(AuxTamaños, 1, AuxPos - 1)
        AuxTamaños = Mid(AuxTamaños, AuxPos + 1, Len(AuxTamaños) - AuxPos)
        If IsNumeric(AuxTamCampo) Then
          If DtgElige.Columns(i).Visible Then
            DtgElige.Columns(i).Width = DtgElige.Columns(i).Width * AuxTamCampo
          End If
        End If
      Else
        If Len(AuxTamaños) > 0 Then
            If IsNumeric(AuxTamaños) Then
              If DtgElige.Columns(i).Visible Then
                DtgElige.Columns(i).Width = DtgElige.Columns(i).Width * AuxTamaños
              End If
            End If
            AuxTamaños = ""
        End If
      End If
    Next i
End Sub

Public Sub NombresColGrid(PAliasGrid As String, PrimeraVez As Boolean)
Dim AuxTamaños As String
Dim AuxTamCampo As String
Dim AuxPos As Integer
Dim i As Integer
Dim Campos As ADODB.Field
  If PrimeraVez Then ReDim ListaCampos(DtgElige.Columns.Count)
  If PAliasGrid <> "" Then
    AuxTamaños = PAliasGrid
    For i = 0 To DtgElige.Columns.Count - 1
      AuxPos = InStr(1, AuxTamaños, "-", vbTextCompare)
      If (AuxPos <> 0) Then
        AuxTamCampo = Mid(AuxTamaños, 1, AuxPos - 1)
        AuxTamaños = Mid(AuxTamaños, AuxPos + 1, Len(AuxTamaños) - AuxPos)
        DtgElige.Columns(i).Caption = AuxTamCampo
        If PrimeraVez Then
          If (i = 0 And POcultarPrimero) Then
            CmbCampo.AddItem ""
          Else
            CmbCampo.AddItem AuxTamCampo
          End If
          ListaCampos(i) = DtgElige.Columns(i).DataField
        End If
      Else
        If Len(AuxTamaños) > 0 Then
          DtgElige.Columns(i).Caption = AuxTamaños
          AuxTamaños = ""
          If PrimeraVez Then
          If (i = 0 And POcultarPrimero) Then
            CmbCampo.AddItem ""
          Else
            CmbCampo.AddItem DtgElige.Columns(i).Caption
          End If
          ListaCampos(i) = DtgElige.Columns(i).DataField
          End If
        End If
      End If
    Next i
  Else
    If PrimeraVez Then
      i = 0
      ReDim ListaCampos(rsElige.Fields.Count)
      For Each Campos In rsElige.Fields
        If (i = 0 And POcultarPrimero) Then
          CmbCampo.AddItem ""
        Else
          CmbCampo.AddItem (Campos.Name)
        End If
        ListaCampos(i) = Campos.Name
        i = i + 1
      Next Campos
    End If
  End If
End Sub

Private Function NombreCampoQuery(NombreDelCampo As String) As String
Dim AuxQuery As String
Dim PosCampo As Integer
Dim i As Integer
Dim Encontro As Boolean
  AuxQuery = Mid(sqlOriginal, 1, InStr(1, sqlOriginal, "FROM", vbTextCompare) - 1)
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


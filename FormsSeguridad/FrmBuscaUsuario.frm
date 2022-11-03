VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmBuscaUsuario 
   Caption         =   "Busca Usuario"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   Icon            =   "FrmBuscaUsuario.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmBuscaUsuario.frx":0442
   ScaleHeight     =   4845
   ScaleWidth      =   8715
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtnAceptar 
      BackColor       =   &H8000000A&
      Caption         =   "Elegir"
      Height          =   720
      Left            =   2640
      Picture         =   "FrmBuscaUsuario.frx":6C474
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton BtnCancelar 
      BackColor       =   &H8000000A&
      Caption         =   "Cancelar"
      Height          =   720
      Left            =   4320
      Picture         =   "FrmBuscaUsuario.frx":6C77E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   855
   End
   Begin VB.Frame FraCriterio 
      Height          =   915
      Left            =   30
      TabIndex        =   1
      Top             =   -30
      Width           =   8655
      Begin VB.ComboBox cmbActivo 
         Height          =   315
         ItemData        =   "FrmBuscaUsuario.frx":6C988
         Left            =   6120
         List            =   "FrmBuscaUsuario.frx":6C995
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   1410
      End
      Begin VB.ComboBox cmbNivelAcceso 
         Height          =   315
         Left            =   3105
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   2490
      End
      Begin VB.ComboBox cmbUsuarios 
         Height          =   315
         ItemData        =   "FrmBuscaUsuario.frx":6C9A8
         Left            =   300
         List            =   "FrmBuscaUsuario.frx":6C9B5
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2235
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Activo"
         Height          =   195
         Left            =   6135
         TabIndex        =   6
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nivel de Acceso"
         Height          =   210
         Left            =   3105
         TabIndex        =   4
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuarios"
         Height          =   210
         Left            =   300
         TabIndex        =   3
         Top             =   210
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid dtgUsuarios 
      Height          =   3015
      Left            =   30
      TabIndex        =   0
      Top             =   960
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   13564411
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "IdFuncionario"
         Caption         =   "Id Funcionario"
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
         DataField       =   "usr_usuario"
         Caption         =   "Usuario"
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
         DataField       =   "Paterno"
         Caption         =   "Primer Apellido"
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
         DataField       =   "Materno"
         Caption         =   "Segundo Apellido"
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
         DataField       =   "Nombres"
         Caption         =   "Nombres"
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
         DataField       =   "idNivelAcceso"
         Caption         =   "Nivel Acceso"
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
         DataField       =   "Usr_Activo"
         Caption         =   "Activo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Si"
            FalseValue      =   "No"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   659.906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmBuscaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUsuarios As New ADODB.Recordset
Dim rsNivelAcceso As New ADODB.Recordset
Dim vUsuarios As String
Dim vNivelAcceso As Byte
Dim vActivo As String

Private Sub cmbActivo_Click()
    vActivo = cmbActivo
    FiltrarUsuarios vUsuarios, vNivelAcceso, vActivo
End Sub

Private Sub cmbNivelAcceso_Click()
    vNivelAcceso = CByte(Mid(cmbNivelAcceso, 1, 1))
    FiltrarUsuarios vUsuarios, vNivelAcceso, vActivo
End Sub

Private Sub cmbUsuarios_Click()
    vUsuarios = Mid(cmbUsuarios, 1, 1)
    FiltrarUsuarios vUsuarios, vNivelAcceso, vActivo
End Sub

Private Sub BtnCancelar_Click()
    GlElegido = ""
    Unload Me
End Sub

Private Sub BtnAceptar_Click()
On Error GoTo Error
    GlElegido = rsUsuarios!usr_usuario
    Unload Me
    Exit Sub
Error:
    GlElegido = ""
    Unload Me
End Sub

Private Sub dtgUsuarios_DblClick()
    BtnAceptar_Click
End Sub

Private Sub Form_Load()
    vUsuarios = "T"
    vNivelAcceso = 0
    vActivo = "Todos"
    
    rsNivelAcceso.Open "Select Distinct IdNivelAcceso, DesNivelAcceso Fromgc_nivelacceso Order by IdNivelAcceso", db, adOpenStatic
    If rsNivelAcceso.RecordCount > 0 Then
        cmbNivelAcceso.AddItem "0 - Todos"
        While Not rsNivelAcceso.EOF
            cmbNivelAcceso.AddItem rsNivelAcceso!IdNivelAcceso & " - " & rsNivelAcceso!DesNivelAcceso
            rsNivelAcceso.MoveNext
        Wend
    Else
        cmbNivelAcceso.AddItem "NO EXISTEN NIVELES"
    End If
    rsNivelAcceso.Close
    
    cmbUsuarios.ListIndex = 0
    cmbNivelAcceso.ListIndex = 0
    cmbActivo.ListIndex = 0
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsUsuarios.Close
End Sub

Private Sub FiltrarUsuarios(pUsuarios As String, pNivelAcceso As Byte, pActivo As String)
If rsUsuarios.State = 1 Then rsUsuarios.Close
If pUsuarios = "T" And pNivelAcceso = 0 And pActivo = "Todos" Then
    rsUsuarios.Open "Select * From gc_usuarios Order by Paterno", db, adOpenStatic
End If

If pUsuarios <> "T" And pNivelAcceso = 0 And pActivo = "Todos" Then
    If pUsuarios = "O" Then
        rsUsuarios.Open "Select * From gc_usuarios Where IdFuncionario = 0 Order by Paterno", db, adOpenStatic
    Else
        rsUsuarios.Open "Select * From gc_usuarios Where IdFuncionario <> 0 Order by Paterno", db, adOpenStatic
    End If
End If

If pUsuarios = "T" And pNivelAcceso <> 0 And pActivo = "Todos" Then
    rsUsuarios.Open "Select * From gc_usuarios Where idNivelAcceso=" & pNivelAcceso & " Order by Paterno", db, adOpenStatic
End If

If pUsuarios = "T" And pNivelAcceso = 0 And pActivo <> "Todos" Then
    If pActivo = "Si" Then
        rsUsuarios.Open "Select * From gc_usuarios Where usr_Activo=" & -1 & " Order by Paterno", db, adOpenStatic
    Else
        rsUsuarios.Open "Select * From gc_usuarios Where usr_Activo=" & 0 & " Order by Paterno", db, adOpenStatic
    End If
End If

If pUsuarios <> "T" And pNivelAcceso <> 0 And pActivo = "Todos" Then
    If pUsuarios = "O" Then
        rsUsuarios.Open "Select * From gc_usuarios Where IdFuncionario=0 and IdNivelAcceso=" & pNivelAcceso & " Order by Paterno", db, adOpenStatic
    Else
        rsUsuarios.Open "Select * From gc_usuarios Where IdFuncionario<>0 and IdNivelAcceso=" & pNivelAcceso & " Order by Paterno", db, adOpenStatic
    End If
End If

If pUsuarios <> "T" And pNivelAcceso = 0 And pActivo <> "Todos" Then
    If pUsuarios = "O" Then
        rsUsuarios.Open "Select * From gc_usuarios Where IdFuncionario=0 and usr_Activo=" & IIf(pActivo = "Si", -1, 0) & " Order by Paterno", db, adOpenStatic
    Else
        rsUsuarios.Open "Select * From gc_usuarios Where IdFuncionario<>0 and usr_Activo=" & IIf(pActivo = "Si", -1, 0) & " Order by Paterno", db, adOpenStatic
    End If
End If

If pUsuarios = "T" And pNivelAcceso <> 0 And pActivo <> "Todos" Then
    rsUsuarios.Open "Select * From gc_usuarios Where IdNivelAcceso=" & pNivelAcceso & " and usr_Activo=" & IIf(pActivo = "Si", -1, 0) & " Order by Paterno", db, adOpenStatic
End If

If pUsuarios <> "T" And pNivelAcceso <> 0 And pActivo <> "Todos" Then
    If pUsuarios = "O" Then
        rsUsuarios.Open "Select * From gc_usuarios Where IdFuncionario=0 and IdNivelAcceso=" & pNivelAcceso & " and usr_Activo=" & IIf(pActivo = "Si", -1, 0) & " Order by Paterno", db, adOpenStatic
    Else
        rsUsuarios.Open "Select * From gc_usuarios Where IdFuncionario<>0 and IdNivelAcceso=" & pNivelAcceso & " and usr_Activo=" & IIf(pActivo = "Si", -1, 0) & " Order by Paterno", db, adOpenStatic
    End If
End If

Set dtgUsuarios.DataSource = rsUsuarios
End Sub

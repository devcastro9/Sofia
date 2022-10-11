VERSION 5.00
Begin VB.Form FrmSisUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuario del Sistema"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "FrmSisUsuarios.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   1890
      Width           =   1335
   End
   Begin VB.Frame FraDatos 
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox cmbNivelAcceso 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2265
         Width           =   3180
      End
      Begin VB.TextBox txtUsuario 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1800
         Width           =   1740
      End
      Begin VB.CommandButton cmdBuscarFuncionario 
         Caption         =   "..."
         Height          =   255
         Left            =   2280
         Picture         =   "FrmSisUsuarios.frx":0442
         TabIndex        =   1
         ToolTipText     =   "Buscar Funcionarios"
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkUsr_Activo 
         Caption         =   "Usuario Activo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3240
         TabIndex        =   4
         Top             =   1815
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Apellido paterno:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Apellido materno:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombres:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Creación:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   1170
      End
      Begin VB.Label lblFechaCrea 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1410
         TabIndex        =   6
         Top             =   2760
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nivel de Acceso:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2340
         Width           =   1215
      End
      Begin VB.Label lblIdFuncionario 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Id. Funcionario:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblPaterno 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label lblMaterno 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label lblNombres 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   600
         Width           =   3135
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   4560
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   120
         X2              =   4560
         Y1              =   1700
         Y2              =   1700
      End
   End
   Begin VB.CommandButton cmdBuscaUsuario 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   2320
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   1000
      Width           =   1335
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   560
      Width           =   1335
   End
End
Attribute VB_Name = "FrmSisUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUsuarios As New ADODB.Recordset
Dim rsFuncionarios As New ADODB.Recordset
Dim rsNivelAcceso As New ADODB.Recordset
Dim Editando As Boolean
Dim i As Byte
Dim vecNivelAcceso(10) As String

Private Sub cmdBuscarFuncionario_Click()
    FrmBuscaFuncionario.Show vbModal
    If GlElegido <> "" Then
        If rsFuncionarios.State = 1 Then rsFuncionarios.Close
        rsFuncionarios.Open "Select * From rc_Personal Where IdFuncionario=" & GlElegido & " Order by Paterno", db, adOpenStatic
        lblIdFuncionario = rsFuncionarios!IdFuncionario
        lblPaterno = rsFuncionarios!Paterno
        lblMaterno = rsFuncionarios!Materno
        lblNombres = rsFuncionarios!Nombres
        'Txtusuario = Mid(rsFuncionarios!Nombres, 1, 1) & Mid(rsFuncionarios!Paterno, 1, 1) & Mid(rsFuncionarios!Materno, 1, 1) & Mid(CStr(1000 + rsFuncionarios!IdFuncionario), 2, 3)
        Txtusuario = Mid(rsFuncionarios!Nombres, 1, 1) & "_" & (rsFuncionarios!Paterno)
    End If
End Sub

Private Sub cmdBuscaUsuario_Click()
Dim Encontrado As Boolean
    FrmBuscaUsuario.Show vbModal
    If GlElegido <> "" Then
        Encontrado = False
        rsUsuarios.MoveFirst
        While Not rsUsuarios.EOF And Not Encontrado
            If rsUsuarios!IdFuncionario = GlElegido Then
                Encontrado = True
            Else
                rsUsuarios.MoveNext
            End If
        Wend
        If Encontrado Then RecuperaUsuario
    End If
End Sub

Private Sub CmdCancelar_Click()
    cmdBuscarFuncionario.Enabled = False
    If rsUsuarios.BOF Or rsUsuarios.EOF Then
        rsUsuarios.MoveFirst
    End If
    RecuperaUsuario
    BotonesNavegar
End Sub

Private Sub Cmdeditar_Click()
    Editando = True
    FraDatos.Enabled = True
    BotonesConfirma
End Sub

Private Sub cmdEliminar_Click()
If Txtusuario = "ADMIN" Then
  MsgBox "El usuario Administrador no puede eliminarse.", vbExclamation + vbOKOnly, "Atencion"
  Exit Sub
End If
If MsgBox("Esta seguro de eliminar al usuario visualizado?", vbExclamation + vbYesNo, "Atención") = vbYes Then
    rsUsuarios.Delete
    If rsUsuarios.RecordCount > 0 Then
        rsUsuarios.MoveNext
        If rsUsuarios.EOF Then rsUsuarios.MoveLast
        RecuperaUsuario
    Else
        VaciaCampos
        BotonesInicio
    End If
End If
End Sub

Private Sub CmdGrabar_Click()
On Error GoTo QueError
    If ValidaCampos Then
        cmdBuscarFuncionario.Enabled = False
        If Not Editando Then
            rsUsuarios.AddNew
        End If
        rsUsuarios!IdFuncionario = lblIdFuncionario
        rsUsuarios!usr_usuario = Txtusuario
        rsUsuarios!usr_clave = Encriptar("x") 'Cuando se crea un nuevo usuario su clave sera "x"
        rsUsuarios!NivelAcceso = CInt(cmbNivelAcceso.ListIndex + 1)
        rsUsuarios!Usr_Activo = CBool(chkUsr_Activo.Value)
        rsUsuarios!FechaCrea = lblFechaCrea
        db.BeginTrans
        rsUsuarios.Update
        db.CommitTrans
        Editando = False
        BotonesNavegar
        MsgBox "Datos del usuario grabado satisfactoriamente", vbInformation + vbOKOnly, "Atención"
    End If
QueError:
    If Err.Number = -2147467259 Then
        db.RollbackTrans
        rsUsuarios.CancelUpdate
        MsgBox "Ha intentado registrar a un usuario ya existente!...", vbCritical + vbOKOnly, "Error..."
        'MsgBox Err.Number & ": " & Err.Description
    End If
End Sub

Private Sub cmdNuevo_Click()
    VaciaCampos
    BotonesConfirma
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblFechaCrea = Date
    cmdBuscarFuncionario.Enabled = False
    FraDatos.Enabled = False
    Editando = False
    'Abrimos la tabla de niveles de acceso
    rsNivelAcceso.Open "Select Distinct IdNivelAcceso, DesNivelAcceso From NivelAcceso", db, adOpenStatic
    If rsNivelAcceso.RecordCount = 0 Then
        cmbNivelAcceso.AddItem "No existen datos"
    Else
        i = 0
        While Not rsNivelAcceso.EOF
            cmbNivelAcceso.AddItem rsNivelAcceso!IdNivelAcceso & "  " & rsNivelAcceso!DesNivelAcceso
            vecNivelAcceso(i) = rsNivelAcceso!IdNivelAcceso & "  " & rsNivelAcceso!DesNivelAcceso
            i = i + 1
            rsNivelAcceso.MoveNext
        Wend
        rsNivelAcceso.Close
    End If
    'Abrimos la tabla de usuarios
    rsUsuarios.Open "Select * From Usuarios_Udapre", db, adOpenKeyset, adLockOptimistic
    If rsUsuarios.RecordCount = 0 Then
       BotonesInicio
    Else
       RecuperaUsuario
       BotonesNavegar
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsUsuarios.Close
    If rsFuncionarios.State = 1 Then rsFuncionarios.Close
End Sub

Private Sub RecuperaUsuario()
    lblIdFuncionario = rsUsuarios!IdFuncionario
    If rsFuncionarios.State = 1 Then rsFuncionarios.Close
    rsFuncionarios.Open "Select * From rc_Personal Where IdFuncionario=" & lblIdFuncionario, db, adOpenStatic
    If rsFuncionarios.RecordCount = 1 Then
        lblPaterno = IIf(IsNull(rsFuncionarios!Paterno), "", rsFuncionarios!Paterno)
        lblMaterno = IIf(IsNull(rsFuncionarios!Materno), "", rsFuncionarios!Materno)
        lblNombres = IIf(IsNull(rsFuncionarios!Nombres), "", rsFuncionarios!Nombres)
        Txtusuario = rsUsuarios!usr_usuario
        cmbNivelAcceso.ListIndex = BuscaNivelAcceso(rsUsuarios!NivelAcceso)
        If CBool(rsUsuarios!Usr_Activo) Then
            chkUsr_Activo.Value = 1
        Else
            chkUsr_Activo.Value = 0
        End If
        lblFechaCrea = IIf(IsNull(rsUsuarios!FechaCrea), Format(Date, "dd/mm/yyyy"), rsUsuarios!FechaCrea)
    End If
End Sub

Private Sub VaciaCampos()
    cmdBuscarFuncionario.Enabled = True
    FraDatos.Enabled = True
    lblIdFuncionario = ""
    lblPaterno = ""
    lblMaterno = ""
    lblNombres = ""
    Txtusuario = ""
    lblFechaCrea = Date
End Sub

Private Function ValidaCampos() As Boolean
    ValidaCampos = True
    If lblIdFuncionario = "" Then
        MsgBox "Debe elegir un funcionario!... ", vbInformation + vbOKOnly, "Atención"
        cmdBuscarFuncionario.SetFocus
        ValidaCampos = False
        Exit Function
    End If
    If Txtusuario = "" Then
        MsgBox "Debe introducir un nombre de usuario", vbInformation + vbOKOnly, "Atención"
        Txtusuario.SetFocus
        ValidaCampos = False
        Exit Function
    End If
    If cmbNivelAcceso.Text = "" Then
        MsgBox "Debe introducir el nivel de acceso", vbInformation + vbOKOnly, "Atención"
        cmbNivelAcceso.SetFocus
        ValidaCampos = False
        Exit Function
    End If
End Function

Private Sub BotonesConfirma()
On Error Resume Next
    cmdNuevo.Enabled = False
    CmdEditar.Enabled = False
    CmdGrabar.Enabled = True
    CmdCancelar.Enabled = True
    cmdEliminar.Enabled = False
    CmdSalir.Enabled = False
    cmdBuscaUsuario.Enabled = False
    FraDatos.Enabled = True
End Sub

Private Sub BotonesNavegar()
On Error Resume Next
    cmdNuevo.Enabled = True
    CmdEditar.Enabled = True
    CmdGrabar.Enabled = False
    CmdCancelar.Enabled = False
    cmdEliminar.Enabled = True
    CmdSalir.Enabled = True
    cmdBuscaUsuario.Enabled = True
    FraDatos.Enabled = False
End Sub

Private Sub BotonesInicio()
On Error Resume Next
    cmdNuevo.Enabled = True
    CmdEditar.Enabled = False
    CmdGrabar.Enabled = False
    CmdCancelar.Enabled = False
    cmdEliminar.Enabled = False
    CmdSalir.Enabled = False
    cmdBuscaUsuario.Enabled = False
End Sub

Private Function BuscaNivelAcceso(pNivelAcceso As Integer) As Integer
Dim j As Integer
    j = 0
    While j < i And CInt(Mid(vecNivelAcceso(j), 1, 3)) <> pNivelAcceso
            j = j + 1
    Wend
    BuscaNivelAcceso = j
End Function

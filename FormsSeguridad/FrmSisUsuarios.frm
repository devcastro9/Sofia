VERSION 5.00
Begin VB.Form FrmSisUsuarios 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuario del Sistema"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "FrmSisUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5985
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtnBuscar 
      BackColor       =   &H8000000A&
      Caption         =   "Buscar"
      Height          =   720
      Left            =   240
      Picture         =   "FrmSisUsuarios.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Busca un Registro"
      Top             =   2400
      Width           =   765
   End
   Begin VB.CommandButton BtnSalir 
      BackColor       =   &H8000000A&
      Caption         =   "Cerrar"
      Height          =   675
      Left            =   240
      Picture         =   "FrmSisUsuarios.frx":09FA
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3360
      Width           =   765
   End
   Begin VB.CommandButton BtnEliminar 
      BackColor       =   &H8000000A&
      Caption         =   "Anular"
      Height          =   675
      Left            =   240
      Picture         =   "FrmSisUsuarios.frx":0C04
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Anula Registro Activo"
      Top             =   1500
      Width           =   765
   End
   Begin VB.CommandButton BtnModificar 
      BackColor       =   &H8000000A&
      Caption         =   "Modificar"
      Height          =   675
      Left            =   240
      Picture         =   "FrmSisUsuarios.frx":18CE
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Modifica Registro Activo"
      Top             =   840
      Width           =   765
   End
   Begin VB.CommandButton BtnAñadir 
      BackColor       =   &H8000000A&
      Caption         =   "Nuevo"
      Height          =   675
      Left            =   240
      Picture         =   "FrmSisUsuarios.frx":1EAE
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Nuevo Registro"
      Top             =   180
      Width           =   765
   End
   Begin VB.Frame FraAcceso 
      Height          =   1890
      Left            =   1245
      TabIndex        =   17
      Top             =   2265
      Width           =   4695
      Begin VB.ComboBox cmbNivelAcceso 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   870
         Width           =   3180
      End
      Begin VB.CheckBox chkUsr_Activo 
         Caption         =   "Usuario Activo"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   90
         TabIndex        =   6
         Top             =   1455
         Width           =   1335
      End
      Begin VB.TextBox txtUsuario 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   765
         MaxLength       =   15
         TabIndex        =   3
         Top             =   240
         Width           =   1380
      End
      Begin VB.TextBox txtClave 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2925
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   240
         Width           =   1620
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Creación:"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1830
         TabIndex        =   21
         Top             =   1470
         Width           =   1170
      End
      Begin VB.Label lblFechaCrea 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   1470
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nivel de Acceso:"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   1
         Left            =   2370
         TabIndex        =   18
         Top             =   270
         Width           =   450
      End
   End
   Begin VB.Frame FraDatos 
      ForeColor       =   &H00000000&
      Height          =   2235
      Left            =   1245
      TabIndex        =   7
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdBuscarFuncionario 
         Caption         =   "..."
         Height          =   255
         Left            =   2475
         Picture         =   "FrmSisUsuarios.frx":24D2
         TabIndex        =   8
         ToolTipText     =   "Buscar Funcionarios"
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtNombres 
         Height          =   285
         Left            =   1500
         MaxLength       =   15
         TabIndex        =   0
         Top             =   720
         Width           =   3075
      End
      Begin VB.TextBox txtPaterno 
         Height          =   285
         Left            =   1500
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1200
         Width           =   3075
      End
      Begin VB.TextBox txtMaterno 
         Height          =   285
         Left            =   1500
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1680
         Width           =   3075
      End
      Begin VB.Label lblNombres 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1500
         TabIndex        =   16
         Top             =   720
         Width           =   3075
      End
      Begin VB.Label lblMaterno 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1500
         TabIndex        =   15
         Top             =   1680
         Width           =   3075
      End
      Begin VB.Label lblPaterno 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1500
         TabIndex        =   14
         Top             =   1200
         Width           =   3075
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Id. Funcionario:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblIdFuncionario 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1620
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombres:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   750
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Segundo Apellido :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Primer Apellido :"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1230
         Width           =   1125
      End
   End
   Begin VB.Frame FraVerificaClave 
      Caption         =   " Verificación de clave de acceso "
      Height          =   885
      Left            =   1245
      TabIndex        =   23
      Top             =   2790
      Width           =   4695
      Begin VB.TextBox txtVerificaClave 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2010
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   1365
         TabIndex        =   25
         Top             =   390
         Width           =   450
      End
   End
   Begin VB.CommandButton BtnGrabar 
      BackColor       =   &H8000000A&
      Caption         =   "Grabar"
      Height          =   675
      Left            =   240
      Picture         =   "FrmSisUsuarios.frx":2914
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1200
      Width           =   765
   End
   Begin VB.CommandButton BtnCancelar 
      BackColor       =   &H8000000A&
      Caption         =   "Cancelar"
      Height          =   675
      Left            =   240
      Picture         =   "FrmSisUsuarios.frx":2B1E
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2100
      Width           =   765
   End
End
Attribute VB_Name = "FrmSisUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUsuarios As New ADODB.Recordset
Dim rsAuxUsuarios As New ADODB.Recordset
Dim rsFuncionarios As New ADODB.Recordset
Dim rsNivelAcceso As New ADODB.Recordset
Dim Editando As Boolean
Dim i As Byte
Dim vecNivelAcceso(99) As Integer 'Máximo se podrá cargar 99 niveles de acceso

Private Sub cmdBuscarFuncionario_Click()
    FrmBuscaFuncionario.Show vbModal
    If GlElegido <> "" Then
        rsFuncionarios.Open "Select * From rc_Personal Where IdFuncionario=" & GlElegido & " Order by Paterno", db, adOpenStatic
        lblIdFuncionario = rsFuncionarios!idfuncionario
        lblPaterno = rsFuncionarios!paterno
        lblMaterno = rsFuncionarios!materno
        lblNombres = rsFuncionarios!NombreS
        txtUsuario = Mid(rsFuncionarios!NombreS, 1, 1) & "_" & (rsFuncionarios!paterno)
        txtUsuario = BuscaPerfiles(rsFuncionarios!idfuncionario, txtUsuario)
        txtUsuario.Enabled = False
        rsFuncionarios.Close
        MostrarTextBox False
    End If
End Sub

Private Sub BtnBuscar_Click()
Dim Encontrado As Boolean
    FrmBuscaUsuario.Show vbModal
    If GlElegido <> "" Then
        Encontrado = False
        rsUsuarios.MoveFirst
        While Not rsUsuarios.EOF And Not Encontrado
            If rsUsuarios!usr_usuario = GlElegido Then
                Encontrado = True
            Else
                rsUsuarios.MoveNext
            End If
        Wend
        If Encontrado Then RecuperaUsuario
    End If
End Sub

Private Sub BtnCancelar_Click()
    cmdBuscarFuncionario.Enabled = False
    If rsUsuarios.RecordCount = 0 Then
        BotonesInicio
    Else
        If rsUsuarios.BOF Or rsUsuarios.EOF Then
            rsUsuarios.MoveFirst
        End If
        RecuperaUsuario
        BotonesNavegar
    End If
    FraDatos.Enabled = False
    FraAcceso.Enabled = False
    FraAcceso.Visible = True
    FraVerificaClave.Visible = False
End Sub

Private Sub BtnModificar_Click()
    Editando = True
    txtUsuario.Enabled = False
    FraDatos.Enabled = True
    FraAcceso.Enabled = True
    If lblIdFuncionario = "0" Then
        txtNombres = lblNombres
        txtPaterno = lblPaterno
        txtMaterno = lblMaterno
        txtUsuario.Enabled = True
        MostrarTextBox True
    End If
    BotonesConfirma
End Sub

Private Sub BtnEliminar_Click()
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

Private Sub BtnGrabar_Click()
On Error GoTo QueError
    If ValidaCampos Then
        cmdBuscarFuncionario.Enabled = False
        If Not Editando Then rsUsuarios.AddNew
        If lblIdFuncionario = "0" Then
            lblNombres = txtNombres
            lblPaterno = txtPaterno
            lblMaterno = txtMaterno
            'txtUsuario = Mid(txtNombres, 1, 1) & "_" & txtPaterno
        End If
        db.BeginTrans
        rsUsuarios!usr_usuario = txtUsuario
        rsUsuarios!idfuncionario = CInt(lblIdFuncionario)
        rsUsuarios!NombreS = lblNombres
        rsUsuarios!paterno = lblPaterno
        rsUsuarios!materno = lblMaterno
        rsUsuarios!usr_clave = Encriptar(txtClave)
        rsUsuarios!IdNivelAcceso = CInt(Mid(cmbNivelAcceso, 1, 2))
        rsUsuarios!Usr_Activo = CBool(chkUsr_Activo.Value)
        rsUsuarios!FechaCrea = Date 'CDate(lblFechaCrea)
        lblFechaCrea = Date
        rsUsuarios.Update
        db.CommitTrans
        Editando = False
        FraDatos.Enabled = False
        FraAcceso.Enabled = False
       
        BotonesNavegar
        FraAcceso.Visible = True
        FraVerificaClave.Visible = False
        MsgBox "Datos del usuario grabado satisfactoriamente", vbInformation + vbOKOnly, "Atención"
    End If
    Exit Sub
QueError:
    db.RollbackTrans
    rsUsuarios.CancelUpdate
    FraAcceso.Visible = True
    FraVerificaClave.Visible = False
    If Err.Number = -2147467259 Then
        MsgBox "Ha intentado registrar ha un usuario ya existente!...", vbCritical + vbOKOnly, "Error..."
    End If
    If Err.Number = -2147217887 Then
        MsgBox "Se producjo un error desconocido!...", vbCritical + vbOKOnly, "Error..."
    End If
End Sub

Private Sub BtnAñadir_Click()
    VaciaCampos
    txtUsuario.Enabled = True
    MostrarTextBox True
    BotonesConfirma
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Habilitacion de Frames
    FraAcceso.Visible = True
    FraVerificaClave.Visible = False
   
    lblFechaCrea = Date
    cmdBuscarFuncionario.Enabled = False
    FraDatos.Enabled = False
    FraAcceso.Enabled = False
    txtUsuario.Enabled = True
    Editando = False
    'Abrimos la tabla de niveles de acceso
    If rsNivelAcceso.State = 1 Then rsNivelAcceso.Close
    rsNivelAcceso.Open "Select Distinct IdNivelAcceso, DesNivelAcceso Fromgc_nivelacceso", db, adOpenStatic
    If rsNivelAcceso.RecordCount = 0 Then
        MsgBox "ATENCION: No existe registros de Niveles de Acceso o se han borrado !!" & vbCr & _
               "Ingrese nuevos Niveles de Acceso para operar normalmente con el sistema.", vbCritical + vbOKOnly, "Atención"
        BotonesInicio
        BtnAñadir.Visible = False
        Exit Sub
    Else
        i = 0
        While Not rsNivelAcceso.EOF
            cmbNivelAcceso.AddItem rsNivelAcceso!IdNivelAcceso & "  " & rsNivelAcceso!DesNivelAcceso
            vecNivelAcceso(i) = rsNivelAcceso!IdNivelAcceso
            i = i + 1
            rsNivelAcceso.MoveNext
        Wend
        rsNivelAcceso.Close
    End If
    
    'Abrimos la tabla de usuarios
    rsUsuarios.Open "Select * From gc_Usuarios", db, adOpenKeyset, adLockOptimistic
    If rsUsuarios.RecordCount = 0 Then
       BotonesInicio
    Else
       'rsUsuarios.MoveNext
       RecuperaUsuario
       BotonesNavegar
    End If
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsUsuarios.State = 1 Then rsUsuarios.Close
    If rsFuncionarios.State = 1 Then rsFuncionarios.Close
End Sub

Private Sub RecuperaUsuario()
    lblIdFuncionario = rsUsuarios!idfuncionario
    lblPaterno = IIf(IsNull(rsUsuarios!paterno), "", rsUsuarios!paterno)
    lblMaterno = IIf(IsNull(rsUsuarios!materno), "", rsUsuarios!materno)
    lblNombres = IIf(IsNull(rsUsuarios!NombreS), "", rsUsuarios!NombreS)
    txtUsuario = rsUsuarios!usr_usuario
    txtClave = IIf(IsNull(rsUsuarios!usr_clave), "", Desencriptar(rsUsuarios!usr_clave))
    cmbNivelAcceso.ListIndex = BuscaNivelAcceso(rsUsuarios!IdNivelAcceso)
    If CBool(rsUsuarios!Usr_Activo) Then
        chkUsr_Activo.Value = 1
    Else
        chkUsr_Activo.Value = 0
    End If
    lblFechaCrea = IIf(IsNull(rsUsuarios!FechaCrea), Format(Date, "dd/mm/yyyy"), rsUsuarios!FechaCrea)
    If rsUsuarios!idfuncionario = 0 Then
        txtNombres = lblNombres
        txtPaterno = lblPaterno
        txtMaterno = lblMaterno
        MostrarTextBox True
    Else
        MostrarTextBox False
    End If
End Sub

Private Sub VaciaCampos()
    cmdBuscarFuncionario.Enabled = True
    FraDatos.Enabled = True
    FraAcceso.Enabled = True
    lblIdFuncionario = "0"
    lblPaterno = ""
    txtPaterno = ""
    lblMaterno = ""
    txtMaterno = ""
    lblNombres = ""
    txtNombres = ""
    txtUsuario = ""
    txtClave = ""
    txtVerificaClave = ""
    chkUsr_Activo = 0
    lblFechaCrea = Date
End Sub

Private Function ValidaCampos() As Boolean
    ValidaCampos = True
    If txtUsuario = "" Then
        MsgBox "Debe introducir el Login de usuario", vbInformation + vbOKOnly, "Atención"
        txtUsuario.SetFocus
        ValidaCampos = False
        Exit Function
    End If
    If cmbNivelAcceso.Text = "" Then
        MsgBox "Debe introducir el nivel de acceso para el usuario", vbInformation + vbOKOnly, "Atención"
        cmbNivelAcceso.SetFocus
        ValidaCampos = False
        Exit Function
    End If
    
    FraAcceso.Visible = False
    FraVerificaClave.Visible = True
    If Encriptar(txtClave) <> rsUsuarios!usr_clave Then
      If txtClave <> txtVerificaClave Then
          MsgBox "Clave de acceso diferente. Intente nuevamente.", vbInformation + vbOKOnly, "Atención"
          txtVerificaClave.SetFocus
          ValidaCampos = False
          Exit Function
      End If
    End If
End Function

Private Sub BotonesConfirma()
On Error Resume Next
    BtnAñadir.Visible = False
    BtnModificar.Visible = False
    BtnGrabar.Visible = True
    BtnCancelar.Visible = True
    BtnEliminar.Visible = False
    BtnSalir.Visible = False
    BtnBuscar.Visible = False
    FraDatos.Enabled = True
    FraAcceso.Enabled = True
End Sub

Private Sub BotonesNavegar()
On Error Resume Next
    BtnAñadir.Visible = True
    BtnModificar.Visible = True
    BtnGrabar.Visible = False
    BtnCancelar.Visible = False
    BtnEliminar.Visible = True
    BtnSalir.Visible = True
    BtnBuscar.Visible = True
    FraDatos.Enabled = False
    FraAcceso.Enabled = False
End Sub

Private Sub BotonesInicio()
On Error Resume Next
    BtnAñadir.Visible = True
    BtnModificar.Visible = False
    BtnGrabar.Visible = False
    BtnCancelar.Visible = False
    BtnEliminar.Visible = False
    BtnSalir.Visible = True
    BtnBuscar.Visible = False
End Sub

Private Function BuscaNivelAcceso(pNivelAcceso As Integer) As Integer
Dim j As Integer
    j = 0
    While j < i And vecNivelAcceso(j) <> pNivelAcceso
            j = j + 1
    Wend
    BuscaNivelAcceso = j
End Function

Private Sub MostrarTextBox(SW As Boolean)
    txtNombres.Visible = SW
    txtPaterno.Visible = SW
    txtMaterno.Visible = SW
End Sub

Function BuscaPerfiles(pIdFuncionario As Integer, pUsuario As String) As String
Dim MaxPerfil As Integer
Dim NroPerfil As String
    MaxPerfil = 0
    rsAuxUsuarios.Open "Select * From GC_Usuarios Where IdFuncionario=" & pIdFuncionario, db, adOpenStatic
    If rsAuxUsuarios.RecordCount = 0 Then
        BuscaPerfiles = pUsuario
    Else
        While Not rsAuxUsuarios.EOF
            NroPerfil = Mid(rsAuxUsuarios!usr_usuario, Len(rsAuxUsuarios!usr_usuario), 1)
            If IsNumeric(NroPerfil) Then
                If CInt(NroPerfil) > MaxPerfil Then MaxPerfil = CInt(NroPerfil)
            End If
            rsAuxUsuarios.MoveNext
        Wend
        BuscaPerfiles = pUsuario & MaxPerfil + 1
    End If
    rsAuxUsuarios.Close
End Function

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 And KeyAscii > 165) Then
        Beep
        KeyAscii = 0
    End If
End Sub

'Option Explicit
'Dim rsUsuarios As New ADODB.Recordset
'Dim rsAuxUsuarios As New ADODB.Recordset
'Dim rsFuncionarios As New ADODB.Recordset
'Dim rsNivelAcceso As New ADODB.Recordset
'Dim Editando As Boolean
'Dim i As Byte
'Dim vecNivelAcceso(99) As Integer 'Máximo se podrá cargar 99 niveles de acceso
'
'Private Sub cmdBuscarFuncionario_Click()
'    FrmBuscaFuncionario.Show vbModal
'    If GlElegido <> "" Then
'        rsFuncionarios.Open "Select * From rc_Personal Where IdFuncionario=" & GlElegido & " Order by Paterno", db, adOpenStatic
'        lblIdFuncionario = rsFuncionarios!idfuncionario
'        lblPaterno = rsFuncionarios!paterno
'        lblMaterno = rsFuncionarios!materno
'        lblNombres = rsFuncionarios!nombres
'        txtUsuario = Mid(rsFuncionarios!nombres, 1, 1) & "_" & (rsFuncionarios!paterno)
'        txtUsuario = BuscaPerfiles(rsFuncionarios!idfuncionario, txtUsuario)
'        txtUsuario.Enabled = False
'        rsFuncionarios.Close
'        MostrarTextBox False
'    End If
'End Sub
'
'Private Sub BtnBuscar_Click()
'Dim Encontrado As Boolean
'    FrmBuscaUsuario.Show vbModal
'    If GlElegido <> "" Then
'        Encontrado = False
'        rsUsuarios.MoveFirst
'        While Not rsUsuarios.EOF And Not Encontrado
'            If rsUsuarios!usr_usuario = GlElegido Then
'                Encontrado = True
'            Else
'                rsUsuarios.MoveNext
'            End If
'        Wend
'        If Encontrado Then RecuperaUsuario
'    End If
'End Sub
'
'Private Sub BtnCancelar_Click()
'    cmdBuscarFuncionario.Enabled = False
'    If rsUsuarios.RecordCount = 0 Then
'        BotonesInicio
'    Else
'        If rsUsuarios.BOF Or rsUsuarios.EOF Then
'            rsUsuarios.MoveFirst
'        End If
'        RecuperaUsuario
'        BotonesNavegar
'    End If
'    FraDatos.Enabled = False
'    FraAcceso.Enabled = False
'End Sub
'
'Private Sub BtnModificar_Click()
'    Editando = True
'    txtUsuario.Enabled = False
'    FraDatos.Enabled = True
'    FraAcceso.Enabled = True
'    If lblIdFuncionario = "0" Then
'        txtNombres = lblNombres
'        txtPaterno = lblPaterno
'        txtMaterno = lblMaterno
'        txtUsuario.Enabled = True
'        MostrarTextBox True
'    End If
'    BotonesConfirma
'End Sub
'
'Private Sub BtnEliminar_Click()
'If MsgBox("Esta seguro de eliminar al usuario visualizado?", vbExclamation + vbYesNo, "Atención") = vbYes Then
'    rsUsuarios.Delete
'    If rsUsuarios.RecordCount > 0 Then
'        rsUsuarios.MoveNext
'        If rsUsuarios.EOF Then rsUsuarios.MoveLast
'        RecuperaUsuario
'    Else
'        VaciaCampos
'        BotonesInicio
'    End If
'End If
'End Sub
'
'Private Sub BtnGrabar_Click()
'On Error GoTo QueError
'    If ValidaCampos Then
'        cmdBuscarFuncionario.Enabled = False
'        If Not Editando Then rsUsuarios.AddNew
'        If lblIdFuncionario = "0" Then
'            lblNombres = txtNombres
'            lblPaterno = txtPaterno
'            lblMaterno = txtMaterno
'            'txtUsuario = Mid(txtNombres, 1, 1) & "_" & txtPaterno
'        End If
'        rsUsuarios!idfuncionario = CInt(lblIdFuncionario)
'        rsUsuarios!nombres = lblNombres
'        rsUsuarios!paterno = lblPaterno
'        rsUsuarios!materno = lblMaterno
'        rsUsuarios!usr_usuario = txtUsuario
'        If rsUsuarios!usr_clave <> txtClave Or IsNull(rsUsuarios!usr_clave) Then rsUsuarios!usr_clave = Encriptar(txtClave)
'        rsUsuarios!IdNivelAcceso = CInt(Mid(cmbNivelAcceso, 1, 2))
'        rsUsuarios!Usr_Activo = CBool(chkUsr_Activo.Value)
'        rsUsuarios!FechaCrea = CDate(lblFechaCrea)
'        db.BeginTrans
'        rsUsuarios.Update
'        db.CommitTrans
'        Editando = False
'        FraDatos.Enabled = False
'        FraAcceso.Enabled = False
'
'        BotonesNavegar
'        MsgBox "Datos del usuario grabado satisfactoriamente", vbInformation + vbOKOnly, "Atención"
'    End If
'    Exit Sub
'QueError:
'    If err.Number = -2147467259 Then
'        db.RollbackTrans
'        rsUsuarios.CancelUpdate
'        MsgBox "Ha intentado registrar ha un usuario ya existente!...", vbCritical + vbOKOnly, "Error..."
'    End If
'End Sub
'
'Private Sub BtnAñadir_Click()
'    VaciaCampos
'    txtUsuario.Enabled = True
'    MostrarTextBox True
'    BotonesConfirma
'End Sub
'
'Private Sub BtnSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
    'Habilitacion de Frames
    FraAcceso.Visible = True
    FraVerificaClave.Visible = False
   
    lblFechaCrea = Date
    cmdBuscarFuncionario.Enabled = False
    FraDatos.Enabled = False
    FraAcceso.Enabled = False
    txtUsuario.Enabled = True
    Editando = False
    'Abrimos la tabla de niveles de acceso
    If rsNivelAcceso.State = 1 Then rsNivelAcceso.Close
    rsNivelAcceso.Open "Select Distinct IdNivelAcceso, DesNivelAcceso Fromgc_nivelacceso", db, adOpenStatic
    If rsNivelAcceso.RecordCount = 0 Then
        MsgBox "ATENCION: No existe registros de Niveles de Acceso o se han borrado !!" & vbCr & _
               "Ingrese nuevos Niveles de Acceso para operar normalmente con el sistema.", vbCritical + vbOKOnly, "Atención"
        BotonesInicio
        BtnAñadir.Visible = False
        Exit Sub
    Else
        i = 0
        While Not rsNivelAcceso.EOF
            cmbNivelAcceso.AddItem rsNivelAcceso!IdNivelAcceso & "  " & rsNivelAcceso!DesNivelAcceso
            vecNivelAcceso(i) = rsNivelAcceso!IdNivelAcceso
            i = i + 1
            rsNivelAcceso.MoveNext
        Wend
        rsNivelAcceso.Close
    End If
    
    'Abrimos la tabla de usuarios
    rsUsuarios.Open "Select * From gc_Usuarios", db, adOpenKeyset, adLockOptimistic
    If rsUsuarios.RecordCount = 0 Then
       BotonesInicio
    Else
       'rsUsuarios.MoveNext
       RecuperaUsuario
       BotonesNavegar
    End If
	Call SeguridadSet(Me)
End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    If rsUsuarios.State = 1 Then rsUsuarios.Close
'    If rsFuncionarios.State = 1 Then rsFuncionarios.Close
'End Sub
'
'Private Sub RecuperaUsuario()
'    lblIdFuncionario = rsUsuarios!idfuncionario
'    lblPaterno = IIf(IsNull(rsUsuarios!paterno), "", rsUsuarios!paterno)
'    lblMaterno = IIf(IsNull(rsUsuarios!materno), "", rsUsuarios!materno)
'    lblNombres = IIf(IsNull(rsUsuarios!nombres), "", rsUsuarios!nombres)
'    txtUsuario = rsUsuarios!usr_usuario
'    txtClave = IIf(IsNull(rsUsuarios!usr_clave), "", rsUsuarios!usr_clave)
'    cmbNivelAcceso.ListIndex = BuscaNivelAcceso(rsUsuarios!IdNivelAcceso)
'    If CBool(rsUsuarios!Usr_Activo) Then
'        chkUsr_Activo.Value = 1
'    Else
'        chkUsr_Activo.Value = 0
'    End If
'    lblFechaCrea = IIf(IsNull(rsUsuarios!FechaCrea), Format(Date, "dd/mm/yyyy"), rsUsuarios!FechaCrea)
'    If rsUsuarios!idfuncionario = 0 Then
'        txtNombres = lblNombres
'        txtPaterno = lblPaterno
'        txtMaterno = lblMaterno
'        MostrarTextBox True
'    Else
'        MostrarTextBox False
'    End If
'End Sub
'
'Private Sub VaciaCampos()
'    cmdBuscarFuncionario.Enabled = True
'    FraDatos.Enabled = True
'    FraAcceso.Enabled = True
'    lblIdFuncionario = "0"
'    lblPaterno = ""
'    txtPaterno = ""
'    lblMaterno = ""
'    txtMaterno = ""
'    lblNombres = ""
'    txtNombres = ""
'    txtUsuario = ""
'    txtClave = ""
'    chkUsr_Activo = 0
'    lblFechaCrea = Date
'End Sub
'
'Private Function ValidaCampos() As Boolean
'    ValidaCampos = True
'    If txtUsuario = "" Then
'        MsgBox "Debe introducir el Login de usuario", vbInformation + vbOKOnly, "Atención"
'        txtUsuario.SetFocus
'        ValidaCampos = False
'        Exit Function
'    End If
'    If cmbNivelAcceso.Text = "" Then
'        MsgBox "Debe introducir el nivel de acceso para el usuario", vbInformation + vbOKOnly, "Atención"
'        cmbNivelAcceso.SetFocus
'        ValidaCampos = False
'        Exit Function
'    End If
'End Function
'
'Private Sub BotonesConfirma()
'On Error Resume Next
'    BtnAñadir.Enabled = False
'    BtnModificar.Enabled = False
'    BtnGrabar.Enabled = True
'    BtnCancelar.Enabled = True
'    BtnEliminar.Enabled = False
'    BtnSalir.Enabled = False
'    BtnBuscar.Enabled = False
'    FraDatos.Enabled = True
'    FraAcceso.Enabled = True
'End Sub
'
'Private Sub BotonesNavegar()
'On Error Resume Next
'    BtnAñadir.Enabled = True
'    BtnModificar.Enabled = True
'    BtnGrabar.Enabled = False
'    BtnCancelar.Enabled = False
'    BtnEliminar.Enabled = True
'    BtnSalir.Enabled = True
'    BtnBuscar.Enabled = True
'    FraDatos.Enabled = False
'    FraAcceso.Enabled = False
'End Sub
'
'Private Sub BotonesInicio()
'On Error Resume Next
'    BtnAñadir.Enabled = True
'    BtnModificar.Enabled = False
'    BtnGrabar.Enabled = False
'    BtnCancelar.Enabled = False
'    BtnEliminar.Enabled = False
'    BtnSalir.Enabled = True
'    BtnBuscar.Enabled = False
'End Sub
'
'Private Function BuscaNivelAcceso(pNivelAcceso As Integer) As Integer
'Dim j As Integer
'    j = 0
'    While j < i And vecNivelAcceso(j) <> pNivelAcceso
'            j = j + 1
'    Wend
'    BuscaNivelAcceso = j
'End Function
'
'Private Sub MostrarTextBox(Sw As Boolean)
'    txtNombres.Visible = Sw
'    txtPaterno.Visible = Sw
'    txtMaterno.Visible = Sw
'End Sub
'
'Function BuscaPerfiles(pIdFuncionario As Integer, pUsuario As String) As String
'Dim MaxPerfil As Integer
'Dim NroPerfil As String
'    MaxPerfil = 0
'    rsAuxUsuarios.Open "Select * From GC_Usuarios Where IdFuncionario=" & pIdFuncionario, db, adOpenStatic
'    If rsAuxUsuarios.RecordCount = 0 Then
'        BuscaPerfiles = pUsuario
'    Else
'        While Not rsAuxUsuarios.EOF
'            NroPerfil = Mid(rsAuxUsuarios!usr_usuario, Len(rsAuxUsuarios!usr_usuario), 1)
'            If IsNumeric(NroPerfil) Then
'                If CInt(NroPerfil) > MaxPerfil Then MaxPerfil = CInt(NroPerfil)
'            End If
'            rsAuxUsuarios.MoveNext
'        Wend
'        BuscaPerfiles = pUsuario & MaxPerfil + 1
'    End If
'    rsAuxUsuarios.Close
'End Function
'
'Private Sub txtClave_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 48 And KeyAscii > 165) Then
'        Beep
'        KeyAscii = 0
'    End If
'End Sub

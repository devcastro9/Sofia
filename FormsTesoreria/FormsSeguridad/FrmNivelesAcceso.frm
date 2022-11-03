VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmNivelesAcceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Nivel de Acceso"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "FrmNivelesAcceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6345
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstPrivacceso 
      Height          =   645
      ItemData        =   "FrmNivelesAcceso.frx":0442
      Left            =   5160
      List            =   "FrmNivelesAcceso.frx":0444
      TabIndex        =   17
      Top             =   1140
      Width           =   855
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "E&liminar"
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   4800
      Width           =   990
   End
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   600
      TabIndex        =   15
      Top             =   4335
      Width           =   375
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5220
      TabIndex        =   9
      Top             =   4800
      Width           =   990
   End
   Begin VB.TextBox txtDesNivelAcceso 
      Height          =   285
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   8
      Top             =   105
      Width           =   3975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1140
      TabIndex        =   6
      Top             =   4800
      Width           =   990
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   4800
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   3180
      TabIndex        =   4
      Top             =   4800
      Width           =   990
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   990
   End
   Begin MSDataGridLib.DataGrid dtgMenuSistema 
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   12708861
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
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
      Caption         =   "MENU SISTEMA SAF-2000"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "DescOpcMenu"
         Caption         =   "Opción de Menu"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "EsTerminal"
         Caption         =   "Terminal"
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
      BeginProperty Column02 
         DataField       =   "Habilitado"
         Caption         =   "Habilitar"
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
         DataField       =   "IdPrivAcceso"
         Caption         =   "Acceso"
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
            Locked          =   -1  'True
            ColumnWidth     =   3014.929
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Button          =   -1  'True
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Button          =   -1  'True
            ColumnWidth     =   854.929
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Acceso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1245
      TabIndex        =   13
      Top             =   4395
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "APR=Aprobacion"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   4005
      TabIndex        =   12
      Top             =   4395
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "CON=Consuta"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2865
      TabIndex        =   11
      Top             =   4395
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "TOT=Total"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1995
      TabIndex        =   10
      Top             =   4395
      Width           =   780
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1080
      Top             =   4320
      Width           =   5175
   End
   Begin VB.Label lblDesNivelAcceso 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblNivelAcceso 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nivel de Acceso:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmNivelesAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMenuSistema As New ADODB.Recordset
Dim rsNivelAcceso As New ADODB.Recordset
Dim rsNiveles As New ADODB.Recordset
Dim rsPrivAcceso As New ADODB.Recordset
Dim vIdNivelAcceso As Integer
Dim vMaxNivelAcceso As Integer
Dim vEditando As Boolean
Dim Fila As Integer

Private Sub CmdAnterior_Click()
    rsNiveles.MovePrevious
    If rsNiveles.BOF Then
       rsNiveles.MoveFirst
       Exit Sub
    End If
    vIdNivelAcceso = rsNiveles!idNivelAcceso
    lblNivelAcceso = vIdNivelAcceso
    If rsMenuSistema.State = 1 Then
        If rsMenuSistema.EditMode <> adEditNone Then rsMenuSistema.CancelUpdate
        rsMenuSistema.Close
    End If
    rsMenuSistema.Open "Select na.IdNivelAcceso, ms.descopcmenu, ms.EsTerminal, na.Habilitado, na.DesNivelAcceso, na.IdPrivAcceso From Menusistema ms, idNivelAcceso na Where ms.NombOpcMenu like na.NombOpcMenu and na.IdNivelAcceso=" & vIdNivelAcceso, db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    lblDesNivelAcceso = rsMenuSistema!DesNivelAcceso
End Sub

Private Sub CmdCancelar_Click()
    If rsMenuSistema.State = 1 Then If rsMenuSistema.EditMode <> adEditNone Then rsMenuSistema.CancelUpdate
    rsMenuSistema.Close
    rsMenuSistema.Open "Select na.IdNivelAcceso, ms.descopcmenu, ms.EsTerminal, na.Habilitado,na.IdPrivAcceso From Menusistema ms, NivelAcceso na Where ms.NombOpcMenu like na.NombOpcMenu and na.IdNivelAcceso=" & vIdNivelAcceso, db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    dtgMenuSistema.Columns(2).Locked = True
    dtgMenuSistema.Columns(2).Button = False
    dtgMenuSistema.Columns(3).Locked = True
    dtgMenuSistema.Columns(3).Button = False
    txtDesNivelAcceso.Visible = False
    lblDesNivelAcceso.Visible = True
    lblNivelAcceso = vIdNivelAcceso
    BotonesNavegar
End Sub

Private Sub Cmdeditar_Click()
    If rsNivelAcceso.State = 1 Then rsNivelAcceso.Close
    rsNivelAcceso.Open "Select * From NivelAcceso Where IdNivelAcceso=" & vIdNivelAcceso, db, adOpenKeyset, adLockOptimistic
    rsNivelAcceso.MoveFirst
    While Not rsNivelAcceso.EOF
        db.Execute "Update MenuSistema Set Habilitado = '" & rsNivelAcceso!Habilitado & "' Where NombOpcMenu = '" & rsNivelAcceso!NombOpcMenu & "'"
        db.Execute "Update Menusistema Set IdPrivAcceso='" & rsNivelAcceso!IdPrivAcceso & "' Where NombOpcMenu = '" & rsNivelAcceso!NombOpcMenu & "'"
        rsNivelAcceso.MoveNext
    Wend
    rsNivelAcceso.MoveFirst
    If rsMenuSistema.State = 1 Then If rsMenuSistema.EditMode <> 0 Then rsMenuSistema.CancelUpdate
    rsMenuSistema.Close
    rsMenuSistema.Open "Select * From Menusistema", db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    vEditando = True
    dtgMenuSistema.Columns(2).Locked = False
    dtgMenuSistema.Columns(2).Button = True
    dtgMenuSistema.Columns(3).Locked = False
    dtgMenuSistema.Columns(3).Button = True
    txtDesNivelAcceso = lblDesNivelAcceso
    txtDesNivelAcceso.Visible = True
    lblDesNivelAcceso.Visible = False
    BotonesConfirma
End Sub

Private Sub cmdEliminar_Click()
  Dim rsUsuarios As New ADODB.Recordset
   rsUsuarios.Open "Select * From Usuarios_Udapre Where NivelAcceso='" & lblNivelAcceso & "'", db, adOpenStatic
  If rsUsuarios.RecordCount > 0 Then
      MsgBox "Existen " & rsUsuarios.RecordCount & " usuarios registrados con este nivel de acceso. Antes de eliminar" & vbCrLf & _
             "el nivel, debe primero cambiar de nivel de acceso de los usuarios afectados.", vbInformation + vbOKOnly, "Atención"
      rsUsuarios.Close
  Else
      rsUsuarios.Close
      If MsgBox("Esta seguro de eliminar este nivel de acceso?", vbCritical + vbYesNo, "Atención") = vbYes Then
         db.Execute "delete from NivelAcceso Where IdNivelAcceso=" & CInt(lblNivelAcceso)
         MsgBox "Nivel de acceso eliminado!", vbInformation + vbOKOnly, "Atención"
         rsNiveles.Requery
         CmdSiguiente_Click
      End If
  End If
End Sub

Private Sub CmdGrabar_Click()
    rsMenuSistema.MoveFirst
    While Not rsMenuSistema.EOF
        If Not vEditando Then
            rsNivelAcceso.AddNew
            rsNivelAcceso!idNivelAcceso = CInt(lblNivelAcceso)
            rsNivelAcceso!NombOpcMenu = rsMenuSistema!NombOpcMenu
            rsNivelAcceso!Habilitado = rsMenuSistema!Habilitado
            rsNivelAcceso!DesNivelAcceso = txtDesNivelAcceso
            rsNivelAcceso!IdPrivAcceso = rsMenuSistema!IdPrivAcceso
            db.BeginTrans
            rsNivelAcceso.Update
            db.CommitTrans
            vMaxNivelAcceso = CInt(lblNivelAcceso)
        Else
            db.Execute "Update NivelAcceso Set Habilitado='" & rsMenuSistema!Habilitado & "' Where IdNivelAcceso= " & vIdNivelAcceso & " and NombOpcMenu='" & rsMenuSistema!NombOpcMenu & "'"
            db.Execute "Update NivelAcceso Set IdPrivAcceso='" & rsMenuSistema!IdPrivAcceso & "' Where IdNivelAcceso= " & vIdNivelAcceso & " and NombOpcMenu='" & rsMenuSistema!NombOpcMenu & "'"
        End If
        rsMenuSistema.MoveNext
    Wend
    rsMenuSistema.Close
    rsMenuSistema.Open "Select na.IdNivelAcceso, ms.descopcmenu, ms.EsTerminal, na.Habilitado,na.IdPrivAcceso From Menusistema ms, NivelAcceso na Where ms.NombOpcMenu like na.NombOpcMenu and na.IdNivelAcceso=" & vIdNivelAcceso, db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    MsgBox "Nivel de acceso grabado satisfactoriamente", vbInformation + vbOKOnly, "Atención"
    dtgMenuSistema.Columns(2).Locked = True
    dtgMenuSistema.Columns(2).Button = False
    dtgMenuSistema.Columns(3).Locked = True
    dtgMenuSistema.Columns(3).Button = False
    txtDesNivelAcceso.Visible = False
    lblDesNivelAcceso.Visible = True
    rsNiveles.Requery
    BotonesNavegar
End Sub

Private Sub cmdNuevo_Click()
    'Abrimos la tabla de niveles de acceso
    If rsNivelAcceso.State = 1 Then rsNivelAcceso.Close
    rsNivelAcceso.Open "Select * From NivelAcceso", db, adOpenKeyset, adLockOptimistic
    
    'Abrimos la tabla de MenuSistema
    db.Execute "Update MenuSistema Set Habilitado='Si'"
    If rsMenuSistema.State = 1 Then rsMenuSistema.Close
    rsMenuSistema.Open "Select * From Menusistema", db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    lblNivelAcceso = vMaxNivelAcceso + 1
    vEditando = False
    dtgMenuSistema.Columns(2).Locked = False
    dtgMenuSistema.Columns(2).Button = True
    dtgMenuSistema.Columns(3).Locked = False
    dtgMenuSistema.Columns(3).Button = True
    txtDesNivelAcceso.Visible = True
    lblDesNivelAcceso.Visible = False
    txtDesNivelAcceso = ""
    BotonesConfirma
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdSiguiente_Click()
    rsNiveles.MoveNext
    If rsNiveles.EOF Then
       rsNiveles.MoveLast
       Exit Sub
    End If
    vIdNivelAcceso = rsNiveles!idNivelAcceso
    lblNivelAcceso = vIdNivelAcceso
    If rsMenuSistema.State = 1 Then
        If rsMenuSistema.EditMode <> adEditNone Then rsMenuSistema.CancelUpdate
        rsMenuSistema.Close
    End If
    rsMenuSistema.Open "Select na.IdNivelAcceso, ms.descopcmenu, ms.EsTerminal, na.Habilitado, na.DesNivelAcceso, na.IdPrivAcceso From Menusistema ms, NivelAcceso na Where ms.NombOpcMenu like na.NombOpcMenu and na.IdNivelAcceso=" & vIdNivelAcceso, db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    lblDesNivelAcceso = rsMenuSistema!DesNivelAcceso
End Sub

Private Sub dtgMenuSistema_ButtonClick(ByVal ColIndex As Integer)
On Error Resume Next
    'Permite la habilitacion de las opciones de menu
    'Si No=No tiene acceso; Si=Si tiene acceso
    If ColIndex = 2 Then
        If dtgMenuSistema.Columns(2).Value = "No" Then
           rsMenuSistema!Habilitado = "Si"
           rsMenuSistema.Update
        Else
           rsMenuSistema!Habilitado = "No"
           rsMenuSistema.Update
        End If
        Set dtgMenuSistema.DataSource = rsMenuSistema
    End If
    'Permite el cambio del estado de acceso a los botones de formulario
    'Si T=Total; Si C=Consulta; Si P=Personalizado
    If ColIndex = 3 Then
        lstPrivacceso.Visible = True
        Fila = dtgMenuSistema.RowBookmark(dtgMenuSistema.Row)
        lstPrivacceso.Top = dtgMenuSistema.RowTop(dtgMenuSistema.Row) + (3 * dtgMenuSistema.RowHeight) + 50
'
'        If dtgMenuSistema.Columns(3).Value = "TOT" Then
'           rsMenuSistema!IdPrivAcceso = "CON"
'           rsMenuSistema.Update
'        ElseIf dtgMenuSistema.Columns(3).Value = "CON" Then
'           rsMenuSistema!IdPrivAcceso = "APR"
'           rsMenuSistema.Update
'        ElseIf dtgMenuSistema.Columns(3).Value = "APR" Then
'           rsMenuSistema!IdPrivAcceso = "TOT"
'           rsMenuSistema.Update
'        End If
'        Set dtgMenuSistema.DataSource = rsMenuSistema
    End If
End Sub

Private Sub dtgMenuSistema_Click()
    lstPrivacceso.Visible = False
End Sub

Private Sub dtgMenuSistema_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
End Sub

Private Sub dtgMenuSistema_Scroll(Cancel As Integer)
    lstPrivacceso.Visible = False
    Cancel = False
End Sub

Private Sub Form_Load()
    'Valores iniciales
    lstPrivacceso.Visible = False
    
    'Verificamos cuantos niveles de acceso existen definidos
    rsNiveles.Open "Select Distinct IdNivelAcceso From NivelAcceso", db, adOpenStatic
    If rsNiveles.RecordCount > 0 Then
        rsNiveles.MoveFirst
        vIdNivelAcceso = rsNiveles!idNivelAcceso
        rsNiveles.MoveLast
        vMaxNivelAcceso = rsNiveles!idNivelAcceso
        rsNiveles.MoveFirst
    Else
        vMaxNivelAcceso = 0
        vIdNivelAcceso = 0
        cmdanterior.Enabled = False
        cmdsiguiente.Enabled = False
        cmdEditar.Enabled = False
    End If
    
    'Abrimos la tabla de privilegios de Operación
    rsPrivAcceso.Open "Select IdPrivAcceso, DesPrivAcceso From PrivilegioAcceso", db, adOpenStatic
    While Not rsPrivAcceso.EOF
        lstPrivacceso.AddItem rsPrivAcceso!IdPrivAcceso
        rsPrivAcceso.MoveNext
    Wend
    rsPrivAcceso.Close
    
    'Abrimos la tabla de niveles de acceso
    rsMenuSistema.Open "Select na.IdNivelAcceso, ms.IdOpcMenu, ms.descopcmenu, ms.EsTerminal, na.Habilitado, na.DesNivelAcceso,na.IdPrivAcceso From Menusistema ms, NivelAcceso na Where ms.NombOpcMenu like na.NombOpcMenu and na.IdNivelAcceso=" & vIdNivelAcceso & " Order by ms.IdOpcMenu", db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    lblNivelAcceso = rsMenuSistema!idNivelAcceso
    lblDesNivelAcceso = rsMenuSistema!DesNivelAcceso
    vEditando = True
    dtgMenuSistema.Columns(2).Locked = True
    dtgMenuSistema.Columns(2).Button = False
    dtgMenuSistema.Columns(3).Locked = True
    dtgMenuSistema.Columns(3).Button = False
    lblDesNivelAcceso.Visible = True
    txtDesNivelAcceso.Visible = False
    BotonesNavegar
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsMenuSistema.State = 1 Then
         rsMenuSistema.MoveFirst
        If rsMenuSistema.EditMode <> adEditNone Then rsMenuSistema.CancelUpdate
        rsMenuSistema.Close
    End If
    If rsNivelAcceso.State = 1 Then rsNivelAcceso.Close
    rsNiveles.Close
End Sub

Private Sub BotonesNavegar()
    cmdanterior.Enabled = True
    cmdsiguiente.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    cmdEliminar.Enabled = True
    CmdGrabar.Enabled = False
    CmdCancelar.Enabled = False
    CmdSalir.Enabled = True
End Sub

Private Sub BotonesConfirma()
    cmdanterior.Enabled = False
    cmdsiguiente.Enabled = False
    cmdNuevo.Enabled = False
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    CmdGrabar.Enabled = True
    CmdCancelar.Enabled = True
    CmdSalir.Enabled = False
End Sub

Private Sub lstPrivacceso_Click()
    lstPrivacceso.Visible = False
    rsMenuSistema!IdPrivAcceso = lstPrivacceso.Text
    rsMenuSistema.Update
    Set dtgMenuSistema.DataSource = rsMenuSistema
End Sub

Private Sub lstPrivacceso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then lstPrivacceso.Visible = False
End Sub

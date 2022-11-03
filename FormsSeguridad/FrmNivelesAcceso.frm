VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmNivelesAcceso 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Nivel de Acceso"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   Icon            =   "FrmNivelesAcceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmNivelesAcceso.frx":0442
   ScaleHeight     =   4740
   ScaleWidth      =   8190
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtnCancelar 
      BackColor       =   &H8000000A&
      Caption         =   "Cancelar"
      Height          =   675
      Left            =   200
      Picture         =   "FrmNivelesAcceso.frx":6C474
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2640
      Width           =   765
   End
   Begin VB.CommandButton BtnSalir 
      BackColor       =   &H8000000A&
      Caption         =   "Cerrar"
      Height          =   720
      Left            =   200
      Picture         =   "FrmNivelesAcceso.frx":6C67E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3480
      Width           =   765
   End
   Begin VB.CommandButton BtnEliminar 
      BackColor       =   &H8000000A&
      Caption         =   "Anular"
      Height          =   720
      Left            =   195
      Picture         =   "FrmNivelesAcceso.frx":6C888
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Anula Registro Activo"
      Top             =   1920
      Width           =   765
   End
   Begin VB.CommandButton BtnModificar 
      BackColor       =   &H8000000A&
      Caption         =   "Modificar"
      Height          =   720
      Left            =   200
      Picture         =   "FrmNivelesAcceso.frx":6D552
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Modifica Registro Activo"
      Top             =   1200
      Width           =   765
   End
   Begin VB.CommandButton BtnAñadir 
      BackColor       =   &H8000000A&
      Caption         =   "Nuevo"
      Height          =   720
      Left            =   200
      Picture         =   "FrmNivelesAcceso.frx":6DB32
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Nuevo Registro"
      Top             =   480
      Width           =   765
   End
   Begin VB.ListBox lstPrivacceso 
      Height          =   645
      ItemData        =   "FrmNivelesAcceso.frx":6E156
      Left            =   5025
      List            =   "FrmNivelesAcceso.frx":6E158
      TabIndex        =   11
      Top             =   1140
      Width           =   2565
   End
   Begin VB.CommandButton cmdSiguiente 
      BackColor       =   &H8000000A&
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
      Left            =   7710
      TabIndex        =   10
      Top             =   60
      Width           =   375
   End
   Begin VB.CommandButton cmdAnterior 
      BackColor       =   &H8000000A&
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
      Left            =   2880
      TabIndex        =   9
      Top             =   60
      Width           =   375
   End
   Begin VB.TextBox txtDesNivelAcceso 
      Height          =   285
      Left            =   3780
      MaxLength       =   15
      TabIndex        =   4
      Top             =   105
      Width           =   3915
   End
   Begin MSDataGridLib.DataGrid dtgMenuSistema 
      Height          =   3735
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   6885
      _ExtentX        =   12144
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
      Caption         =   "MENU DEL SISTEMA"
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
            ColumnWidth     =   4004.788
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Button          =   -1  'True
            Locked          =   -1  'True
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Button          =   -1  'True
            Locked          =   -1  'True
            ColumnWidth     =   780.095
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton BtnGrabar 
      BackColor       =   &H8000000A&
      Caption         =   "Grabar"
      Height          =   675
      Left            =   200
      Picture         =   "FrmNivelesAcceso.frx":6E15A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1680
      Width           =   765
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
      Left            =   2265
      TabIndex        =   8
      Top             =   4395
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "APR=Aprobacion"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5580
      TabIndex        =   7
      Top             =   4395
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "CON=Consuta"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   4260
      TabIndex        =   6
      Top             =   4395
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "TOT=Total"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3165
      TabIndex        =   5
      Top             =   4395
      Width           =   780
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1200
      Top             =   4320
      Width           =   6885
   End
   Begin VB.Label lblDesNivelAcceso 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3780
      TabIndex        =   3
      Top             =   120
      Width           =   3915
   End
   Begin VB.Label lblNivelAcceso 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3270
      TabIndex        =   1
      Top             =   105
      Width           =   530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Perfil de Acceso -->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1695
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
Dim vTipoAcceso As String
Dim Fila As Integer

Private Sub cmdanterior_Click()
    rsNiveles.MovePrevious
    If rsNiveles.BOF Then
       rsNiveles.MoveFirst
    End If
    vIdNivelAcceso = rsNiveles!IdNivelAcceso
    lblNivelAcceso = vIdNivelAcceso
    If rsMenuSistema.State = 1 Then
        If rsMenuSistema.EditMode <> adEditNone Then rsMenuSistema.CancelUpdate
        rsMenuSistema.Close
    End If
    rsMenuSistema.Open "Select na.IdNivelAcceso, ms.menu_descripcion, ms.EsTerminal, na.Habilitado, na.DesNivelAcceso, na.IdPrivAcceso From gc_Menu_sistema ms, gc_nivelacceso na Where ms.NombOpcMenu like na.NombOpcMenu and na.IdNivelAcceso=" & vIdNivelAcceso & " Order by ms.menu_codigo", db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    lblDesNivelAcceso = rsMenuSistema!DesNivelAcceso
End Sub

Private Sub BtnCancelar_Click()
    If rsMenuSistema.State = 1 Then If rsMenuSistema.EditMode <> adEditNone Then rsMenuSistema.CancelUpdate
    rsMenuSistema.Close
    rsMenuSistema.Open "Select na.IdNivelAcceso, ms.menu_descripcion, ms.EsTerminal, na.Habilitado,na.IdPrivAcceso From gc_Menu_sistema ms, gc_nivelacceso na Where ms.NombOpcMenu like na.NombOpcMenu and na.IdNivelAcceso=" & vIdNivelAcceso & " Order by ms.menu_codigo", db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    dtgMenuSistema.Columns(2).Button = False
    dtgMenuSistema.Columns(3).Button = False
    txtDesNivelAcceso.Visible = False
    lblDesNivelAcceso.Visible = True
    lblNivelAcceso = vIdNivelAcceso
    If rsMenuSistema.RecordCount > 0 Then
        BotonesNavegar
    Else
        BotonesInicio
    End If
End Sub

Private Sub BtnModificar_Click()
If rsNiveles.RecordCount > 0 And rsNiveles!IdNivelAcceso <> 0 Then
    If rsNivelAcceso.State = 1 Then rsNivelAcceso.Close
    rsNivelAcceso.Open "Select * From gc_nivelacceso Where IdNivelAcceso=" & vIdNivelAcceso, db, adOpenKeyset, adLockOptimistic
    rsNivelAcceso.MoveFirst
    While Not rsNivelAcceso.EOF
        db.Execute "Update gc_Menu_sistema Set Habilitado = '" & rsNivelAcceso!habilitado & "' Where NombOpcMenu = '" & rsNivelAcceso!NombOpcMenu & "'"
        db.Execute "Update gc_Menu_sistema Set IdPrivAcceso='" & rsNivelAcceso!IdPrivAcceso & "' Where NombOpcMenu = '" & rsNivelAcceso!NombOpcMenu & "'"
        rsNivelAcceso.MoveNext
    Wend
    rsNivelAcceso.MoveFirst
    If rsMenuSistema.State = 1 Then If rsMenuSistema.EditMode <> 0 Then rsMenuSistema.CancelUpdate
    rsMenuSistema.Close
    rsMenuSistema.Open "Select * From gc_Menu_sistema Order by menu_codigo", db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    vEditando = True
    dtgMenuSistema.Columns(2).Button = True
    dtgMenuSistema.Columns(3).Button = True
    txtDesNivelAcceso = lblDesNivelAcceso
    txtDesNivelAcceso.Visible = True
    lblDesNivelAcceso.Visible = False
    BotonesConfirma
End If
End Sub

Private Sub BtnEliminar_Click()
Dim rsUsuarios As New ADODB.Recordset
If rsNiveles.RecordCount > 0 Then
  rsUsuarios.Open "Select * From GC_Usuarios Where IdNivelAcceso=" & CInt(lblNivelAcceso), db, adOpenStatic
  If rsUsuarios.RecordCount > 0 Then
      MsgBox "Existen " & rsUsuarios.RecordCount & " usuarios registrados con este nivel de acceso. Antes de eliminar" & vbCrLf & _
             "el nivel, debe primero cambiar de nivel de acceso de los usuarios afectados.", vbInformation + vbOKOnly, "Atención"
  Else
      If MsgBox("Esta seguro de eliminar este nivel de acceso?", vbCritical + vbYesNo, "Atención") = vbYes Then
         db.Execute "delete from gc_nivelacceso Where IdNivelAcceso=" & CInt(lblNivelAcceso)
         MsgBox "Nivel de acceso eliminado!", vbInformation + vbOKOnly, "Atención"
         rsNiveles.Requery
         'Refrescamos el datagrid de Niveles de acceso
         If rsNiveles.RecordCount > 0 Then
            vIdNivelAcceso = CInt(rsNiveles!IdNivelAcceso)
            lblNivelAcceso = vIdNivelAcceso
            lblDesNivelAcceso = rsNiveles!DesNivelAcceso
            txtDesNivelAcceso = rsNiveles!DesNivelAcceso
            'Actualiza el contador de niveles de acceso
            rsNiveles.MoveLast
            vMaxNivelAcceso = rsNiveles!IdNivelAcceso
            rsNiveles.MoveFirst
            BotonesNavegar
         Else
            lblNivelAcceso = ""
            lblDesNivelAcceso = ""
            txtDesNivelAcceso = ""
            'Actualiza el contador de niveles de acceso
            vMaxNivelAcceso = 0
            BotonesInicio
         End If
         If rsMenuSistema.State = 1 Then
            If rsMenuSistema.EditMode <> adEditNone Then rsMenuSistema.CancelUpdate
               rsMenuSistema.Close
         End If
         rsMenuSistema.Open "Select na.IdNivelAcceso, ms.menu_descripcion, ms.EsTerminal, na.Habilitado, na.DesNivelAcceso, na.IdPrivAcceso From gc_Menu_sistema ms, gc_nivelacceso na Where ms.NombOpcMenu like na.NombOpcMenu and na.IdNivelAcceso=" & vIdNivelAcceso & " Order by ms.menu_codigo", db, adOpenKeyset, adLockOptimistic
         Set dtgMenuSistema.DataSource = rsMenuSistema
      End If
  End If
  rsUsuarios.Close
End If
End Sub

Private Sub BtnGrabar_Click()
    rsMenuSistema.MoveFirst
    While Not rsMenuSistema.EOF
        If Not vEditando Then
            rsNivelAcceso.AddNew
            rsNivelAcceso!IdNivelAcceso = CInt(lblNivelAcceso)
            rsNivelAcceso!NombOpcMenu = rsMenuSistema!NombOpcMenu
            rsNivelAcceso!habilitado = rsMenuSistema!habilitado
            rsNivelAcceso!DesNivelAcceso = txtDesNivelAcceso
            rsNivelAcceso!IdPrivAcceso = rsMenuSistema!IdPrivAcceso
            rsNivelAcceso!EsTerminal = True
            rsNivelAcceso.Update
            
            vMaxNivelAcceso = CInt(lblNivelAcceso)
        Else
            db.Execute "Update gc_nivelacceso Set Habilitado='" & rsMenuSistema!habilitado & "' Where IdNivelAcceso= " & vIdNivelAcceso & " and NombOpcMenu='" & rsMenuSistema!NombOpcMenu & "'"
            db.Execute "Update gc_nivelacceso Set IdPrivAcceso='" & rsMenuSistema!IdPrivAcceso & "' Where IdNivelAcceso= " & vIdNivelAcceso & " and NombOpcMenu='" & rsMenuSistema!NombOpcMenu & "'"
        End If
        rsMenuSistema.MoveNext
    Wend
    rsNiveles.Requery
    rsNiveles.MoveLast
    
    vIdNivelAcceso = CInt(rsNiveles!IdNivelAcceso)
    lblNivelAcceso = vIdNivelAcceso
    lblDesNivelAcceso = rsNiveles!DesNivelAcceso
    txtDesNivelAcceso = rsNiveles!DesNivelAcceso
    
    rsMenuSistema.Close
    rsMenuSistema.Open "Select na.IdNivelAcceso, ms.menu_descripcion, ms.EsTerminal, na.Habilitado,na.IdPrivAcceso From gc_Menu_sistema ms, gc_nivelacceso na Where ms.NombOpcMenu like na.NombOpcMenu and na.IdNivelAcceso=" & vIdNivelAcceso & " Order by ms.menu_codigo", db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    MsgBox "Nivel de acceso grabado satisfactoriamente", vbInformation + vbOKOnly, "Atención"
    dtgMenuSistema.Columns(2).Button = False
    dtgMenuSistema.Columns(3).Button = False
    txtDesNivelAcceso.Visible = False
    lblDesNivelAcceso.Visible = True

    BotonesNavegar
End Sub

Private Sub BtnAñadir_Click()
    'Abrimos la tabla de niveles de acceso
    If rsNivelAcceso.State = 1 Then rsNivelAcceso.Close
    rsNivelAcceso.Open "Select * From gc_nivelacceso", db, adOpenKeyset, adLockOptimistic
    
    'Abrimos la tabla de MenuSistema
    db.Execute "Update gc_Menu_Sistema Set acceso_codigo='Si'"
    If rsMenuSistema.State = 1 Then rsMenuSistema.Close
    rsMenuSistema.Open "Select * From gc_Menu_sistema Order by menu_codigo", db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    lblNivelAcceso = vMaxNivelAcceso + 1
    vEditando = False
    dtgMenuSistema.Columns(2).Button = True
    dtgMenuSistema.Columns(3).Button = True
    txtDesNivelAcceso.Visible = True
    lblDesNivelAcceso.Visible = False
    txtDesNivelAcceso = ""
    BotonesConfirma
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub cmdsiguiente_Click()
    rsNiveles.MoveNext
    If rsNiveles.EOF Then
       rsNiveles.MoveLast
    End If
    vIdNivelAcceso = CInt(rsNiveles!IdNivelAcceso)
    lblNivelAcceso = vIdNivelAcceso
    lblDesNivelAcceso = rsNiveles!DesNivelAcceso
    txtDesNivelAcceso = rsNiveles!DesNivelAcceso
    
    If rsMenuSistema.State = 1 Then
        If rsMenuSistema.EditMode <> adEditNone Then rsMenuSistema.CancelUpdate
        rsMenuSistema.Close
    End If
    rsMenuSistema.Open "Select na.IdNivelAcceso, ms.menu_descripcion, ms.EsTerminal, na.Habilitado, na.DesNivelAcceso, na.IdPrivAcceso From gc_Menu_sistema ms, gc_nivelacceso na Where ms.NombOpcMenu like na.NombOpcMenu and na.IdNivelAcceso=" & vIdNivelAcceso & " Order by ms.menu_codigo", db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    'lblDesNivelAcceso = rsMenuSistema!DesNivelAcceso
End Sub

Private Sub dtgMenuSistema_ButtonClick(ByVal ColIndex As Integer)
On Error Resume Next
    'Permite la habilitacion de las opciones de menu
    'Si No=No tiene acceso; Si=Si tiene acceso
    If ColIndex = 2 Then
        If dtgMenuSistema.Columns(2).Value = "No" Then
           dtgMenuSistema.Columns(2).Value = "Si"
        Else
           dtgMenuSistema.Columns(2).Value = "No"
        End If
        Set dtgMenuSistema.DataSource = rsMenuSistema
    End If
    'Permite el cambio del estado de operacion a los botones de formulario
    'Si TOT=Total; Si CON=Consulta; Si APR=Aprobar
    If ColIndex = 3 Then
        lstPrivacceso.Visible = True
        Fila = dtgMenuSistema.RowBookmark(dtgMenuSistema.Row)
        lstPrivacceso.Top = dtgMenuSistema.RowTop(dtgMenuSistema.Row) + (3 * dtgMenuSistema.RowHeight) + 50
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
    vTipoAcceso = GlTipoAcceso
    'Verificamos cuantos niveles de acceso existen definidos
    rsNiveles.Open "Select Distinct IdNivelAcceso,DesNivelAcceso From gc_nivelacceso Order By IdNivelAcceso", db, adOpenStatic
    If rsNiveles.RecordCount > 0 Then
        rsNiveles.MoveFirst
        vIdNivelAcceso = rsNiveles!IdNivelAcceso
        rsNiveles.MoveLast
        vMaxNivelAcceso = rsNiveles!IdNivelAcceso
        rsNiveles.MoveFirst
    Else
        vMaxNivelAcceso = 0
        vIdNivelAcceso = 0
        cmdAnterior.Enabled = False
        cmdSiguiente.Enabled = False
        BtnModificar.Enabled = False
    End If
    
    'Abrimos la tabla de privilegios de Operación
    rsPrivAcceso.Open "Select IdPrivAcceso, DesPrivAcceso From gc_Privilegio_Acceso", db, adOpenStatic
    While Not rsPrivAcceso.EOF
        lstPrivacceso.AddItem rsPrivAcceso!IdPrivAcceso & "  " & rsPrivAcceso!DesPrivAcceso
        rsPrivAcceso.MoveNext
    Wend
    rsPrivAcceso.Close
    
    'Abrimos la tabla de niveles de acceso
    rsMenuSistema.Open "Select na.IdNivelAcceso, ms.menu_codigo, ms.menu_descripcion, ms.menu_es_terminal, na.Habilitado, na.DesNivelAcceso,na.IdPrivAcceso From gc_Menu_sistema ms,gc_nivelacceso na Where ms.menu_name = na.NombOpcMenu and na.IdNivelAcceso=" & vIdNivelAcceso & " Order by ms.menu_codigo", db, adOpenKeyset, adLockOptimistic
    Set dtgMenuSistema.DataSource = rsMenuSistema
    If rsMenuSistema.RecordCount > 0 Then
        lblNivelAcceso = rsMenuSistema!IdNivelAcceso
        lblDesNivelAcceso = rsMenuSistema!DesNivelAcceso
        vEditando = True
        dtgMenuSistema.Columns(2).Button = False
        dtgMenuSistema.Columns(3).Button = False
        lblDesNivelAcceso.Visible = True
        txtDesNivelAcceso.Visible = False
        BotonesNavegar
    Else
        BotonesInicio
    End If
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsMenuSistema.State = 1 Then
        rsMenuSistema.Close
    End If
    If rsNivelAcceso.State = 1 Then rsNivelAcceso.Close
    rsNiveles.Close
End Sub

Private Sub lstPrivacceso_Click()
    lstPrivacceso.Visible = False
    rsMenuSistema!IdPrivAcceso = Mid(lstPrivacceso.Text, 1, 3)
    rsMenuSistema.Update
    Set dtgMenuSistema.DataSource = rsMenuSistema
End Sub

Private Sub lstPrivacceso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then lstPrivacceso_Click
End Sub

Private Sub BotonesNavegar()
    cmdAnterior.Enabled = True
    cmdSiguiente.Enabled = True
    BtnAñadir.Visible = True
    BtnModificar.Visible = True
    BtnEliminar.Visible = True
    BtnGrabar.Visible = False
    BtnCancelar.Visible = False
    BtnSalir.Visible = True
End Sub

Private Sub BotonesConfirma()
    cmdAnterior.Enabled = False
    cmdSiguiente.Enabled = False
    BtnAñadir.Visible = False
    BtnModificar.Visible = False
    BtnEliminar.Visible = False
    BtnGrabar.Visible = True
    BtnCancelar.Visible = True
    BtnSalir.Visible = False
End Sub

Private Sub BotonesInicio()
    cmdAnterior.Enabled = False
    cmdSiguiente.Enabled = False
    BtnAñadir.Visible = True
    BtnModificar.Visible = False
    BtnEliminar.Visible = False
    BtnGrabar.Visible = False
    BtnCancelar.Visible = False
    BtnSalir.Visible = True
End Sub

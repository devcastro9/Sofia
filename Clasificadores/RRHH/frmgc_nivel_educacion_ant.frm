VERSION 5.00
Begin VB.Form frmgc_nivel_educacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clasificadores - RR.HH. - Nivel de Educacion"
   ClientHeight    =   2970
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmgc_nivel_educacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmgc_nivel_educacion.frx":0A02
   ScaleHeight     =   2970
   ScaleWidth      =   6960
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      BackColor       =   &H00FFC0C0&
      Height          =   660
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   6900
      TabIndex        =   13
      Top             =   0
      Width           =   6960
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Cerrar"
         Height          =   480
         Left            =   5880
         Picture         =   "frmgc_nivel_educacion.frx":C9744
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Salir de Personas"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         Picture         =   "frmgc_nivel_educacion.frx":CA146
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Nuevo Registro"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Modif."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1080
         Picture         =   "frmgc_nivel_educacion.frx":CA6D0
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Aprobar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3000
         Picture         =   "frmgc_nivel_educacion.frx":CAC5A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Aprueba Registro"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "AnuLar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2040
         Picture         =   "frmgc_nivel_educacion.frx":CB1E4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Anula Registro Activo"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Grabar"
         Height          =   480
         Left            =   3960
         Picture         =   "frmgc_nivel_educacion.frx":CBBE6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   480
         Left            =   4920
         Picture         =   "frmgc_nivel_educacion.frx":CC170
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6960
      TabIndex        =   6
      Top             =   2670
      Width           =   6960
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   6585
         Picture         =   "frmgc_nivel_educacion.frx":CC6FA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   6240
         Picture         =   "frmgc_nivel_educacion.frx":CCA3C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmgc_nivel_educacion.frx":CCD7E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmgc_nivel_educacion.frx":CD0C0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Top             =   0
         Width           =   5520
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "estado_codigo"
      Height          =   285
      Index           =   2
      Left            =   2130
      TabIndex        =   5
      Top             =   1890
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nivel_educ_descripcion"
      Height          =   285
      Index           =   1
      Left            =   2130
      TabIndex        =   3
      Top             =   1440
      Width           =   4455
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nivel_educ_codigo"
      Height          =   285
      Index           =   0
      Left            =   2130
      TabIndex        =   1
      Top             =   1005
      Width           =   1095
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIVEL DE EDUCACION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   105
      TabIndex        =   12
      Top             =   105
      Width           =   3105
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado del Registro:"
      Height          =   255
      Index           =   2
      Left            =   210
      TabIndex        =   4
      Top             =   1890
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion Nivel Educ.:"
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo Nivel:"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   1005
      Width           =   1815
   End
End
Attribute VB_Name = "frmgc_nivel_educacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents adoPrimaryRS As Recordset
Dim adoPrimaryRS As Recordset
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'  Dim db As Connection
'  Set db = New Connection
'  db.CursorLocation = adUseClient
'  db.Open "PROVIDER=MSDASQL;dsn=OdbcMalaria;uid=;pwd=;"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select * from rc_nivel_educacional Order by nivel_educ_descripcion", DB, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Enlaza los cuadros de texto con el proveedor de datos
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next

  mbDataChanged = False
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  If adoPrimaryRS!estado_codigo = "N" Then
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Else
    MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
  End If
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  If adoPrimaryRS!estado_codigo = "N" Then
    adoPrimaryRS!estado_codigo = "S"
    adoPrimaryRS!fecha_registro = Date
    adoPrimaryRS!usr_codigo = GlUsuario
  Else
    MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
  End If
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False
  
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveLast
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveFirst
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub


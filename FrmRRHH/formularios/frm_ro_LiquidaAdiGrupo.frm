VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_ro_LiquidaAdiGrupo 
   BackColor       =   &H00000000&
   Caption         =   "Grupo de Planilla"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFF80&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   14
      Top             =   720
      Width           =   2175
   End
   Begin VB.PictureBox fraToolBarGuarda 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   30
      Picture         =   "frm_ro_LiquidaAdiGrupo.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   1035
      TabIndex        =   10
      Top             =   30
      Width           =   1095
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "frm_ro_LiquidaAdiGrupo.frx":6C032
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cancelar"
         Top             =   1320
         Width           =   765
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   120
         Picture         =   "frm_ro_LiquidaAdiGrupo.frx":6C23C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   9900
         TabIndex        =   13
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.TextBox txtCodGrupo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00FFFF80&
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtDesGrupoLiq 
      Height          =   765
      Left            =   2520
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1680
      Width           =   5535
   End
   Begin VB.TextBox txtCodUnidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFF80&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame fraModalidadLiq 
      BackColor       =   &H00404040&
      Caption         =   "Modalidad de Planilla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   855
      Left            =   3960
      TabIndex        =   1
      Top             =   360
      Width           =   4095
      Begin VB.OptionButton optModalidadLiq 
         BackColor       =   &H00404040&
         Caption         =   "Consultor Individual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton optModalidadLiq 
         BackColor       =   &H00404040&
         Caption         =   "Por Planilla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSDataListLib.DataCombo cboUnidad 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cod. Planilla:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblDescrip 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción de Planilla:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   615
      Left            =   1320
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frm_ro_LiquidaAdiGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQLs As String ' usado para la elaboración de los querys
Dim NroCon As Integer ' para guaradar el nuemro de consultoria
Dim ModLiq As String ' para guaradar la modalidad de liquidacion

Private Sub cboUnidad_Change()
    cboUnidad.ToolTipText = cboUnidad.Text
    'asigna el codigo de proyecto a la caja código
    txtCodUnidad.Text = cboUnidad.BoundText

End Sub

Private Sub cboUnidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub cmdGuardar_Click()
    Dim swGuardar As Integer ' usado para saber si efectivamente se almaceno o elimino los datos en la base
                          ' swGuarda -> 0 si se realizo el proceso satisfactoriamente
                          ' swGuarda -> 1 si se produjo un evento de cancelar por parte del usuario en el proceso
                          ' swGuarda -> 2 si se produjo un error de integridad de la base de datos en el servidor por el proceso
    
    Screen.MousePointer = vbHourglass
    If fl_VerificaGrupoLiq Then ' verificamos si la información está correcta antes de actualizar la BD
        
        Call pl_GuardarGrupoLiq(swGuardar)

        If swGuardar = 0 Then   ' si el proceso se realizo satisfactoriamente
            Unload Me
        ElseIf swGuardar = 2 Then ' si se cancelo el proceso por un error controlado por el servidor
            txtDesGrupoLiq.SetFocus
        End If
        
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSalir_Click()
    If vbYes = MsgBox("Desea mostrar los valores originales, perdiendo cualquier modificación realizada?", vbDefaultButton2 + vbYesNo + vbQuestion, "Aviso") Then
        Unload Me
      Else
        txtDesGrupoLiq.SetFocus
    End If

End Sub

Private Sub Form_Load()
    
    Dim NumGrupo As Integer
    Dim rstTemp As ADODB.Recordset ' usado para la carga de los combos de base
    
    ' carga la unidad solicitante
    Set rstTemp = New ADODB.Recordset
    SQLs = "select codigo_unidad, codigo_unidad + ' - '+ uni_descripcion_larga as des_unidad from fc_unidad_ejecutora where uni_activo = 'S' ORDER BY codigo_unidad"
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        Set cboUnidad.RowSource = rstTemp
        cboUnidad.BoundColumn = "codigo_unidad"
        cboUnidad.ListField = "des_unidad"
        
      Else
        MsgBox "El catalogo de unidad solicitante no esta actualizado.", vbInformation, "Aviso"
    End If
    
    'carga los valores usados como parametros
    cboUnidad.BoundText = frm_ro_LiquidaMain.lblCodUniSol.Caption
    txtCodGrupo.Text = Val(frm_ro_LiquidaMain.lblCodGrupo.Caption)
    txtDesGrupoLiq.Text = frm_ro_LiquidaMain.lblDesGrupo.Caption
    
    NroCon = Val(frm_ro_LiquidaMain.lblCodGrupo.Tag)
    ModLiq = frm_ro_LiquidaMain.lblDesGrupo.Tag
    
    If ModLiq = "P" Then
        optModalidadLiq(0).Value = True
        optModalidadLiq(1).Value = False
      Else
        optModalidadLiq(0).Value = False
        optModalidadLiq(1).Value = True
    End If
    Screen.MousePointer = vbDefault
    
	Call SeguridadSet(Me)
End Sub


Private Sub pl_GuardarGrupoLiq(ByRef swGuarda As Integer)
    ' guarda la información en la base de datos
    ' swGuarda -> 0 si se realizo el proceso satisfactoriamente
    ' swGuarda -> 1 si se produjo un evento de cancelar por parte del usuario en el proceso
    ' swGuarda -> 2 si se produjo un error de integridad de la base de datos en el servidor por el proceso

    Dim codigo_grupo_out As Integer
    
    On Error GoTo EtiqError
    
'    Select Case frm_ro_LiquidaMain.lblEstadoGrupoLiq.Caption
'      Case "N" ' se está en modo de adición
''        inser into (ges_gestion, codigo_unidad, codigo_grupo, numero_consultoria, descripcion_grupo, modalidad_pago, estado_aprobado, Usr_Usuario, Fecha_Registro, Hora_Registro)
'        'JQ QR
'        'De.dbo_ap_PagosGrabaGrupo Trim(Str(Year(GlFechaProceso))), cboUnidad.BoundText, 0, NroCon, pg_ReemplazaCarater(pg_QuitaEspBlanco(txtDesGrupoLiq.Text), Chr(34), Chr(39)), IIf(optModalidadLiq(0).Value = True, "P", "I"), "N", "", codigo_grupo_out, GlUsuario
'
'      Case "E" ' se esta editando el registro actual
''        inser into (ges_gestion, codigo_unidad, codigo_grupo, numero_consultoria, descripcion_grupo, modalidad_pago, estado_aprobado, Usr_Usuario, Fecha_Registro, Hora_Registro)
'        'JQ QR
'        'De.dbo_ap_PagosGrabaGrupo frm_ro_LiquidaMain.lblGestion.Caption, cboUnidad.BoundText, Val(txtCodGrupo.Text), NroCon, pg_ReemplazaCarater(pg_QuitaEspBlanco(txtDesGrupoLiq.Text), Chr(34), Chr(39)), IIf(optModalidadLiq(0).Value = True, "P", "I"), "N", cboUnidad.BoundText, codigo_grupo_out, GlUsuario
'
'    End Select
    db.Execute "update ro_pagos_grupos set descripcion_grupo = '" & txtDesGrupoLiq & "' where ges_gestion = '" & frm_ro_LiquidaMain.lblGestion.Caption & "' and planilla_codigo = '" & txtCodGrupo.Text & "' mes_grupo = " & txtCodUnidad & " "
    'ges_gestion, planilla_codigo, mes_grupo, descripcion_grupo, unidad_codigo, depto_codigo, clasif_codigo, doc_codigo, solicitud_tipo, estado_codigo,
'                      usr_codigo , Fecha_Registro, Hora_Registro, usr_aprueba, fecha_aprueba, hora_aprueba

    If codigo_grupo_out > 0 Then
        frm_ro_LiquidaMain.lblEstadoGrupoLiq.Caption = codigo_grupo_out ' se guarda temporalmente el id_contrato para posicionarse el el registro editado o nuevo
        swGuarda = 0 ' si llego a esta parte es porque los cambios se realizaron efectivamente en la base de datos
      Else
        swGuarda = 2
        GoTo EtiqError:
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub
    
EtiqError:
    Select Case Err.Number
      Case -2147217900
        MsgBox "Error: No puede existir dos códigos iguales, corrija el error y vuelva a intentarlo." & Chr(13) & "Los cambios no se llevaron a cabo." & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description, vbCritical, "Error"
        swGuarda = 2
      Case Else ' si se produjo otro tipo de error
        MsgBox "Error: Los cambios no se llevaron a cabo." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
        swGuarda = 2
    End Select
    
End Sub

Private Function fl_VerificaGrupoLiq() As Boolean
    'TITULO:                Función fl_VerificaGrupoLiq
    'PROPOSITO:             Verifica los datos para el registro de grupo de liquidacion
    'EJEMPLO DE LLAMADA:    fl_VerificaGrupoLiq()
    
    fl_VerificaGrupoLiq = True ' asuminos que se cuenta con los datos mnimos para grabar
        
    ' verificamos unidad solicitante
    If Len(RTrim(LTrim(txtCodUnidad.Text))) = 0 Then
        MsgBox "La unidad solicitante no es vália." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        cboUnidad.SetFocus
        fl_VerificaGrupoLiq = False
        Exit Function
    End If
        
    ' verificamos descripcion de grupo
    If Len(RTrim(LTrim(txtDesGrupoLiq.Text))) = 0 Then
        MsgBox "La descripción del grupo de liquidación no es vália." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        txtDesGrupoLiq.SetFocus
        fl_VerificaGrupoLiq = False
        Exit Function
    End If
    
End Function

Private Sub txtCodGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub txtCodUnidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub txtDesGrupoLiq_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
      Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

End Sub

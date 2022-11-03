VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ro_LiquidaAdiPago 
   BackColor       =   &H00000000&
   Caption         =   "RRHH - Pagos por Planilla"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox fraToolBarGuarda 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   0
      Picture         =   "frm_ro_LiquidaAdiPago.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   1035
      TabIndex        =   14
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   120
         Picture         =   "frm_ro_LiquidaAdiPago.frx":6C032
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   480
         Width           =   765
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "frm_ro_LiquidaAdiPago.frx":6C23C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Cancelar"
         Top             =   2280
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
         TabIndex        =   17
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.TextBox tdnNveces 
      Height          =   285
      Left            =   7100
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1160
      Width           =   400
   End
   Begin VB.TextBox txtConcepto 
      Height          =   885
      Left            =   2400
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2160
      Width           =   5655
   End
   Begin VB.TextBox txtAntecedente 
      Height          =   885
      Left            =   2400
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3120
      Width           =   5655
   End
   Begin VB.TextBox txtNroLiq 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   9
      Top             =   1680
      Width           =   855
   End
   Begin VB.Frame fraNveces 
      BackColor       =   &H00404040&
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   960
      Width           =   4095
      Begin MSComCtl2.UpDown udwNveces 
         Height          =   375
         Left            =   3735
         TabIndex        =   8
         Top             =   160
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1
         OrigLeft        =   3735
         OrigTop         =   120
         OrigRight       =   3975
         OrigBottom      =   495
         Max             =   24
         Min             =   1
         Enabled         =   0   'False
      End
      Begin VB.Label lblNveces 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número de veces a repetir registro:"
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
   End
   Begin MSComCtl2.DTPicker tdpFEstimadaLiq 
      Height          =   315
      Left            =   6600
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   58654721
      CurrentDate     =   36882
   End
   Begin VB.Label Label74 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha estimada de liquidación:"
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
      Left            =   3840
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
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
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblGrupo 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2400
      TabIndex        =   11
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto de pago:"
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
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Antece- dentes:"
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
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label53 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Liquidación:"
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
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "frm_ro_LiquidaAdiPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQLs As String ' usado para la elaboración de los querys

Private Sub cmdSalir_Click()
    If vbYes = MsgBox("Desea mostrar los valores originales, perdiendo cualquier modificación realizada?", vbDefaultButton2 + vbYesNo + vbQuestion, "Aviso") Then
        Unload Me
      Else
        TxtConcepto.SetFocus
    End If

End Sub

Private Sub cmdGuardar_Click()
    Dim swGuardar As Integer ' usado para saber si efectivamente se almaceno o elimino los datos en la base
                          ' swGuarda -> 0 si se realizo el proceso satisfactoriamente
                          ' swGuarda -> 1 si se produjo un evento de cancelar por parte del usuario en el proceso
                          ' swGuarda -> 2 si se produjo un error de integridad de la base de datos en el servidor por el proceso
    
    Screen.MousePointer = vbHourglass
    If fl_VerificaLiquida Then ' verificamos si la información está correcta antes de actualizar la BD
        Call pl_GuardarLiquida(swGuardar)

        If swGuardar = 0 Then   ' si el proceso se realizo satisfactoriamente
            Unload Me
        ElseIf swGuardar = 2 Then ' si se cancelo el proceso por un error controlado por el servidor
            TxtConcepto.SetFocus
        End If
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
    Dim NumGrupo As Integer
    Dim rstTemp As ADODB.Recordset ' usado para la carga de los datos
    
    lblGrupo.Caption = "[" & frm_ro_LiquidaMain.lblCodGrupo & "] " & frm_ro_LiquidaMain.lblDesGrupo
    Select Case frm_ro_LiquidaMain.lblEstadoLiquida.Caption
      Case "N"
        ' obtiene el ultimo pago
        SQLs = "SELECT max(numero_pago) as num_pago "
        SQLs = SQLs & "FROM ao_pagos_cronograma "
        SQLs = SQLs & "WHERE estado_aprobado <> 'E' AND estado_devengado <> 'E' AND ges_gestion = '" & frm_ro_LiquidaMain.lblGestion.Caption & "' AND codigo_unidad ='" & frm_ro_LiquidaMain.lblCodUniSol.Caption & "' AND codigo_grupo = " & Val(frm_ro_LiquidaMain.lblCodGrupo.Caption)
        Set rstTemp = New ADODB.Recordset
        rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
        txtNroLiq.Text = IIf(IsNull(rstTemp!num_pago), 0, rstTemp!num_pago) + 1   ' nuevo numero de liquidación
        
        SQLs = "SELECT numero_pago, concepto, tipo_moneda, monto_us, monto_bs, estado_aprobado, estado_devengado, antecedente, codigo_orden, fecha_estimada_liq "
        SQLs = SQLs & "FROM ao_pagos_cronograma "
        SQLs = SQLs & "WHERE estado_aprobado <> 'E' AND estado_devengado <> 'E' AND ges_gestion = '" & frm_ro_LiquidaMain.lblGestion.Caption & "' AND codigo_unidad ='" & frm_ro_LiquidaMain.lblCodUniSol.Caption & "' AND codigo_grupo = " & Val(frm_ro_LiquidaMain.lblCodGrupo.Caption) & " AND numero_pago = " & Val(txtNroLiq.Text) - 1
        Set rstTemp = New ADODB.Recordset
        rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
        lblNveces.Enabled = True
        tdnNveces.Enabled = True
        udwNveces.Enabled = True
        tdnNveces.Text = 1
        If rstTemp.RecordCount > 0 Then
            tdpFEstimadaLiq.Value = Date
            TxtConcepto.Text = rstTemp!Concepto & ""
            txtAntecedente.Text = rstTemp!antecedente & ""
          Else
            tdpFEstimadaLiq.Value = Date
            TxtConcepto.Text = ""
            txtAntecedente.Text = ""
        End If
        
      Case "E"
        
        SQLs = "SELECT numero_pago, concepto, tipo_moneda, monto_us, monto_bs, estado_aprobado, estado_devengado, antecedente, codigo_orden, fecha_estimada_liq "
        SQLs = SQLs & "FROM ao_pagos_cronograma "
        SQLs = SQLs & "WHERE ges_gestion = '" & frm_ro_LiquidaMain.lblGestion.Caption & "' AND codigo_unidad ='" & frm_ro_LiquidaMain.lblCodUniSol.Caption & "' "
        SQLs = SQLs & " AND codigo_grupo =" & Val(frm_ro_LiquidaMain.lblCodGrupo.Caption) & " AND numero_pago= " & Val(frm_ro_LiquidaMain.lblEstadoLiquida.Tag)
        Set rstTemp = New ADODB.Recordset
        rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
        lblNveces.Enabled = False
        tdnNveces.Enabled = False
        udwNveces.Enabled = False
        tdnNveces.Text = 1
        If rstTemp.RecordCount > 0 Then
            txtNroLiq.Text = rstTemp!NUMERO_PAGO & ""
            tdpFEstimadaLiq.Value = IIf(IsNull(rstTemp!fecha_estimada_liq), Date, rstTemp!fecha_estimada_liq)
            TxtConcepto.Text = rstTemp!Concepto & ""
            txtAntecedente.Text = rstTemp!antecedente & ""
          Else
            txtNroLiq.Text = ""
            tdpFEstimadaLiq.Value = Date
            TxtConcepto.Text = ""
            txtAntecedente.Text = ""
        End If
    End Select
    Screen.MousePointer = vbDefault
    
	Call SeguridadSet(Me)
End Sub

Private Function fl_VerificaLiquida() As Boolean
    'TITULO:                Función fl_VerificaLiquida
    'PROPOSITO:             Verifica los datos para el registro de una liquidacion
    'EJEMPLO DE LLAMADA:    fl_VerificaLiquida
    
    fl_VerificaLiquida = True ' asuminos que se cuenta con los datos mnimos para grabar
    
    On Error GoTo EtiqError
     
    ' verificamos concepto de liquidación
    If Len(RTrim(LTrim(TxtConcepto.Text))) = 0 Then
        MsgBox "El concepto de liquidación no es vália." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        TxtConcepto.SetFocus
        fl_VerificaLiquida = False
        Exit Function
    End If
        
    ' verificamos fecha estimada de liquidación
    If tdpFEstimadaLiq.Value = 0 Then
        MsgBox "La fecha estimada de liquidación no es válida." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        tdpFEstimadaLiq.SetFocus
        fl_VerificaLiquida = False
        Exit Function
    End If

    ' verificamos coherencia de fecha estimada

    SQLs = "SELECT max(fecha_estimada_liq) as fecha_estimada_liq "
    SQLs = SQLs & "FROM ao_pagos_cronograma "
    SQLs = SQLs & "WHERE estado_devengado <> 'E' and ges_gestion = '" & frm_ro_LiquidaMain.lblGestion.Caption & "' AND codigo_unidad ='" & frm_ro_LiquidaMain.lblCodUniSol.Caption & "' AND codigo_grupo = " & Val(frm_ro_LiquidaMain.lblCodGrupo.Caption) & " AND fecha_estimada_liq > '" & tdpFEstimadaLiq.Value & "' and numero_pago < " & Val(txtNroLiq.Text)
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        If Not IsNull(rstTemp!fecha_estimada_liq) Then
            MsgBox "La fecha estimada de liquidación [" & tdpFEstimadaLiq.Value & "] no puede ser inferior a la fecha estimada de una anterior liquidación [" & rstTemp!fecha_estimada_liq & "]." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
            tdpFEstimadaLiq.SetFocus
            fl_VerificaLiquida = False
            Exit Function
        End If
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Function

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Function


Private Sub pl_GuardarLiquida(ByRef swGuarda As Integer)
    ' guarda la información en la base de datos
    ' swGuarda -> 0 si se realizo el proceso satisfactoriamente
    ' swGuarda -> 1 si se produjo un evento de cancelar por parte del usuario en el proceso
    ' swGuarda -> 2 si se produjo un error de integridad de la base de datos en el servidor por el proceso

    Dim i As Integer
    Dim numero_pago_out As Integer
    
    On Error GoTo EtiqError
    
    Select Case frm_ro_LiquidaMain.lblEstadoLiquida.Caption
      Case "N" ' se está en modo de adición
        
        For i = 1 To tdnNveces.Text ' numero de adiciones de numero de liquidaciones
            'JQ QR
            'De.dbo_ap_PagosGrabaPago frm_ro_LiquidaMain.lblGestion.Caption, frm_ro_LiquidaMain.lblCodUniSol.Caption, Val(frm_ro_LiquidaMain.lblCodGrupo.Caption), 0, pg_ReemplazaCarater(pg_QuitaEspBlanco(txtConcepto.Text), Chr(34), Chr(39)), pg_ReemplazaCarater(pg_QuitaEspBlanco(txtAntecedente.Text), Chr(34), Chr(39)), CDate(tdpFEstimadaLiq.Value), GlUsuario, numero_pago_out
        Next i
        
      Case "E" ' se esta editando el registro actual
        'JQ QR
        'De.dbo_ap_PagosGrabaPago frm_ro_LiquidaMain.lblGestion.Caption, frm_ro_LiquidaMain.lblCodUniSol.Caption, Val(frm_ro_LiquidaMain.lblCodGrupo.Caption), Val(txtNroLiq.Text), pg_ReemplazaCarater(pg_QuitaEspBlanco(txtConcepto.Text), Chr(34), Chr(39)), pg_ReemplazaCarater(pg_QuitaEspBlanco(txtAntecedente.Text), Chr(34), Chr(39)), CDate(tdpFEstimadaLiq.Value), GlUsuario, numero_pago_out
        
    End Select
    
    If numero_pago_out > 0 Then
        frm_ro_LiquidaMain.lblEstadoLiquida.Caption = numero_pago_out ' se guarda temporalmente el id_contrato para posicionarse el el registro editado o nuevo
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

Private Sub tdpFEstimadaLiq_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case 13 ' si presiono enter
        SendKeys "{Tab}"
      Case 27 ' si presiono escape
        Call cmdSalir_Click
        
    End Select

End Sub

Private Sub tdnNveces_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub txtAntecedente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
      Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
      Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

End Sub

Private Sub txtNroLiq_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

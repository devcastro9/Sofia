VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frm_ro_LiquidaConformidad 
   Caption         =   "Consultoría: Registra conformidad"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFactura 
      Height          =   975
      Left            =   3240
      TabIndex        =   10
      Top             =   2160
      Width           =   3015
      Begin VB.OptionButton optFactura 
         Caption         =   "CON retención (NO emite factura)"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   12
         Top             =   360
         Width           =   2775
      End
      Begin VB.OptionButton optFactura 
         Caption         =   "SIN retención (SI emite factura)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame fraConformidad 
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
      Begin VB.TextBox txtNroCITE 
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Text            =   "S/N"
         Top             =   600
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker tddFCITE 
         Height          =   315
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   92078081
         CurrentDate     =   36882
      End
      Begin VB.Label Label2 
         Caption         =   "Nro. CITE conformidad:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha CITE conformidad:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame fraPlanilla 
      Height          =   5535
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   6255
      Begin VB.OptionButton optConformidad 
         Caption         =   "Sin conformidad"
         Height          =   195
         Index           =   1
         Left            =   4320
         TabIndex        =   18
         Top             =   1440
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optConformidad 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   4320
         TabIndex        =   17
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   840
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Sale del formulario"
         Top             =   3360
         Width           =   765
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Grabar"
         Height          =   840
         Left            =   240
         MousePointer    =   4  'Icon
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2400
         Width           =   765
      End
      Begin VB.ListBox lstBeneficiario 
         Height          =   3660
         Left            =   1200
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   1680
         Width           =   4935
      End
      Begin VB.Label Label27 
         Caption         =   "Descripción beneficiario:"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   1440
         Width           =   3135
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Número de Liquidación:"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblNroLiq 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblNroLiq"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblDesGrupo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDesGrupo"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label lblDesUnidad 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDesUnidad"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "Descripción Grupo:"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Unidad Grupo:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "REGISTRA COFORMIDAD PARA LA LIQUIDACIÓN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frm_ro_LiquidaConformidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQLs As String ' usado para la elaboración de los querys
Dim gp_ges_gestion As String
Dim gp_codigo_unidad As String
Dim gp_codigo_grupo As Integer
Dim gp_numero_pago As Integer
Dim DesUnidad  As String
Dim DesGrupo As String
Dim TipoFrame As String
Public FechaControl As String ' para controlar el registro de fechas
Public HoraControl As String ' para controlar el registro de fechas

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    ' actualiza la tabla ao_pago_cronograma_detalle
    Dim i As Integer
    
    On Error GoTo EtiqError
    
    Select Case TipoFrame ' tipo de formulario
      Case "C"
        If fl_VerificaDatos Then
            If vbYes = MsgBox("Desea registrar y aprobar los CITES de beneficiarios seleccionados?", vbDefaultButton2 + vbYesNo + vbQuestion, "Aviso") Then
                If lstBeneficiario.Selected(0) = True Then ' selecciono a todos
                    For i = 1 To lstBeneficiario.ListCount - 1 ' para seleccionar a los beneficiarios
                        SQLs = "UPDATE ao_pagos_cronograma_detalle SET ncite_conformidad ='" & pg_ReemplazaCarater(pg_QuitaEspBlanco(txtNroCITE.Text), Chr(34), Chr(39)) & "', "
                        SQLs = SQLs & "fcite_conformidad = '" & tddFCITE.Value & "', estado_conformidad = 'S' "
                        SQLs = SQLs & "WHERE codigo_beneficiario = '" & fl_ObtieneCodBen(lstBeneficiario.List(i)) & "' AND ges_gestion = '" & gp_ges_gestion & "' AND codigo_unidad = '" & gp_codigo_unidad & "' AND numero_pago = " & gp_numero_pago
                        'JQ QR
                        'DE.dbo_apGeneralSearching SQLs
                    Next
                  Else ' algunos
                    For i = 1 To lstBeneficiario.ListCount - 1 ' para seleccionar a los beneficiarios
                        If lstBeneficiario.Selected(i) = True Then
                            SQLs = "UPDATE ao_pagos_cronograma_detalle SET ncite_conformidad ='" & pg_ReemplazaCarater(pg_QuitaEspBlanco(txtNroCITE.Text), Chr(34), Chr(39)) & "', "
                            SQLs = SQLs & "fcite_conformidad = '" & tddFCITE.Value & "', estado_conformidad = 'S' "
                            SQLs = SQLs & "WHERE codigo_beneficiario = '" & fl_ObtieneCodBen(lstBeneficiario.List(i)) & "' AND ges_gestion = '" & gp_ges_gestion & "' AND codigo_unidad = '" & gp_codigo_unidad & "' AND numero_pago = " & gp_numero_pago
                            'JQ QR
                            'DE.dbo_apGeneralSearching SQLs
                        End If
                    Next
                End If
                
                Call pl_ListaBeneficiario
                
            End If
        End If
        
      Case "F"

        If vbYes = MsgBox("Desea registrar que los beneficiarios seleccionados " & IIf(optFactura(0).Value = True, "SI", "NO") & " emiten factura?", vbDefaultButton2 + vbYesNo + vbQuestion, "Aviso") Then
            Select Case True
            
              Case optFactura(0).Value  ' SI emite factira
                  
                If lstBeneficiario.Selected(0) = True Then
                    For i = 1 To lstBeneficiario.ListCount - 1 ' para seleccionar a los beneficiarios
                        SQLs = "UPDATE ao_pagos_cronograma_detalle SET emite_factura ='S' "
                        SQLs = SQLs & "WHERE codigo_beneficiario = '" & fl_ObtieneCodBen(lstBeneficiario.List(i)) & "' AND ges_gestion = '" & gp_ges_gestion & "' AND codigo_unidad = '" & gp_codigo_unidad & "' AND numero_pago = " & gp_numero_pago
                        'JQ QR
                        'DE.dbo_apGeneralSearching SQLs
                    Next
                  Else
                    For i = 1 To lstBeneficiario.ListCount - 1 ' para seleccionar a los beneficiarios
                        If lstBeneficiario.Selected(i) = True Then
                            SQLs = "UPDATE ao_pagos_cronograma_detalle SET emite_factura ='S' "
                            SQLs = SQLs & "WHERE codigo_beneficiario = '" & fl_ObtieneCodBen(lstBeneficiario.List(i)) & "' AND ges_gestion = '" & gp_ges_gestion & "' AND codigo_unidad = '" & gp_codigo_unidad & "' AND numero_pago = " & gp_numero_pago
                            'JQ QR
                            'DE.dbo_apGeneralSearching SQLs
                          Else
                            SQLs = "UPDATE ao_pagos_cronograma_detalle SET emite_factura ='N' "
                            SQLs = SQLs & "WHERE codigo_beneficiario = '" & fl_ObtieneCodBen(lstBeneficiario.List(i)) & "' AND ges_gestion = '" & gp_ges_gestion & "' AND codigo_unidad = '" & gp_codigo_unidad & "' AND numero_pago = " & gp_numero_pago
                            'JQ QR
                            'DE.dbo_apGeneralSearching SQLs
                        End If
                    Next
                End If
                
              Case optFactura(1).Value  ' NO emite factira
                
                If lstBeneficiario.Selected(0) = True Then
                    For i = 1 To lstBeneficiario.ListCount - 1 ' para seleccionar a los beneficiarios
                        SQLs = "UPDATE ao_pagos_cronograma_detalle SET emite_factura ='N' "
                        SQLs = SQLs & "WHERE codigo_beneficiario = '" & fl_ObtieneCodBen(lstBeneficiario.List(i)) & "' AND ges_gestion = '" & gp_ges_gestion & "' AND codigo_unidad = '" & gp_codigo_unidad & "' AND numero_pago = " & gp_numero_pago
                        'JQ QR
                        'DE.dbo_apGeneralSearching SQLs
                    Next
                  Else
                    For i = 1 To lstBeneficiario.ListCount - 1 ' para seleccionar a los beneficiarios
                        If lstBeneficiario.Selected(i) = True Then
                            SQLs = "UPDATE ao_pagos_cronograma_detalle SET emite_factura ='N' "
                            SQLs = SQLs & "WHERE codigo_beneficiario = '" & fl_ObtieneCodBen(lstBeneficiario.List(i)) & "' AND ges_gestion = '" & gp_ges_gestion & "' AND codigo_unidad = '" & gp_codigo_unidad & "' AND numero_pago = " & gp_numero_pago
                            'JQ QR
                            'DE.dbo_apGeneralSearching SQLs
                          Else
                            SQLs = "UPDATE ao_pagos_cronograma_detalle SET emite_factura ='S' "
                            SQLs = SQLs & "WHERE codigo_beneficiario = '" & fl_ObtieneCodBen(lstBeneficiario.List(i)) & "' AND ges_gestion = '" & gp_ges_gestion & "' AND codigo_unidad = '" & gp_codigo_unidad & "' AND numero_pago = " & gp_numero_pago
                            'JQ QR
                            'DE.dbo_apGeneralSearching SQLs
                        End If
                    Next
                End If
              
            End Select
            Call pl_ListaBeneficiario
            
        End If
    End Select
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub Form_Load()

    On Error GoTo EtiqError
   'CAPTURA FECHA DEL SERVIDOR
    'JQ QR
    'DE.dbo_ap_GetServDateTime FechaControl, HoraControl

    gp_ges_gestion = frm_ro_LiquidaMain.lblGestion.Caption ' pago gestion
    gp_codigo_unidad = frm_ro_LiquidaMain.lblCodUniSol.Caption  ' copdigo unidad
    gp_codigo_grupo = Val(frm_ro_LiquidaMain.lblCodGrupo.Caption) ' codigo grupo
    gp_numero_pago = Val(frm_ro_LiquidaMain.grdBeneficiario.Tag)  ' numero liquidación
    DesUnidad = frm_ro_LiquidaMain.lblDesUniSol ' unidad
    DesGrupo = frm_ro_LiquidaMain.lblDesGrupo ' grupo
    lblDesUnidad.Caption = gp_codigo_unidad & " - " & DesUnidad
    lblDesGrupo.Caption = gp_codigo_grupo & " - " & DesGrupo
    lblNroLiq.Caption = gp_numero_pago
    TipoFrame = frm_ro_LiquidaMain.lblEstadoBeneficiario.Caption     ' tipo de frame C=conformidad ó F=emite factura

    Call pl_ListaBeneficiario

    If TipoFrame = "C" Then ' tipo de formulario
        Me.Caption = "Registra conformidad"
        lblTitulo.Caption = "REGISTRA COFORMIDAD PARA LA LIQUIDACIÓN"
        fraConformidad.Top = 2160
        fraConformidad.Left = 240
        fraConformidad.Width = 6015
        fraConformidad.Visible = True
        fraFactura.Visible = False
        optConformidad(0).Visible = True
        optConformidad(1).Visible = True
        
      Else
        Me.Caption = "Registra emite factura"
        lblTitulo.Caption = "REGISTRA SI O NO EMITE FACTURA BENEFICIARIO"
        fraFactura.Top = 2160
        fraFactura.Left = 240
        fraFactura.Width = 6015
        fraConformidad.Visible = False
        fraFactura.Visible = True
        optConformidad(0).Visible = False
        optConformidad(1).Visible = False
    End If
    Screen.MousePointer = vbDefault
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
    
End Sub

Private Sub pl_ListaBeneficiario()
    'TITULO:                Procedimiento pl_listaBeneficirio
    'PROPOSITO:             ' obtiene datos de beneficiarios sin registro de conformidad
    'EJEMPLO DE LLAMADA:    call pl_listaBeneficirio
    
    On Error GoTo EtiqError

    Select Case TipoFrame
      Case "C" ' cite de coformidad
    
        tddFCITE.Value = Date
        txtNroCITE.Text = "S/N"
        
        lstBeneficiario.Clear
        
        Select Case glProceso
          Case "F05"
            SQLs = "SELECT '[' + ao_pagos_cronograma_detalle.codigo_beneficiario + '] ' + fc_beneficiario.nombres_beneficiario + fc_beneficiario.paterno_beneficiario + ' ' + fc_beneficiario.materno_beneficiario + ' (' + ncite_conformidad + '-' + case when fcite_conformidad is null then '' else CONVERT ( varchar, fcite_conformidad, 3) end   + ')' AS des_beneficiario, estado_conformidad "
            SQLs = SQLs & "FROM ao_pagos_cronograma_detalle INNER JOIN fc_beneficiario ON ao_pagos_cronograma_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario "
            SQLs = SQLs & "WHERE ao_pagos_cronograma_detalle.ges_gestion = '" & gp_ges_gestion & "' AND ao_pagos_cronograma_detalle.codigo_unidad = '" & gp_codigo_unidad & "' AND ao_pagos_cronograma_detalle.codigo_grupo = " & gp_codigo_grupo & " AND ao_pagos_cronograma_detalle.numero_pago = " & gp_numero_pago
            SQLs = SQLs & " ORDER BY fc_beneficiario.paterno_beneficiario, fc_beneficiario.materno_beneficiario, fc_beneficiario.nombres_beneficiario"
          Case "F10"
            SQLs = "SELECT '[' + ao_pagos_cronograma_detalle.codigo_beneficiario + '] ' + RC_Personal.nombres + RC_Personal.paterno + ' ' + RC_Personal.materno + ' (' + ncite_conformidad + '-' + case when fcite_conformidad is null then '' else CONVERT ( varchar, fcite_conformidad, 3) end   + ')' AS des_beneficiario, estado_conformidad "
            SQLs = SQLs & "FROM ao_pagos_cronograma_detalle INNER JOIN RC_Personal ON ao_pagos_cronograma_detalle.codigo_beneficiario = RC_Personal.ci "
            SQLs = SQLs & "WHERE ao_pagos_cronograma_detalle.ges_gestion = '" & gp_ges_gestion & "' AND ao_pagos_cronograma_detalle.codigo_unidad = '" & gp_codigo_unidad & "' AND ao_pagos_cronograma_detalle.codigo_grupo = " & gp_codigo_grupo & " AND ao_pagos_cronograma_detalle.numero_pago = " & gp_numero_pago
            SQLs = SQLs & "ORDER BY RC_Personal.paterno, RC_Personal.materno, RC_Personal.nombres"
        End Select
        
        Set rstTemp = New ADODB.Recordset
        rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
        optConformidad(1).Value = True
        Call optConformidad_Click(1)
        
      Case "F" ' factura
      
        optFactura(0).Value = True ' por defecto SI emite factura
        optFactura(1).Value = False
        
        lstBeneficiario.Clear
        Select Case glProceso
          Case "F05"
            SQLs = "SELECT '[' + ao_pagos_cronograma_detalle.codigo_beneficiario + '] ' + fc_beneficiario.nombres_beneficiario + fc_beneficiario.paterno_beneficiario + ' ' + fc_beneficiario.materno_beneficiario + ' [' + case when ao_pagos_cronograma_detalle.emite_factura ='S' then 'Si' else 'No' end + ']' AS des_beneficiario, ao_pagos_cronograma_detalle.emite_factura "
            SQLs = SQLs & "FROM ao_pagos_cronograma_detalle INNER JOIN fc_beneficiario ON ao_pagos_cronograma_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario "
            SQLs = SQLs & "WHERE ao_pagos_cronograma_detalle.ges_gestion = '" & gp_ges_gestion & "' AND ao_pagos_cronograma_detalle.codigo_unidad = '" & gp_codigo_unidad & "' AND ao_pagos_cronograma_detalle.codigo_grupo = " & gp_codigo_grupo & " AND ao_pagos_cronograma_detalle.numero_pago = " & gp_numero_pago
            SQLs = SQLs & " ORDER BY fc_beneficiario.paterno_beneficiario, fc_beneficiario.materno_beneficiario, fc_beneficiario.nombres_beneficiario"
          Case "F10"
            SQLs = "SELECT '[' + ao_pagos_cronograma_detalle.codigo_beneficiario + '] ' + RC_Personal.nombres + RC_Personal.paterno + ' ' + RC_Personal.materno + ' [' + case when ao_pagos_cronograma_detalle.emite_factura ='S' then 'Si' else 'No' end + ']' AS des_beneficiario, ao_pagos_cronograma_detalle.emite_factura "
            SQLs = SQLs & "FROM ao_pagos_cronograma_detalle INNER JOIN RC_Personal ON ao_pagos_cronograma_detalle.codigo_beneficiario = RC_Personal.ci "
            SQLs = SQLs & "WHERE ao_pagos_cronograma_detalle.ges_gestion = '" & gp_ges_gestion & "' AND ao_pagos_cronograma_detalle.codigo_unidad = '" & gp_codigo_unidad & "' AND ao_pagos_cronograma_detalle.codigo_grupo = " & gp_codigo_grupo & " AND ao_pagos_cronograma_detalle.numero_pago = " & gp_numero_pago
            SQLs = SQLs & "ORDER BY RC_Personal.paterno, RC_Personal.materno, RC_Personal.nombres"
        End Select
        
        ' carga datos de beneficiarios
        Set rstTemp = New ADODB.Recordset
        rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
        If rstTemp.RecordCount > 0 Then ' si devolvio por lo menos un registro
            lstBeneficiario.AddItem "Todos"
            lstBeneficiario.ItemData(lstBeneficiario.NewIndex) = 999 ' para filtrar todos
            While Not rstTemp.EOF
                lstBeneficiario.AddItem rstTemp!des_beneficiario
     '''        lstBeneficiario.ItemData(lstBeneficiario.NewIndex) = rstTemp!codigo_beneficiario ' solo tipo integer
                If rstTemp!emite_factura = "S" Then
                    lstBeneficiario.Selected(lstBeneficiario.NewIndex) = True
                  Else
                    lstBeneficiario.Selected(lstBeneficiario.NewIndex) = False
                End If
                rstTemp.MoveNext
            Wend
            optFactura(0).Enabled = True
            optFactura(1).Enabled = True
          
          Else
            optFactura(0).Enabled = False
            optFactura(1).Enabled = False
            cmdGuardar.Enabled = False
            MsgBox "No existe registro de beneficiarios.", vbInformation, "Aviso"
        End If
    End Select

    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub pl_CargaListaBen()
    
    ' carga datos de beneficiarios a la lista
    If rstTemp.RecordCount > 0 Then ' si devolvio por lo menos un registro
        lstBeneficiario.AddItem "Seleccionar todos"
        lstBeneficiario.ItemData(lstBeneficiario.NewIndex) = 999 ' para filtrar todos
        lstBeneficiario.Selected(lstBeneficiario.NewIndex) = True
        While Not rstTemp.EOF
            lstBeneficiario.AddItem rstTemp!des_beneficiario
    '''        lstBeneficiario.ItemData(lstBeneficiario.NewIndex) = rstTemp!codigo_beneficiario ' solo tipo integer
            rstTemp.MoveNext
        Wend
        tddFCITE.Enabled = True
        txtNroCITE.Locked = False
        cmdGuardar.Enabled = True
      Else
        tddFCITE.Enabled = False
        txtNroCITE.Locked = True
        cmdGuardar.Enabled = False
''        MsgBox "No existe registro de beneficiarios sin registro de conformidad.", vbInformation, "Aviso"
    End If

End Sub

Private Sub lstBeneficiario_ItemCheck(Item As Integer)
    Dim i As Integer
    If Item = 0 Then
        If lstBeneficiario.Selected(0) = True Then
          For i = 1 To lstBeneficiario.ListCount - 1
            lstBeneficiario.Selected(i) = False
          Next
        End If
    Else
        lstBeneficiario.Selected(0) = False
    End If
    lstBeneficiario.ToolTipText = lstBeneficiario.List(Item)

End Sub

Private Function fl_ObtieneCodBen(Item As String) As String
    'TITULO:                Función fl_ObtieneCodBen
    'PROPOSITO:             Obtiene el código a partir de la cadena item, se supone q todos los item contienen su código
    '                       para cadenas donde el código se encuentra al principio entre "[]", es decir del tipo "[Código]- descripción"
    'EJEMPLO DE LLAMADA:    fl_ObtieneCodBen( item )
    Dim i As Integer
    i = 1
    While Mid(Item, i, 1) <> "]" And i <= Len(Item)
        i = i + 1
    Wend
    fl_ObtieneCodBen = Mid(Item, 2, i - 2)
End Function
    
Private Function fl_VerificaDatos() As Boolean
    'TITULO:                Función fl_VerificaDatos
    'PROPOSITO:             Verifica los datos para el registro
    'EJEMPLO DE LLAMADA:    fl_VerificaDatos
    Dim i As Integer
    
    fl_VerificaDatos = True ' asuminos que se cuenta con los datos mnimos para grabar
        
    ' verificamos fecha cite
    If tddFCITE.Value = 0 Then
        MsgBox "La fecha de CITE de conformidad no es válida." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        tddFCITE.SetFocus
        fl_VerificaDatos = False
        Exit Function
    End If
    
    ' verificamos coherencia mínima de la fecha
    If tddFCITE.Value > CDate(FechaControl) Then
        MsgBox "La fecha de conformidad [" & tddFCITE.Value & "] no puede ser mayor a la actual [" & FechaControl & "]." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        tddFCITE.SetFocus
        fl_VerificaDatos = False
        Exit Function
    End If
    
    ' verificamos nro. cite
    If Len(RTrim(LTrim(txtNroCITE.Text))) = 0 Then
        MsgBox "El numero de CITE no es vália." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        txtNroCITE.SetFocus
        fl_VerificaDatos = False
        Exit Function
    End If

    ' verificamos eleccionde beneficiarios
    For i = 0 To lstBeneficiario.ListCount - 1
        If lstBeneficiario.Selected(i) = True Then
            Exit For
        End If
    Next i
    
    If i > lstBeneficiario.ListCount - 1 Then
        MsgBox "No se selecciono ningún beneficiario." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        lstBeneficiario.SetFocus
        fl_VerificaDatos = False
        Exit Function
    End If

End Function

Private Sub lstBeneficiario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub optConformidad_Click(Index As Integer)
    lstBeneficiario.Clear
    If Index = 0 Then
        rstTemp.Filter = 0
        Call pl_CargaListaBen
      Else
        rstTemp.Filter = "estado_conformidad like 'N'"
        Call pl_CargaListaBen
    End If

End Sub

Private Sub tddFCITE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub txtNroCITE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
      Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form af_PagosPrintOrdenPago 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultoría - Impresión de Ordenes de Liquidación"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCronograma 
      Caption         =   "Print &Cronograma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   12
      Top             =   6480
      Width           =   1935
   End
   Begin VB.ListBox lstDocAdjunto 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Columns         =   2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4890
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   480
      Width           =   10815
   End
   Begin Crystal.CrystalReport CR 
      Left            =   480
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCancelBtn=   0   'False
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Print Orden Liquidación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   1
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   0
      Top             =   7320
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpFOrdenPago 
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   5880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      Format          =   54722561
      CurrentDate     =   36882
   End
   Begin MSDataListLib.DataCombo cboDeQuien 
      Height          =   315
      Left            =   4080
      TabIndex        =   6
      Top             =   6480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboAQuien 
      Height          =   315
      Left            =   4080
      TabIndex        =   7
      Top             =   7080
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboPuestoDeQuien 
      Height          =   315
      Left            =   2160
      TabIndex        =   10
      Top             =   6480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboPuestoAQuien 
      Height          =   315
      Left            =   2160
      TabIndex        =   11
      Top             =   7080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
      Text            =   ""
   End
   Begin Crystal.CrystalReport CRCrono 
      Left            =   480
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCancelBtn=   0   'False
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   9
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "De:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Documentos Adjuntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Orden Liquidación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      Top             =   5880
      Width           =   3015
   End
End
Attribute VB_Name = "af_PagosPrintOrdenPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQLs As String ' usado para la elaboración de los querys
Dim DocAdj As String ' usado para la lista de documentos adjuntos
Dim CodGestion As String
Dim CodUnidad As String
Dim CodGrupo As Integer
Dim NumPago As Integer
Dim rs_DeQuien As New ADODB.Recordset
Dim rs_AQuien As New ADODB.Recordset
    
Private Sub cboAQuien_Change()
    On Error GoTo EtiqError
    
    cboAQuien.ToolTipText = cboAQuien.Text
    If cboPuestoAQuien.MatchedWithList = False Then
        cboPuestoAQuien.BoundText = rs_AQuien!codigo_puesto
    End If
    If rs_AQuien.RecordCount > 0 Then
        rs_AQuien.MoveFirst
        rs_AQuien.Find "ci = '" & cboAQuien.BoundText & "'"
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub cboDeQuien_Change()
    On Error GoTo EtiqError
    
    cboDeQuien.ToolTipText = cboDeQuien.Text
    If cboPuestoDeQuien.MatchedWithList = False Then
        cboPuestoDeQuien.BoundText = rs_DeQuien!codigo_puesto
    End If
    If rs_DeQuien.RecordCount > 0 Then
        rs_DeQuien.MoveFirst
        rs_DeQuien.Find "ci = '" & cboDeQuien.BoundText & "'"
        
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub cboPuestoAQuien_Change()
    On Error GoTo EtiqError
    
    Dim aux As String
    cboPuestoAQuien.ToolTipText = cboPuestoAQuien.Text
    If rs_AQuien.RecordCount > 0 Then
        If cboAQuien.MatchedWithList = True And rs_AQuien!codigo_puesto = cboPuestoAQuien.BoundText Then
            aux = cboAQuien.BoundText
         Else
            aux = ""
        End If
    End If
    
    SQLs = "select * from ac_unidad_respondable_para_rep where ci not in ('" & cboDeQuien.BoundText & "') and ges_gestion = '" & CodGestion & "' and activo='S' and da in('DGF','" & GldaCodigo & "') and codigo_puesto = " & Val(cboPuestoAQuien.BoundText) & " ORDER BY codigo_puesto DESC"
    Set rs_AQuien = New ADODB.Recordset
    rs_AQuien.Open SQLs, db, adOpenStatic, adLockReadOnly
    Set cboAQuien.RowSource = rs_AQuien
    cboAQuien.BoundColumn = "ci"
    cboAQuien.ListField = "des_responsable"
    cboAQuien.BoundText = ""
    cboAQuien.Refresh
    If Len(Trim(aux)) = 0 And rs_AQuien.RecordCount > 0 Then
        rs_AQuien.MoveFirst
        cboAQuien.BoundText = rs_AQuien!ci
      Else
        cboAQuien.BoundText = aux
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
    
End Sub

Private Sub cboPuestoAQuien_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EtiqError
    
    Select Case KeyCode
      Case 46 ' si presiono suprimir
'''        cboPuestoAQuien.BoundText = ""
'''        SQLs = "select * from ac_unidad_respondable_para_rep where ges_gestion = '" & CodGestion & "' and activo='S' and da in('DGF','" & GldaCodigo & "')"
'''        Set rs_AQuien = New ADODB.Recordset
'''        rs_AQuien.Open SQLs, db, adOpenStatic, adLockReadOnly
'''        If rstTemp.RecordCount > 0 Then
'''            Set cboAQuien.RowSource = rs_AQuien
'''            cboAQuien.BoundColumn = "ci"
'''            cboAQuien.ListField = "des_responsable"
'''            cboAQuien.Refresh
'''            cboAQuien.BoundText = ""
'''          Else
'''            MsgBox "El catalogo de responsables esta vació", vbInformation, "Aviso"
'''        End If
        
      Case 27 ' esc
        Call cmdSalir_Click
      Case 13 ' si presiono enter
        SendKeys "{Tab}"
    End Select
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub cboPuestoDeQuien_Change()
        
    Dim aux As String
    
    On Error GoTo EtiqError
    
    cboPuestoDeQuien.ToolTipText = cboPuestoDeQuien.Text
    If rs_DeQuien.RecordCount > 0 Then
        If cboDeQuien.MatchedWithList = True And rs_DeQuien!codigo_puesto = cboPuestoDeQuien.BoundText Then
            aux = cboDeQuien.BoundText
         Else
            aux = ""
        End If
    End If
    
    SQLs = "select * from ac_unidad_respondable_para_rep where ges_gestion = '" & CodGestion & "' and activo='S' and da in('" & GldaCodigo & "') and codigo_puesto = " & Val(cboPuestoDeQuien.BoundText) & " ORDER BY codigo_puesto DESC"
    Set rs_DeQuien = New ADODB.Recordset
    rs_DeQuien.Open SQLs, db, adOpenStatic, adLockReadOnly
    Set cboDeQuien.RowSource = rs_DeQuien
    cboDeQuien.BoundColumn = "ci"
    cboDeQuien.ListField = "des_responsable"
    cboDeQuien.Refresh
    cboDeQuien.BoundText = ""
    If Len(Trim(aux)) = 0 And rs_DeQuien.RecordCount > 0 Then
        rs_DeQuien.MoveFirst
        cboDeQuien.BoundText = rs_DeQuien!ci
      Else
        cboDeQuien.BoundText = aux
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub cboPuestoDeQuien_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EtiqError
    
    Select Case KeyCode
      Case 46 ' si presiono suprimir
'''        cboPuestoDeQuien.BoundText = ""
'''        SQLs = "select * from ac_unidad_respondable_para_rep where ges_gestion = '" & CodGestion & "' and activo='S' and da in('" & GldaCodigo & "')"
'''        Set rs_DeQuien = New ADODB.Recordset
'''        rs_DeQuien.Open SQLs, db, adOpenStatic, adLockReadOnly
'''        If rs_DeQuien.RecordCount > 0 Then
'''            Set cboDeQuien.RowSource = rs_DeQuien
'''            cboDeQuien.BoundColumn = "ci"
'''            cboDeQuien.ListField = "des_responsable"
'''            cboDeQuien.BoundText = ""
'''            cboDeQuien.Refresh
'''          Else
'''            MsgBox "El catalogo de responsables esta vació", vbInformation, "Aviso"
'''        End If
      Case 27 ' esc
        Call cmdSalir_Click
      Case 13 ' si presiono enter
        SendKeys "{Tab}"
    End Select
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub cmdCronograma_Click()
    Dim rs As New ADODB.Recordset
    'imprime la orden de pago
    Dim IResult As Variant
    Dim i As Integer
    Dim j As Integer
    Dim c As Integer
    Dim d As Integer
    Dim ncite As String
    Dim ObjetivoCons As String
    Dim COD_ORDEN As String
    Dim Poa As String
    Dim NroPagos As Currency
    Dim Factura As String
    Dim DocAdjunto As String

    On Error GoTo EtiqError
    ' obtiene el numero de pagos
    SQLs = "SELECT 'NroPagos'=MAX(numero_pago) FROM ao_pagos_cronograma WHERE estado_aprobado in('N','S') and estado_devengado IN ('N','S') and ges_gestion = '" & CodGestion & "' and codigo_unidad = '" & CodUnidad & "' and codigo_grupo = " & CodGrupo
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        NroPagos = rstTemp!NroPagos
      Else
        NroPagos = 0
    End If

    Screen.MousePointer = vbHourglass
    SQLs = "SELECT ao_pagos_cronograma.*, ao_pagos_grupos.modalidad_pago AS modalidad_pago FROM ao_pagos_cronograma INNER JOIN ao_pagos_grupos ON ao_pagos_cronograma.ges_gestion = ao_pagos_grupos.ges_gestion AND "
    SQLs = SQLs & "ao_pagos_cronograma.codigo_unidad = ao_pagos_grupos.codigo_unidad AND ao_pagos_cronograma.codigo_grupo = ao_pagos_grupos.codigo_grupo "
    SQLs = SQLs & "WHERE ao_pagos_cronograma.estado_aprobado in('N','S') and ao_pagos_cronograma.estado_devengado IN ('N','S') and ao_pagos_cronograma.ges_gestion = '" & CodGestion & "' and ao_pagos_cronograma.codigo_unidad = '" & CodUnidad & "' and ao_pagos_cronograma.codigo_grupo = " & CodGrupo
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    If rstTemp.RecordCount > 0 Then ' si devolvio por lo menos un registro
            Screen.MousePointer = vbHourglass
            
            Select Case rstTemp!modalidad_pago
              Case "I" 'modalidad de pago individual
                DE.dbo_ap_PagosRptOP_SacaSaldo rstTemp!ges_gestion, rstTemp!codigo_unidad, rstTemp!codigo_grupo, NroPagos, 0
                
                DE.dbo_ap_PagosRptCronogramaOP_c rstTemp!ges_gestion, rstTemp!codigo_unidad, rstTemp!codigo_grupo, NroPagos

                Set rs = DE.rsdbo_ap_PagosRptCronogramaOP_c.Clone
                DE.rsdbo_ap_PagosRptCronogramaOP_c.Close
                
                CRCrono.Formulas(20) = "Fecha='" & Me.dtpFOrdenPago.Value & "'"
                CRCrono.Formulas(30) = "POAs='" & (Poa) & "'"
                
                CRCrono.Formulas(45) = "AQuien='" & rs_AQuien!des_responsable & "'"
                CRCrono.Formulas(50) = "PuestoAQuien='" & rs_AQuien!des_puesto & "'"
                CRCrono.Formulas(55) = "DeQuien='" & rs_DeQuien!des_responsable & "'"
                CRCrono.Formulas(60) = "PuestoDeQuien='" & rs_DeQuien!des_puesto & "'"

                CRCrono.WindowShowGroupTree = False
                
                CRCrono.ReportFileName = App.Path & "\consultoria\rptCronogramaOrdenPago_C.rpt"
                            
                CRCrono.StoredProcParam(0) = rstTemp!ges_gestion
                CRCrono.StoredProcParam(1) = rstTemp!codigo_unidad
                CRCrono.StoredProcParam(2) = rstTemp!codigo_grupo
                CRCrono.StoredProcParam(3) = NroPagos
                
                IResult = CRCrono.PrintReport
                If IResult <> 0 Then MsgBox CRCrono.LastErrorNumber & " : " & CRCrono.LastErrorString, vbCritical, "Error de impresión"
              
              Case "P" ' se trata de modalidad por planilla de grupo
              
                DE.dbo_ap_PagosRptOP_SacaSaldo rstTemp!ges_gestion, rstTemp!codigo_unidad, rstTemp!codigo_grupo, NroPagos, 0

                DE.dbo_ap_PagosRptCronoOPPlanilla_c rstTemp!ges_gestion, rstTemp!codigo_unidad, rstTemp!codigo_grupo, NroPagos

                Set rs = DE.rsdbo_ap_PagosRptCronoOPPlanilla_c.Clone
                DE.rsdbo_ap_PagosRptCronoOPPlanilla_c.Close
    
                CRCrono.Formulas(20) = "Fecha='" & Me.dtpFOrdenPago.Value & "'"
                CRCrono.Formulas(30) = "POAs='" & (Poa) & "'"

                CRCrono.Formulas(45) = "AQuien='" & rs_AQuien!des_responsable & "'"
                CRCrono.Formulas(50) = "PuestoAQuien='" & rs_AQuien!des_puesto & "'"
                CRCrono.Formulas(55) = "DeQuien='" & rs_DeQuien!des_responsable & "'"
                CRCrono.Formulas(60) = "PuestoDeQuien='" & rs_DeQuien!des_puesto & "'"
                
                CRCrono.WindowShowGroupTree = False
                
                CRCrono.ReportFileName = App.Path & "\consultoria\rptCronoOrdenPago_Planilla_C.rpt"
                            
                CRCrono.StoredProcParam(0) = rstTemp!ges_gestion
                CRCrono.StoredProcParam(1) = rstTemp!codigo_unidad
                CRCrono.StoredProcParam(2) = rstTemp!codigo_grupo
                CRCrono.StoredProcParam(3) = NroPagos
                
                IResult = CRCrono.PrintReport
                If IResult <> 0 Then MsgBox CRCrono.LastErrorNumber & " : " & CRCrono.LastErrorString, vbCritical, "Error de impresión"
              
            End Select
      Else
        MsgBox "No existen registros para ser impresos.", vbInformation, "Aviso"
    End If

    Screen.MousePointer = vbDefault
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    Dim rs As New ADODB.Recordset
    'imprime la orden de pago
    Dim IResult As Variant
    Dim i As Integer
    Dim j As Integer
    Dim ncite As String
    Dim ObjetivoCons As String
    Dim COD_ORDEN As String
    Dim Poa As String
    Dim CONCEPTO_CABECERA As String
    Dim MONTO_US_CABE As Currency
    Dim MONTO_BS_CABE As Currency
    Dim Factura As String

    On Error GoTo EtiqError
    
    If fl_VerificaPrintOP Then
            
        Screen.MousePointer = vbHourglass
        SQLs = "SELECT ao_pagos_cronograma.*, ao_pagos_grupos.modalidad_pago AS modalidad_pago FROM ao_pagos_cronograma INNER JOIN ao_pagos_grupos ON ao_pagos_cronograma.ges_gestion = ao_pagos_grupos.ges_gestion AND "
        SQLs = SQLs & "ao_pagos_cronograma.codigo_unidad = ao_pagos_grupos.codigo_unidad AND ao_pagos_cronograma.codigo_grupo = ao_pagos_grupos.codigo_grupo "
        SQLs = SQLs & "WHERE ao_pagos_cronograma.estado_aprobado ='S' and ao_pagos_cronograma.estado_devengado = 'S' and ao_pagos_cronograma.ges_gestion = '" & CodGestion & "' and ao_pagos_cronograma.codigo_unidad = '" & CodUnidad & "' and ao_pagos_cronograma.codigo_grupo = " & CodGrupo & " and ao_pagos_cronograma.numero_pago = " & NumPago
        
        Set rstTemp = New ADODB.Recordset
        rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
        
        If rstTemp.RecordCount > 0 Then ' si devolvio por lo menos un registro

            Select Case rstTemp!modalidad_pago
              ''********************************************
              ''********INDIVUDUAL
              ''********************************************
              Case "I" 'modalidad de pago individual
                DE.dbo_ap_PagosRptOP_SacaSaldo CodGestion, CodUnidad, CodGrupo, NumPago, rstTemp!codigo_orden

                DE.dbo_ap_PagosRptOP_c CodGestion, CodUnidad, CodGrupo, NumPago

''                Set rs = DE.rsdbo_ap_PagosRptOP_c.Clone
''                DE.rsdbo_ap_PagosRptOP_c.Close

    
                DE.dbo_ap_SacaMontosDol CodGestion, CodUnidad, CodGrupo, NumPago, COD_ORDEN, MONTO_US_CABE, MONTO_BS_CABE, Poa, ncite, Factura, ObjetivoCons
                
                If Factura = "S" Then
                    Factura = "No" ' significa SIN (NO)retención por impuestos de ley
                Else
                    Factura = "Si" ' significa CON (SI)retención por impuestos de ley
                End If
                
                If rstTemp!tipo_moneda = "$US" Then
                    CR.Formulas(10) = "literal='" & Literal(Val(MONTO_US_CABE)) & "'"
                  Else
                    CR.Formulas(10) = "literal='" & Literal(Val(MONTO_BS_CABE)) & "'"
                End If
                
                CR.Formulas(20) = "Fecha='" & Me.dtpFOrdenPago.Value & "'"
                CR.Formulas(30) = "POAs='" & (Poa) & "'"
                CR.Formulas(40) = "XXFACTURA='" & (Factura) & "'"
                CR.Formulas(45) = "AQuien='" & rs_AQuien!des_responsable & "'"
                CR.Formulas(50) = "PuestoAQuien='" & rs_AQuien!des_puesto & "'"
                CR.Formulas(55) = "DeQuien='" & rs_DeQuien!des_responsable & "'"
                CR.Formulas(60) = "PuestoDeQuien='" & rs_DeQuien!des_puesto & "'"
                                
                'limpia los docs
                For i = 1 To 18
                        CR.Formulas(65 + i) = "DocAdjunto_" & i & "=''"
                Next i
                    
                ' carga los documentos adjuntos
                If lstDocAdjunto.Selected(0) = True Then ' todos los docuemntos adjuntos
                    For i = 1 To lstDocAdjunto.ListCount - 1
                        CR.Formulas(65 + i) = "DocAdjunto_" & i & "='" & Chr(164) & " " & lstDocAdjunto.List(i) & "'"
                    
                    Next i
                  Else
                    j = 1
                    For i = 1 To lstDocAdjunto.ListCount - 1
                        If lstDocAdjunto.Selected(i) = True Then
                            CR.Formulas(65 + i) = "DocAdjunto_" & j & "='" & Chr(164) & " " & lstDocAdjunto.List(i) & "'"
                            j = j + 1
                        End If
                    Next i
                End If
                
                CR.WindowShowGroupTree = False
                
                CR.ReportFileName = App.Path & "\consultoria\rptOrdenPago_C.rpt"
                            
                CR.StoredProcParam(0) = CodGestion
                CR.StoredProcParam(1) = CodUnidad
                CR.StoredProcParam(2) = CodGrupo
                CR.StoredProcParam(3) = NumPago
                
                IResult = CR.PrintReport
                If IResult <> 0 Then MsgBox CR.LastErrorNumber & " : " & CR.LastErrorString, vbCritical, "Error de impresión"
              
              ''********************************************
              ''********PLANILLA
              ''********************************************
              Case "P" ' se trata de modalidad por planilla de grupo
              
                DE.dbo_ap_PagosRptOP_SacaSaldo CodGestion, CodUnidad, CodGrupo, NumPago, rstTemp!codigo_orden

                DE.dbo_ap_PagosRptOPPlanilla_c CodGestion, CodUnidad, CodGrupo, NumPago
'''
'''                Set rs = DE.rsdbo_ap_PagosRptOPPlanilla_c.Clone
'''                DE.rsdbo_ap_PagosRptOPPlanilla_c.Close
    
                DE.dbo_ap_SacaMontosDol CodGestion, CodUnidad, CodGrupo, NumPago, COD_ORDEN, MONTO_US_CABE, MONTO_BS_CABE, Poa, ncite, Factura, ObjetivoCons
                
                If rstTemp!tipo_moneda = "$US" Then
                    CR.Formulas(10) = "literal='" & Literal(Val(MONTO_US_CABE)) & "'"
                  Else
                    CR.Formulas(10) = "literal='" & Literal(Val(MONTO_BS_CABE)) & "'"
                End If
                
                CR.Formulas(20) = "Fecha='" & Me.dtpFOrdenPago.Value & "'"
                CR.Formulas(30) = "POAs='" & (Poa) & "'"
                CR.Formulas(45) = "AQuien='" & rs_AQuien!des_responsable & "'"
                CR.Formulas(50) = "PuestoAQuien='" & rs_AQuien!des_puesto & "'"
                CR.Formulas(55) = "DeQuien='" & rs_DeQuien!des_responsable & "'"
                CR.Formulas(60) = "PuestoDeQuien='" & rs_DeQuien!des_puesto & "'"
                                
                'limpia los docs objeto de crystal
                For i = 1 To 16
                        CR.Formulas(65 + i) = "DocAdjunto_" & i & "=''"
                Next i
                
                ' carga los documentos adjuntos
                If lstDocAdjunto.Selected(0) = True Then ' todos los docuemntos adjuntos
                    For i = 1 To lstDocAdjunto.ListCount - 1
                        CR.Formulas(65 + i) = "DocAdjunto_" & i & "='" & Chr(164) & " " & lstDocAdjunto.List(i) & "'"
                    
                    Next i
                  Else
                    j = 1
                    For i = 1 To lstDocAdjunto.ListCount - 1
                        If lstDocAdjunto.Selected(i) = True Then
                            CR.Formulas(65 + i) = "DocAdjunto_" & j & "='" & Chr(164) & " " & lstDocAdjunto.List(i) & "'"
                            j = j + 1
                        End If
                    Next i
                End If
                
                CR.WindowShowGroupTree = False
                
                CR.ReportFileName = App.Path & "\consultoria\rptOrdenPago_Planilla_C.rpt"
                            
                CR.StoredProcParam(0) = CodGestion
                CR.StoredProcParam(1) = CodUnidad
                CR.StoredProcParam(2) = CodGrupo
                CR.StoredProcParam(3) = NumPago
                
                IResult = CR.PrintReport
                If IResult <> 0 Then MsgBox CR.LastErrorNumber & " : " & CR.LastErrorString, vbCritical, "Error de impresión"
              
            End Select
            
          Else
            MsgBox "La liquidación seleccionada Nro.[" & NumPago & "] no se encuentra devengado." & Chr(13) & "Corrija el error e intente imprimir nuevamente.", vbInformation, "Aviso"
        End If
    End If
    Screen.MousePointer = vbDefault
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
End Sub

Private Function fl_VerificaPrintOP() As Boolean
    'TITULO:                Función fl_VerificaPrintOP
    'PROPOSITO:             Verifica los datos la impresion
    'EJEMPLO DE LLAMADA:    fl_VerificaPrintOP
    Dim i As Integer
    fl_VerificaPrintOP = True ' asuminos que se cuenta con los datos
    
    ' si seleccionno docs adjuntos
    For i = 0 To lstDocAdjunto.ListCount - 1
        If lstDocAdjunto.Selected(i) = True Then
            Exit For
        End If
    Next i
    
    If i >= lstDocAdjunto.ListCount Then
        MsgBox "No seleciono ningún documento adjunto." & Chr(13) & "Corrija el error e intente imprimir nuevamente.", vbInformation, "Aviso"
        fl_VerificaPrintOP = False
        lstDocAdjunto.SetFocus ' se posiciona en el boton
        Exit Function
    End If
    ' selecciono correctamente dequien y aquien
    If Len(Trim(cboDeQuien.BoundText)) = 0 Then
        MsgBox "No selecciono 'DE QUIEN' es la orden de liquidación." & Chr(13) & "Corrija el error e intente procesar nuevamente.", vbInformation, "Aviso"
        fl_VerificaPrintOP = False
        Exit Function
    End If
    
    ' selecciono correctamente dequien y aquien
    If Len(Trim(cboAQuien.BoundText)) = 0 Then
        MsgBox "No selecciono 'A QUIEN' es la orden de liquidación." & Chr(13) & "Corrija el error e intente procesar nuevamente.", vbInformation, "Aviso"
        fl_VerificaPrintOP = False
        Exit Function
    End If
    
    ' existe coherencia dequien y aquien
    If cboDeQuien.BoundText = cboAQuien.BoundText Then
        MsgBox "No existe coherencia 'DE QUIEN'  y 'A QUIEN' es la orden de liquidación." & Chr(13) & "Corrija el error e intente procesar nuevamente.", vbInformation, "Aviso"
        fl_VerificaPrintOP = False
        Exit Function
    End If
End Function

Private Sub Form_Load()
    On Error GoTo EtiqError
    
    If GlProceso = "F05" Then
        Me.Caption = "SAF - Consultoría - Impresión de Ordenes de Pago"
    Else
        Me.Caption = "SAF - Recursos Humanos - Impresión de Ordenes de Pago"
    End If
    
    CodGestion = af_LiquidaMain_c.lblGestion.Caption
    CodUnidad = af_LiquidaMain_c.lblCodUniSol.Caption
    CodGrupo = Val(af_LiquidaMain_c.lblCodGrupo.Caption)
    NumPago = Val(af_LiquidaMain_c.lblEstadoBeneficiario.Tag)
    
    Call pl_Llena_Combos_Base
    
    Call pl_ValoresPorDefecto

    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

	Call SeguridadSet(Me)
End Sub

Private Sub pl_ValoresPorDefecto()
    Dim fechax As String
    Dim hhx As String
    
    DE.dbo_ap_GetServDateTime fechax, hhx
    dtpFOrdenPago.Value = fechax
    
    Select Case GldaCodigo
      Case "01" ' definido por el responsable de adqusiciones DGAAYRRHH
        cboPuestoDeQuien.BoundText = 2 ' director DGAAYRRHH
'        cboPuestoAQuien.BoundText = 1 ' director DGA GRAL
        cboPuestoAQuien.BoundText = 2 ' director DGF
      Case "52" ' definido por el responsable de adqusiciones DAP
        cboPuestoDeQuien.BoundText = 3 ' director RESPONSABLE DE ADQ
        cboPuestoAQuien.BoundText = 2 ' director DAP
      
    End Select
    
    
End Sub

Private Sub pl_Llena_Combos_Base()
    ' llena los combos y listas base para la carga del formulario
    
    Dim rs As ADODB.Recordset ' usado para la carga de los combos de base
    
    ' DATOS DEL RESPONSABLES DE LA UNIDAD para el reporte
    
    SQLs = "select * from ac_unidad_puesto_para_rep where ges_gestion = '" & CodGestion & "' and activo='S' and codigo_puesto<>1 "
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        Set cboPuestoDeQuien.RowSource = rstTemp
        cboPuestoDeQuien.BoundColumn = "codigo_puesto"
        cboPuestoDeQuien.ListField = "des_corta_puesto"
    End If
    
    SQLs = "select * from ac_unidad_puesto_para_rep where ges_gestion = '" & CodGestion & "' and activo='S' "
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        Set cboPuestoAQuien.RowSource = rstTemp
        cboPuestoAQuien.BoundColumn = "codigo_puesto"
        cboPuestoAQuien.ListField = "des_corta_puesto"
      Else
        MsgBox "El catalogo de puestos de la unidad vacio.", vbInformation, "Aviso"
    End If
    
    SQLs = "select * from ac_unidad_respondable_para_rep where ges_gestion = '" & CodGestion & "' and activo='S' and da in('DGF','" & GldaCodigo & "')"
    Set rs_DeQuien = New ADODB.Recordset
    Set rs_AQuien = New ADODB.Recordset
    rs_DeQuien.Open SQLs, db, adOpenStatic, adLockReadOnly
    rs_AQuien.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rs_DeQuien.RecordCount > 0 Then
        Set cboDeQuien.RowSource = rs_DeQuien
        cboDeQuien.BoundColumn = "ci"
        cboDeQuien.ListField = "des_responsable"
        
        Set cboAQuien.RowSource = rs_AQuien
        cboAQuien.BoundColumn = "ci"
        cboAQuien.ListField = "des_responsable"
      Else
        MsgBox "El catalogo de responsables de unidad esta vacio.", vbInformation, "Aviso"
    End If

    ' documentos adjuntos
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open "select cod_doc_adjunto, des_doc_adjunto from ac_doc_adjunto_rpt_c where activo='S'", db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then ' si devolvio por lo menos un registro
        lstDocAdjunto.AddItem "Todos"
        lstDocAdjunto.ItemData(lstDocAdjunto.NewIndex) = 999 ' para filtrar todos
        While Not rstTemp.EOF
        lstDocAdjunto.AddItem rstTemp!des_doc_adjunto
        lstDocAdjunto.ItemData(lstDocAdjunto.NewIndex) = rstTemp!cod_doc_adjunto
        rstTemp.MoveNext
        Wend
      Else
        MsgBox "No se tiene elementos en el catálogo de documentos adjuntos.", vbInformation, "Aviso"
    End If

End Sub

Private Sub lstDocAdjunto_ItemCheck(Item As Integer)
    Dim i As Integer
    If Item = 0 Then
        If lstDocAdjunto.Selected(0) = True Then
          For i = 1 To lstDocAdjunto.ListCount - 1
            lstDocAdjunto.Selected(i) = False
          Next
        End If
    Else
        lstDocAdjunto.Selected(0) = False
    End If
    lstDocAdjunto.ToolTipText = lstDocAdjunto.List(Item)

End Sub

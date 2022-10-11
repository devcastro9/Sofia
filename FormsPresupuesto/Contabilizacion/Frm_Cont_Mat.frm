VERSION 5.00
Begin VB.Form Frm_Cont_Mat 
   Caption         =   "."
   ClientHeight    =   1665
   ClientLeft      =   6675
   ClientTop       =   4440
   ClientWidth     =   3480
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3480
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "APROBACION . . ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1680
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   3450
      Begin VB.CommandButton Cmd_ContaConf 
         Caption         =   "Comprobante Aprobado . . . Click para continuar"
         Height          =   1080
         Left            =   336
         Picture         =   "Frm_Cont_Mat.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   2820
      End
   End
End
Attribute VB_Name = "Frm_Cont_Mat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************
'******** MODULO CORREGIDO 2 ********

Private Sub Cmd_Asiento_Click()
    Frm_Asiento.Show vbModal
End Sub

Private Sub Cmd_Aux_Mayor_Click()
    Frm_Aux_Conta.Show vbModal
End Sub

Private Sub Cmd_contaCancel_Click()
On Error GoTo errorComp2

db.RollbackTrans
MsgBox "Cancelando......."
Exit Sub

errorComp2:

MsgBox "Error al intentar Cancelar"

End Sub

Private Sub Cmd_ContaConf_Click()
Dim SW As Boolean
Dim Sw_Fuente As Boolean
Dim Cont_Comp As Long
Dim rstipopy As ADODB.Recordset
Set rstipopy = New ADODB.Recordset
Dim aux_T As String

'On Error GoTo errorComp
db.BeginTrans


'********* Para obtener en el recordset recsetAuxComp los datos necesarios para almacenar*********"
    Set recSetAuxcomp = New ADODB.Recordset
    recSetAuxcomp.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
    If recSetAuxcomp.State = 1 Then recSetAuxcomp.Close
    recSetAuxcomp.Open "SELECT pagos.codigo_convenio, pago_detalle.codigo_Pago,pagos.codigo_solicitud,pago_detalle.codigo_Pago_detalle,Pagos.Fte_Codigo,pagos.Ges_Gestion," & _
    " Pago_Detalle.Codigo_Beneficiario,pagos.Justificacion,pagos.fecha_aprueba as fecha_pago,pago_detalle.par_codigo,pago_detalle.Monto_total,Pagos.org_Codigo,pagos.Codigo_orden,Pagos.Codigo_documento," & _
    " pago_detalle.pro_programa, pago_detalle.pro_subprograma, pago_detalle.pro_proyecto, pago_detalle.pro_actividad, " & _
    " pago_detalle.Monto_Dolares,pago_detalle.Tipo_Cambio,pago_detalle.estado_aprobacion, PAGOS.CODIGO_UNIDAD From pago_detalle,pagos Where pago_detalle.codigo_Pago = pagos.codigo_Pago and pago_detalle.Org_Codigo = pagos.Org_codigo and TIPO_COMP='DAC' AND " & _
    " pago_detalle.Ges_Gestion = pagos.Ges_Gestion AND pagos.codigo_Pago= " & ff_egresos.AdoRegularizacion.Recordset!codigo_pago & " and pagos.Org_Codigo='" & ff_egresos.AdoRegularizacion.Recordset!org_codigo & "' and pago_detalle.Ges_Gestion = '" & ff_egresos.AdoRegularizacion.Recordset!ges_gestion & "'", db, adOpenKeyset, adLockOptimistic
    If recSetAuxcomp.RecordCount > 0 Then recSetAuxcomp.MoveFirst
    ' -> add si emite factura ?  -----------------------------------------------
While Not (recSetAuxcomp.EOF)
      If rstipopy.State = 1 Then rstipopy.Close
      Dim sqlpy  As String
      Dim tipopy As String
      'rstipopy.Open "select tipo_proyecto from fc_estructura_programatica where Pro_programa='" & recSetAuxcomp!pro_programa & "' and  Pro_subprograma='" & recSetAuxcomp!pro_subprograma & "' and Pro_proyecto='" & recSetAuxcomp!pro_proyecto & "' and Pro_actividad='" & recSetAuxcomp!pro_actividad & "'", db, adOpenKeyset, adLockReadOnly
      rstipopy.Open "select tipo_proyecto from fc_estructura_programatica where Pro_programa='" & recSetAuxcomp!pro_programa & "' and Pro_proyecto='" & recSetAuxcomp!pro_proyecto & "' and Pro_actividad='" & recSetAuxcomp!pro_actividad & "'", db, adOpenKeyset, adLockReadOnly
      If rstipopy.RecordCount <> 0 Then
          tipopy = rstipopy!tipo_proyecto
      Else
         ' MsgBox "La Categoria Programática elegida no existe"
         MsgBox "Error en la contabilización, No se encontró la Estructura Programática", vbExclamation + vbDefaultButton1
         Exit Sub
      End If
      Set recSetPartida = New ADODB.Recordset
      If recSetPartida.State = 1 Then recSetPartida.Close
      Select Case tipopy
        Case "N"
            recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H,CC_Cuentas_D" & _
                " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst='DEV' and CC_Cuenta_H.Inst='DEV' and" & _
                " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
    
        Case "S"
            recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H,CC_Cuentas_D" & _
                " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst='PSD' and CC_Cuenta_H.Inst='PSD' and" & _
                " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
    
        Case "F"
            recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H,CC_Cuentas_D" & _
                " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst='PFD' and CC_Cuenta_H.Inst='PFD' and" & _
                " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
      End Select
    
    If recSetPartida.BOF And recSetPartida.EOF Then
        MsgBox "No Existe Partida"
    Else
    
      '************Abrimos un record set para adicionar datos*********************'
      Set recSetAuxActualizar = New ADODB.Recordset
      If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
      recSetAuxActualizar.Open " select * from CO_Comprobante_M  where Cod_Trans=" & recSetAuxcomp!codigo_pago & " and Org_Codigo='" & recSetAuxcomp!org_codigo & "' " & _
      " and Ges_Gestion='" & recSetAuxcomp!ges_gestion & "' and tipo_comp='DAC' and Cod_Trans_Detalle='" & recSetAuxcomp!codigo_pago_detalle & "'", db, adOpenDynamic, adLockOptimistic, adCmdText
      'MsgBox recSetAuxActualizar.RecordCount
      If Not recSetAuxActualizar.BOF Then recSetAuxActualizar.MoveFirst
      If (recSetAuxActualizar.BOF) And (recSetAuxActualizar.EOF) Then
        '************* GENERA EL CODIGO DE COMPROBANTE**********
         Set recSetGenera = New ADODB.Recordset
         recSetGenera.CursorLocation = adUseClient
         If recSetGenera.State = 1 Then recSetGenera.Close
         recSetGenera.Open "select * from fc_Correl  where tipo_tramite='cmbte'", db, adOpenDynamic, adLockOptimistic, adCmdText
         If recSetGenera.RecordCount > 0 Then
                Cont_Comp = Val(recSetGenera!numero_correlativo)
                Cont_Comp = Cont_Comp + 1
                recSetGenera!numero_correlativo = Trim(Str(Cont_Comp))
                recSetGenera.Update
         End If
         If recSetGenera.State = 1 Then recSetGenera.Close
        '************TERMINA GENERACION DE COMPROBANTE********
        ' Datos Para co_Comprobante_M
    
         recSetAuxActualizar.AddNew
            
         recSetAuxActualizar!usr_usuario = GlUsuario
         recSetAuxActualizar!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
         recSetAuxActualizar!hora_registro = Format(Time, "hh:mm:ss")
         recSetAuxActualizar!Cod_Comp = Cont_Comp
         recSetAuxActualizar!cod_trans = recSetAuxcomp!codigo_pago
         recSetAuxActualizar!cod_trans_detalle = recSetAuxcomp!codigo_pago_detalle
         recSetAuxActualizar!org_codigo = recSetAuxcomp!org_codigo
         recSetAuxActualizar!codigo_beneficiario = recSetAuxcomp!codigo_beneficiario
         recSetAuxActualizar!ges_gestion = recSetAuxcomp!ges_gestion
         recSetAuxActualizar!num_respaldo = recSetAuxcomp!Codigo_orden
         recSetAuxActualizar!codigo_documento = IIf(IsNull(recSetAuxcomp!codigo_documento), "-", recSetAuxcomp!codigo_documento)
         recSetAuxActualizar!fecha_A = IIf(IsNull(recSetAuxcomp!fecha_pago), (Format(Date, "dd/mm/yyyy")), CDate(recSetAuxcomp!fecha_pago))
         recSetAuxActualizar!glosa = recSetAuxcomp!justificacion
         recSetAuxActualizar!tipo_comp = "DAC"
         recSetAuxActualizar!Status = "S"
         recSetAuxActualizar!codigo_unidad = recSetAuxcomp!codigo_unidad
         recSetAuxActualizar.Update
         If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
         Set recSetAuxActualizar1 = New ADODB.Recordset
         If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
         recSetAuxActualizar1.Open " select * from CO_Diario where cod_Comp = " & Cont_Comp & " ", db, adOpenKeyset, adLockOptimistic, adCmdText
         ' -- adicionar detalle CO_DIARIO -------------------------------------
         If (recSetAuxActualizar1.BOF) And (recSetAuxActualizar1.EOF) Then
            ' --------- DETALLE ASiENTO 1 - NORMAL -----------------------
             recSetAuxActualizar1.AddNew
             recSetAuxActualizar1!usr_usuario = GlUsuario
             recSetAuxActualizar1!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
             recSetAuxActualizar1!hora_registro = Format(Time, "hh:mm:ss")
             recSetAuxActualizar1!Cod_Comp = Cont_Comp
             recSetAuxActualizar1!Cod_Comp_C = 1
             recSetAuxActualizar1!tipo_comp = "DAC"
             recSetAuxActualizar1!d_cuenta = recSetPartida!cuenta
             recSetAuxActualizar1!D_Nombre = recSetPartida!NombreCta
             recSetAuxActualizar1!d_subcta1 = recSetPartida!subcta1
             recSetAuxActualizar1!d_subcta2 = recSetPartida!subcta2
             recSetAuxActualizar1!d_Aux1 = recSetPartida!aux1
             recSetAuxActualizar1!d_Aux2 = recSetPartida!AUX2
             recSetAuxActualizar1!d_Aux3 = recSetPartida!aux3
             ''******* ADICION DE AUXILIARES A DETALLE DEBITO*******
             Select Case recSetPartida!aux1
              Case "01"
                    recSetAuxActualizar1!d_cta_larga = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!d_cta_larga = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!d_cta_larga = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!d_cta_larga = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
             End Select
             '---------auxiliar 2
             Select Case recSetPartida!AUX2
              Case "01"
                    recSetAuxActualizar1!d_ctaaux2 = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!d_ctaaux2 = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!d_ctaaux2 = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!d_ctaaux2 = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
             End Select
             '----------auxiliar 3
             Select Case recSetPartida!aux3
              Case "01"
                   recSetAuxActualizar1!d_CtaAux3 = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!d_CtaAux3 = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!d_CtaAux3 = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!d_CtaAux3 = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
             End Select
             ''****************** finaliza sesion de auxiliares DEBITO
             recSetAuxActualizar1!h_cuenta = recSetPartida!h_cuenta
             recSetAuxActualizar1!H_Nombre = recSetPartida!H_NombCta
             recSetAuxActualizar1!h_subcta1 = recSetPartida!h_subcta1
             recSetAuxActualizar1!h_subcta2 = recSetPartida!h_subcta2
             recSetAuxActualizar1!h_Aux1 = recSetPartida!h_Aux1
             recSetAuxActualizar1!h_Aux2 = recSetPartida!h_Aux2
             recSetAuxActualizar1!h_Aux3 = recSetPartida!h_Aux3
             ''******* ADICION DE AUXILIARES AL HABER *******
             Select Case recSetPartida!h_Aux1
             Case "01"
                    recSetAuxActualizar1!h_cta_larga = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!h_cta_larga = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!h_cta_larga = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!h_cta_larga = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
             End Select
             '-------Haber-auxiliar 2
             Select Case recSetPartida!h_Aux2
              Case "01"
                    recSetAuxActualizar1!h_ctaaux2 = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!h_ctaaux2 = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!h_ctaaux2 = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!h_ctaaux2 = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
             End Select
             '----------auxiliar 3
             Select Case recSetPartida!h_Aux3
              Case "01"
                   recSetAuxActualizar1!h_CtaAux3 = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!h_CtaAux3 = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!h_CtaAux3 = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!h_CtaAux3 = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
             End Select
             ''****************** finaliza sesion de auxiliares en el haber
                  
             recSetAuxActualizar1!d_montoBs = Round(recSetAuxcomp!monto_total * 0.87, 2)
             recSetAuxActualizar1!d_montoDl = Round(recSetAuxcomp!monto_dolares * 0.87, 2)
             recSetAuxActualizar1!d_Cambio = recSetAuxcomp!tipo_cambio
             recSetAuxActualizar1!h_montoBs = Round(recSetAuxcomp!monto_total * 0.87, 2)
             recSetAuxActualizar1!h_montoDl = Round(recSetAuxcomp!monto_dolares * 0.87, 2)
             recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
             recSetAuxActualizar1!usr_usuario = GlUsuario
             recSetAuxActualizar1!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
             recSetAuxActualizar1!hora_registro = Format(Time, "hh:mm:ss")
             recSetAuxActualizar1.Update
             'If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
             ' --------- DETALLE ASIENTO 2 - CREDITO FISCAL -----------------------
             recSetAuxActualizar1.AddNew
             recSetAuxActualizar1!usr_usuario = GlUsuario
             recSetAuxActualizar1!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
             recSetAuxActualizar1!hora_registro = Format(Time, "hh:mm:ss")
             recSetAuxActualizar1!Cod_Comp = Cont_Comp
             recSetAuxActualizar1!Cod_Comp_C = 2
             recSetAuxActualizar1!tipo_comp = "DAC"
             recSetAuxActualizar1!d_cuenta = "1150"
             recSetAuxActualizar1!D_Nombre = "-"
             recSetAuxActualizar1!d_subcta1 = "01"
             recSetAuxActualizar1!d_subcta2 = "00"
             recSetAuxActualizar1!d_Aux1 = "01"
             recSetAuxActualizar1!d_Aux2 = "09"
             recSetAuxActualizar1!d_Aux3 = "00"
             ''******* ADICION DE AUXILIARES A DETALLE DEBITO*******
             Select Case recSetPartida!aux1
              Case "01"
                    recSetAuxActualizar1!d_cta_larga = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!d_cta_larga = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!d_cta_larga = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!d_cta_larga = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
             End Select
             '---------auxiliar 2
             Select Case recSetPartida!AUX2
              Case "01"
                    recSetAuxActualizar1!d_ctaaux2 = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!d_ctaaux2 = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!d_ctaaux2 = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!d_ctaaux2 = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
             End Select
             '----------auxiliar 3
             Select Case recSetPartida!aux3
              Case "01"
                   recSetAuxActualizar1!d_CtaAux3 = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!d_CtaAux3 = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!d_CtaAux3 = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!d_CtaAux3 = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
             End Select
             ''****************** finaliza sesion de auxiliares DEBITO
             recSetAuxActualizar1!h_cuenta = recSetPartida!h_cuenta
             recSetAuxActualizar1!H_Nombre = recSetPartida!H_NombCta
             recSetAuxActualizar1!h_subcta1 = recSetPartida!h_subcta1
             recSetAuxActualizar1!h_subcta2 = recSetPartida!h_subcta2
             recSetAuxActualizar1!h_Aux1 = recSetPartida!h_Aux1
             recSetAuxActualizar1!h_Aux2 = recSetPartida!h_Aux2
             recSetAuxActualizar1!h_Aux3 = recSetPartida!h_Aux3
             ''******* ADICION DE AUXILIARES AL HABER *******
             Select Case recSetPartida!h_Aux1
             Case "01"
                    recSetAuxActualizar1!h_cta_larga = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!h_cta_larga = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!h_cta_larga = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!h_cta_larga = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
             End Select
             '-------Haber-auxiliar 2
             Select Case recSetPartida!h_Aux2
              Case "01"
                    recSetAuxActualizar1!h_ctaaux2 = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!h_ctaaux2 = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!h_ctaaux2 = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!h_ctaaux2 = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
             End Select
             '----------auxiliar 3
             Select Case recSetPartida!h_Aux3
              Case "01"
                   recSetAuxActualizar1!h_CtaAux3 = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!h_CtaAux3 = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!h_CtaAux3 = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!h_CtaAux3 = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
             End Select
             ''****************** finaliza sesion de auxiliares en el haber
                  
             recSetAuxActualizar1!d_montoBs = Round(recSetAuxcomp!monto_total * 0.13, 2)
             recSetAuxActualizar1!d_montoDl = Round(recSetAuxcomp!monto_dolares * 0.13, 2)
             recSetAuxActualizar1!d_Cambio = recSetAuxcomp!tipo_cambio
             recSetAuxActualizar1!h_montoBs = Round(recSetAuxcomp!monto_total * 0.13, 2)
             recSetAuxActualizar1!h_montoDl = Round(recSetAuxcomp!monto_dolares * 0.13, 2)
             recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
             recSetAuxActualizar1!usr_usuario = GlUsuario
             recSetAuxActualizar1!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
             recSetAuxActualizar1!hora_registro = Format(Time, "hh:mm:ss")
             recSetAuxActualizar1.Update
             If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
            
         End If 'Adicion del diario
      Else
            MsgBox "Ya fue contabilizado anteriormente...  ", vbOKOnly, "Se reemplaran los datos del Comprobante....  "
            'Modifica registro existente
            'recSetAuxActualizar!Cod_Comp = Cont_Comp
            recSetAuxActualizar!usr_usuario = GlUsuario
            Cont_Comp = recSetAuxActualizar!Cod_Comp
            recSetAuxActualizar!cod_trans = recSetAuxcomp!codigo_pago
            recSetAuxActualizar!cod_trans_detalle = recSetAuxcomp!codigo_pago_detalle
            recSetAuxActualizar!org_codigo = recSetAuxcomp!org_codigo
            recSetAuxActualizar!codigo_beneficiario = recSetAuxcomp!codigo_beneficiario
            recSetAuxActualizar!ges_gestion = recSetAuxcomp!ges_gestion
            recSetAuxActualizar!num_respaldo = recSetAuxcomp!Codigo_orden
            recSetAuxActualizar!codigo_documento = recSetAuxcomp!codigo_documento
'            If IsNull(recSetAuxcomp!fecha_pago) Then
'             FECHA = Date
'            Else
'             FECHA = recSetAuxcomp!fecha_pago
'            End If
            recSetAuxActualizar!fecha_A = IIf(IsNull(recSetAuxcomp!fecha_pago), Format(Date, "dd/mm/yyyy"), CDate(Format(recSetAuxcomp!fecha_pago, "dd/mm/yyyy")))
            recSetAuxActualizar!glosa = recSetAuxcomp!justificacion
            recSetAuxActualizar!usr_usuario = GlUsuario
            recSetAuxActualizar!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
            recSetAuxActualizar!hora_registro = Format(Time, "hh:mm:ss")
            'recSetAuxActualizar!Tipo_Comp = "DAC"
            recSetAuxActualizar!Status = "S"
            recSetAuxActualizar.Update
            If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
   'Modificacione en el diario

            Set recSetAuxActualizar1 = New ADODB.Recordset
            If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
            recSetAuxActualizar1.Open " select * from CO_Diario where cod_Comp = " & Cont_Comp, db, adOpenDynamic, adLockOptimistic  ', adCmdText
            If (recSetAuxActualizar1.BOF) And (recSetAuxActualizar1.EOF) Then
                recSetAuxActualizar1.AddNew
                recSetAuxActualizar1!tipo_comp = "DAC"
                recSetAuxActualizar1!Cod_Comp = Cont_Comp
            Else
                    If (Not recSetAuxActualizar1.BOF) Then recSetAuxActualizar1.MoveFirst
            End If
            recSetAuxActualizar1!usr_usuario = GlUsuario + "MOD"
            recSetAuxActualizar1!d_cuenta = recSetPartida!cuenta
            recSetAuxActualizar1!D_Nombre = recSetPartida!NombreCta
            recSetAuxActualizar1!d_subcta1 = recSetPartida!subcta1
            recSetAuxActualizar1!d_subcta2 = recSetPartida!subcta2
            recSetAuxActualizar1!d_Aux1 = recSetPartida!aux1
            recSetAuxActualizar1!d_Aux2 = recSetPartida!AUX2
            recSetAuxActualizar1!d_Aux3 = recSetPartida!aux3

        ''******* ADICION DE AUXILIARES A DETALLE DEBITO*******
           ''******* ADICION DE AUXILIARES A DETALLE DEBITO*******
            Select Case recSetPartida!aux1
              Case "01"
                    recSetAuxActualizar1!d_cta_larga = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!d_cta_larga = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!d_cta_larga = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!d_cta_larga = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
            End Select
            '---------auxiliar 2
            Select Case recSetPartida!AUX2
              Case "01"
                    recSetAuxActualizar1!d_ctaaux2 = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!d_ctaaux2 = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!d_ctaaux2 = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!d_ctaaux2 = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
            End Select
            '----------auxiliar 3
            Select Case recSetPartida!aux3
              Case "01"
                   recSetAuxActualizar1!d_CtaAux3 = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!d_CtaAux3 = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!d_CtaAux3 = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!d_CtaAux3 = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
            End Select
        ''****************** finaliza sesion de auxiliares DEBITO
        
            recSetAuxActualizar1!h_cuenta = recSetPartida!h_cuenta
           ' recSetAuxActualizar1!H_Nombre = recSetPartida!H_NombCta
            recSetAuxActualizar1!h_subcta1 = recSetPartida!h_subcta1
            recSetAuxActualizar1!h_subcta2 = recSetPartida!h_subcta2
            recSetAuxActualizar1!h_Aux1 = recSetPartida!h_Aux1
            recSetAuxActualizar1!h_Aux2 = recSetPartida!h_Aux2
            recSetAuxActualizar1!h_Aux3 = recSetPartida!h_Aux3
        ''******* ADICION DE AUXILIARES A DETALLE*******
          ''******* ADICION DE AUXILIARES AL HABER *******
            Select Case recSetPartida!h_Aux1
              Case "01"
                    recSetAuxActualizar1!h_cta_larga = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!h_cta_larga = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!h_cta_larga = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!h_cta_larga = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
            Case Else
            End Select
             '-------Haber-auxiliar 2
            Select Case recSetPartida!h_Aux2
              Case "01"
                    recSetAuxActualizar1!h_ctaaux2 = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!h_ctaaux2 = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!h_ctaaux2 = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!h_ctaaux2 = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
            End Select
            '----------auxiliar 3
            Select Case recSetPartida!h_Aux3
              Case "01"
                   recSetAuxActualizar1!h_CtaAux3 = IIf(IsNull(recSetAuxcomp!codigo_beneficiario), "", recSetAuxcomp!codigo_beneficiario)
              Case "02"
                    recSetAuxActualizar1!h_CtaAux3 = IIf(IsNull(recSetAuxcomp!Cta_Codigo), "", recSetAuxcomp!Cta_Codigo)
              Case "08"
                    recSetAuxActualizar1!h_CtaAux3 = IIf(IsNull(recSetAuxcomp!org_codigo), "", recSetAuxcomp!org_codigo)
              Case "09"
                    recSetAuxActualizar1!h_CtaAux3 = IIf(IsNull(recSetAuxcomp!codigo_convenio), "", recSetAuxcomp!codigo_convenio)
              Case Else
            End Select
        ''****************** finaliza sesion de auxiliares
        ''****************** finaliza sesion de auxiliares
        
            
            recSetAuxActualizar1!d_montoBs = Round(recSetAuxcomp!monto_total, 2)
            recSetAuxActualizar1!d_montoDl = Round(recSetAuxcomp!monto_dolares, 2)
            recSetAuxActualizar1!d_Cambio = recSetAuxcomp!tipo_cambio
            
            recSetAuxActualizar1!h_montoBs = Round(recSetAuxcomp!monto_total, 2)
            recSetAuxActualizar1!h_montoDl = Round(recSetAuxcomp!monto_dolares, 2)
            recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
            recSetAuxActualizar1!usr_usuario = GlUsuario
            recSetAuxActualizar1!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
            recSetAuxActualizar1!hora_registro = Format(Time, "hh:mm:ss")
            
            recSetAuxActualizar1.Update
            If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
                       
            
      End If
    
    End If ' De Partida
recSetAuxcomp.MoveNext

Wend
                       

db.CommitTrans
MsgBox "Contabilizo....."
Unload Frm_Cont_Mat


Exit Sub
errorComp:

db.RollbackTrans
MsgBox "error al recuperar datos"
Unload Frm_Cont_Mat

End Sub

Private Sub Cmd_ContaGrab_Click()
On Error GoTo errorComp1

db.CommitTrans
MsgBox "Grabando.........."
Exit Sub
errorComp1:
MsgBox "Error al intentar grabar"

'On Error GoTo errorcomp
'db.Execute "Insert into Co_Comprobante_C SELECT  "",codigo_orden,codigo_orden_detalle,fecha_pago,concepto_pago,codigo_solicitud,compromiso_numero,'DEV','Devengado'," & _
'" From orden_pago_detalle,orden_pago WHERE orden_pago_detalle.codigo_orden = orden_pago.codigo_orden and estado_aprobacion='S' and " & _
'Exit Sub
'
End Sub

Private Sub Cmd_Salir_Click()
'Unload Frm_Cont_Mat
End Sub

Private Sub Form_Load()

Set recSetBusqueda = New ADODB.Recordset
Set recSetAuxActualizar = New ADODB.Recordset
Set recSetAuxActualizar1 = New ADODB.Recordset

Set recsetAdicion = New ADODB.Recordset

Set recSetAuxcomp = New ADODB.Recordset
Set recSetAuxcomp1 = New ADODB.Recordset
Set recSetPartida = New ADODB.Recordset
Set recSetPartida1 = New ADODB.Recordset

Set recSetAuxcomp1 = New ADODB.Recordset
Set recSetAuxbenefi1 = New ADODB.Recordset
Set recSetPartid1 = New ADODB.Recordset



recSetPartida.CursorLocation = adUseClient


recSetAuxcomp.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
recSetAuxActualizar.CursorLocation = adUseClient



'recSetPartida.Open "SELECT Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta From CC_Cuenta_H, CC_Cuentas_D" & _
'" WHERE CC_Cuenta_H.h_Cuenta <> CC_Cuentas_D.Cuenta AND CC_Cuenta_H.H_SubCta1 <> CC_Cuentas_D.SubCta1 AND" & _
'" CC_Cuenta_H.H_SubCta2 <> CC_Cuentas_D.SubCta2 AND CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND" & _
'" cc_Cuenta_H.Par_I<='" & aux & "' and  cc_Cuenta_H.Par_F>='" & aux & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText

'Select Case recSetAuxcomp!tipo_Beneficiario
'Case "R"
'
'recSetAuxcomp1.Open "SELECT Ruc_Documento_Id ,Ruc_Descripcion_Larga From orden_pago_detalle, ac_proveedor" & _
'" WHERE orden_pago_detalle.codigo_beneficiario = ac_proveedor.Ruc_Documento_Id and orden_pago_detalle.codigo_beneficiario ='" & recSetAuxcomp!codigo_beneficiario & " '", db, adOpenDynamic, adLockOptimistic, adCmdText
'Text11.Text = recSetAuxcomp1!Ruc_Descripcion_Larga
'
'Case "C", "U", "I", "P"
'recSetBusqueda.Open "Select CI,Paterno,Materno,Nombres from Funcionario,orden_pago_detalle " & _
'" where orden_pago_detalle.Codigo_Beneficiario=Funcionario.CI and orden_pago_detalle.codigo_beneficiario ='" & recSetAuxcomp!codigo_beneficiario & " '", db, adOpenDynamic, adLockOptimistic, adCmdText
'Text11.Text = recSetAuxcomp1!Nombres & " " & recSetAuxcomp1!Paterno & " " & recSetAuxcomp1!Materno
'
'Case Else
'MsgBox "Benenficiario no encontrado"
'End Select
                      
               
End Sub


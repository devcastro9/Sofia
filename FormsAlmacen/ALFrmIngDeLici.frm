VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form ALFrmIngDeLici 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingreso por Compras"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   9120
      TabIndex        =   3
      Top             =   6405
      Width           =   9120
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   435
         Left            =   7650
         TabIndex        =   2
         Top             =   150
         Width           =   1380
      End
      Begin VB.CommandButton cmdElegir 
         Caption         =   "Elegir"
         Height          =   435
         Left            =   6195
         TabIndex        =   1
         Top             =   150
         Width           =   1380
      End
   End
   Begin TrueOleDBGrid60.TDBGrid tdbgIngreso 
      Align           =   3  'Align Left
      Height          =   6405
      Left            =   0
      OleObjectBlob   =   "ALFrmIngDeLici.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   9120
   End
End
Attribute VB_Name = "ALFrmIngDeLici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsIngreso As ADODB.Recordset
Public QResp As Boolean
Public NoLicitacion As Long
Private Sub CmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdElegir_Click()
    If RsIngreso.RecordCount <= 0 Then Exit Sub
    NoLicitacion = RsIngreso!NoAdjudica
    If RsIngreso.EditMode <> adEditNone Then RsIngreso.Update
    If valida Then
        QResp = True
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    QResp = False
    With Me
        .Top = 0
        .Left = 0
        .Height = 5685
        .Width = 9210
    End With
    GlSqlAux = "DELETE FROM ALIngDeLici WHERE Usuario = '" & NombreTerminal & "'"
    db.Execute GlSqlAux
'    GlSqlAux = "INSERT INTO ALIngDeLici(Usuario, NoLici, FechLici, DescItem, CodProveedor, NomProv, CodGrupo, CodDetalle, CantSaldo) " & _
'               "SELECT '" & NombreTerminal & "', ao_Cotizaciones.Nro_Licitacion, " & _
'               "ao_Cotizaciones.fecha_cotizacion, ALCLDetalle.DescDetalle, " & _
'               "ao_cotizacion_detalle.Codigo_proveedor, " & _
'               "ac_proveedor.Descripcion_Larga, " & _
'               "ALCLDetalle.CodGrupo, " & _
'               "ALCLDetalle.CodDetalle, " & _
'               "ao_cotizacion_detalle.cantidad_cotizada - ao_cotizacion_detalle.cantidad_entregada " & _
'               "AS Saldo " & _
'               "FROM ao_Cotizaciones INNER JOIN " & _
'               "ao_cotizacion_detalle ON " & _
'               "ao_Cotizaciones.Nro_Licitacion = ao_cotizacion_detalle.Nro_Licitacion " & _
'               "Inner Join ALCLDetalle ON " & _
'               "ao_cotizacion_detalle.CodGrupo = ALCLDetalle.CodGrupo AND " & _
'               "ao_cotizacion_detalle.CodDetalle = ALCLDetalle.CodDetalle INNER " & _
'               "Join ac_proveedor ON ao_cotizacion_detalle.Codigo_proveedor = ac_proveedor.Ruc_Id " & _
'               "WHERE (ao_Cotizaciones.estado = 'S') AND (ao_cotizacion_detalle.cantidad_cotizada - ao_cotizacion_detalle.cantidad_entregada > 0)"
    GlSqlAux = "INSERT INTO ALIngDeLici(Usuario, CodGrupo, CodDetalle, NoLici, FechLici, CodProveedor, NomProv, DescItem, CantSaldo, NoAdjudica) " & _
               "SELECT '" & NombreTerminal & "', " & _
               "ao_adjudica_d.CodGrupo, " & _
               "ao_adjudica_d.codDetalle, " & _
               "Convert(VarChar(10),ao_adjudica_d.nro_licitacion) + '/' + Convert(VarChar(10),ao_adjudica_d.ges_gestion) AS NoLici, " & _
               "ao_adjudica_d.fecha_adjudicacion, " & _
               "aO_ADJUDICA_D.Ruc_Id, " & _
               "Fc_BENEFICIARIO.denominacion_beneficiario, " & _
               "ALCLDetalle.DescDetalle, " & _
               "ao_adjudica_d.cantidad_cotizada - ao_adjudica_d.cantidad_entregada AS Saldo, " & _
               "ao_adjudica_d.nro_adjudica " & _
               "FROM ao_adjudica_d INNER JOIN " & _
               "ALCLDetalle ON " & _
               "ao_adjudica_d.CodGrupo = ALCLDetalle.CodGrupo AND " & _
               "ao_adjudica_d.codDetalle = ALCLDetalle.CodDetalle INNER JOIN " & _
               "FC_BENEFICIARIO ON ao_adjudica_d.Ruc_Id = FC_BENEFICIARIO.CODIGO_BENEFICIARIO " & _
               "WHERE ao_adjudica_d.habilitado = 'N' AND (ao_adjudica_d.cantidad_cotizada - ao_adjudica_d.cantidad_entregada) > 0"
    db.Execute GlSqlAux
    GlSqlAux = "SELECT * FROM ALIngDeLici WHERE Usuario = '" & NombreTerminal & "'"
    Set RsIngreso = New ADODB.Recordset
    RsIngreso.Open GlSqlAux, db, adOpenStatic, adLockOptimistic
    Set tdbgIngreso.DataSource = RsIngreso
    cmdElegir.Enabled = RsIngreso.RecordCount > 0
    Screen.MousePointer = vbDefault
    '--
	Call SeguridadSet(Me)
End Sub
Private Sub Form_Resize()
On Error Resume Next
    tdbgIngreso.Width = Me.ScaleWidth
End Sub

Private Sub tdbgIngreso_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Select Case ColIndex
        Case 5 ' Cantidad a Entregar
            If CLng(tdbgIngreso.Columns("CantEnt").Value) > CLng(tdbgIngreso.Columns("CantSaldo").Value) Then
                MsgBox "No se entregar una cantidad del Item '" & tdbgIngreso.Columns("DescItem").Value & "' mayor a su Saldo disponible de entrega(Saldo = " & tdbgIngreso.Columns("CantSaldo").Text & ")", vbInformation + vbOKOnly, "Atención"
                tdbgIngreso.Columns("CantEnt").Value = OldValue
                Cancel = True
            End If
    End Select
End Sub

Private Sub tdbgIngreso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If RsIngreso.EditMode <> adEditNone Then RsIngreso.Update
    ElseIf KeyAscii = 27 Then
        If RsIngreso.EditMode <> adEditNone Then RsIngreso.CancelUpdate
    Else
        KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]", KeyAscii, 0)
    End If
End Sub
Private Function valida() As Boolean
Dim rs As ADODB.Recordset
Dim AuxValida As Boolean
Dim TodoCero As Boolean
Dim CadenaError As String
    valida = False
    '-- Validamos la cantidad
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT * FROM ALIngDeLici WHERE Usuario = '" & NombreTerminal & "' AND NoAdjudica = '" & NoLicitacion & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    If rs.RecordCount <= 0 Then
        MsgBox "Ooops!!!. No se encontro la Licitacion.", vbCritical + vbOKOnly, "Atención"
        Exit Function
    End If
    AuxValida = True
    TodoCero = True
    CadenaError = CadenaError & "------------------------------------------" & vbCrLf
    While Not rs.EOF
        If rs!CantEnt > 0 Then
            CadenaError = CadenaError & vbTab & "El Item '" & rs!DescItem & "' del Proveedor '" & rs!NomProv & "', tiene cantidad a entregar IGUAL A " & Format(rs!CantEnt, "###,###,##0") & "." & vbCrLf
            AuxValida = False
            TodoCero = False
        End If
        If rs!CantEnt <= 0 Then
            CadenaError = CadenaError & vbTab & "El Item '" & rs!DescItem & "' del Proveedor '" & rs!NomProv & "', tiene cantidad a entregar IGUAL A CERO." & vbCrLf
            AuxValida = False
        End If
        CadenaError = CadenaError & "------------------------------------------" & vbCrLf
        rs.MoveNext
    Wend
    If TodoCero Then
        MsgBox "Debe ingresar la Cantidad a Entregar de algún Item de la Lictación '" & NoLicitacion & "', para realizar esta operación.", vbInformation + vbOKOnly, "Atención"
        Exit Function
    End If
    If Not AuxValida Then
        If MsgBox("Para la Licitación '" & NoLicitacion & "' se tiene las siguientes caracteristicas de Entrega: " & vbCrLf & CadenaError & "Esta seguro de continuar con la Operación?.", vbQuestion + vbYesNo, "Atención") = vbNo Then
            Exit Function
        End If
    End If
    valida = True
End Function

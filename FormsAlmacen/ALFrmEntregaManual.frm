VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form ALFrmEntregaManual 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Entrega Manual"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBoton 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8925
      TabIndex        =   5
      Top             =   5145
      Width           =   8925
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar Detalle..."
         Height          =   360
         Left            =   60
         TabIndex        =   1
         Top             =   75
         Width           =   2475
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cantidad de Cajas a Entregar del Item = "
         Height          =   195
         Left            =   5925
         TabIndex        =   6
         Top             =   135
         Width           =   2850
      End
   End
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      Picture         =   "ALFrmEntregaManual.frx":0000
      ScaleHeight     =   690
      ScaleWidth      =   8865
      TabIndex        =   4
      Top             =   5640
      Width           =   8925
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   435
         Left            =   1725
         TabIndex        =   3
         Top             =   135
         Width           =   1530
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Continuar"
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   135
         Width           =   1530
      End
   End
   Begin TrueOleDBGrid60.TDBGrid tdbgIngreso 
      Align           =   3  'Align Left
      Height          =   5145
      Left            =   0
      OleObjectBlob   =   "ALFrmEntregaManual.frx":329A
      TabIndex        =   0
      Top             =   0
      Width           =   8925
   End
End
Attribute VB_Name = "ALFrmEntregaManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'--
Public QResp As Boolean
'--
Dim CodArt As String
Dim Item As String
Dim CantidadC As Long
Dim CantidadE As Long
Dim RsIngreso As ADODB.Recordset
Dim rsAux As ADODB.Recordset
Private Sub cmdAceptar_Click()
On Error GoTo QError
    If RsIngreso.EditMode <> adEditNone Then RsIngreso.Update
    If valida Then
        '--
        GlSqlAux = "DELETE FROM ALAuxEntDeIng WHERE Usuario = '" & NombreTerminal & "'"
        db.Execute GlSqlAux
        GlSqlAux = "INSERT INTO ALAuxEntDeIng(Usuario, CodArt, CodProveedor, IdIngreso, CantidadEntCaj, CantidadEntEj) " & _
                   "SELECT '" & NombreTerminal & "', CodArt, CodProveedor, IdIngreso, CantCaja, CantEjm FROM ALAuxEntregaManual WHERE Usuario = '" & NombreTerminal & "' AND CodArt = '" & CodArt & "' AND Elegido = 1"
        db.Execute GlSqlAux
        '--
        QResp = True
        Unload Me
    End If
    Exit Sub
QError:
    MsgBox err.Number & " : " & err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdModificar_Click()
    If RsIngreso.RecordCount <= 0 Then Beep: Exit Sub
    With ALFrmEntDet
        .estado = 2
        .IdIngreso = RsIngreso!IdIngreso
        .Codigo = CodArt
        .Item = Item
        .NoLici = RsIngreso!NoLicitacion
        .SaldoCaja = RsIngreso!SaldoCaja
        .SaldoEj = RsIngreso!SaldoEj
        .CantCaja = RsIngreso!CantCaja
        .CantEjm = RsIngreso!CantEjm
        .UnidadCaja = RsIngreso!UnidadCaja
        .Show vbModal
        If .QResp Then
            GlSqlAux = "UPDATE ALAuxEntregaManual " & _
                       "SET CantCaja = " & .CantCaja & ", " & _
                       "CantEjm = " & .CantEjm & ", " & _
                       "Elegido = " & IIf(.CantCaja > 0 And .CantEjm > 0, 1, 0) & " " & _
                       "WHERE Usuario = '" & NombreTerminal & "' AND IdIngreso = " & .IdIngreso & " AND CodArt = '" & .Codigo & "'"
            db.Execute GlSqlAux
            '--
            GlSqlAux = "SELECT * FROM ALAuxEntregaManual WHERE Usuario = '" & NombreTerminal & "' AND (SaldoCaja > 0 AND SaldoEj > 0) ORDER BY IdIngreso"
            '--
            Set RsIngreso = New ADODB.Recordset
            RsIngreso.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
            Set tdbgIngreso.DataSource = RsIngreso
            RsIngreso.Find "IdIngreso = " & .IdIngreso
            '--
            Totales
        End If
    End With
End Sub

Private Sub Form_Load()
    '--
    QResp = False
    '--
    With Me
        .Top = 0
        .Left = 0
    End With
    '--
    Set rsAux = New ADODB.Recordset
    GlSqlAux = "SELECT DescDetalle FROM ALCLDetalle WHERE CodGrupo + '-' + CodDetalle = '" & CodArt & "'"
    rsAux.Open GlSqlAux, db, adOpenStatic
    Item = rsAux!descdetalle
    Me.Caption = "Entrega Manual del Item - '" & CodArt & " : " & rsAux!descdetalle & "'"
    lbl.Caption = "Cantidad de Cajas a Entregar del Item = " & Format(CantidadC, "###,###,##0")
    '--
    Set tdbgIngreso.DataSource = RsIngreso
    Totales
	Call SeguridadSet(Me)
End Sub
Private Sub Form_Resize()
On Error Resume Next
    tdbgIngreso.Width = Me.ScaleWidth
End Sub
Public Sub Totales()
Dim rs As ADODB.Recordset
Dim SaldoCaja As Long
Dim SaldoEjm As Long
Dim CantCaja As Long
Dim CantEjm As Long
Dim UnidadCaja As Integer
    Set rs = New ADODB.Recordset
    Set rs = RsIngreso.Clone
    SaldoCaja = 0
    SaldoEjm = 0
    CantCaja = 0
    CantEjm = 0
    UnidadCaja = 0
    While Not rs.EOF
        SaldoCaja = SaldoCaja + IIf(IsNull(rs!SaldoCaja), 0, rs!SaldoCaja)
        SaldoEjm = SaldoEjm + IIf(IsNull(rs!SaldoEj), 0, rs!SaldoEj)
        UnidadCaja = UnidadCaja + IIf(IsNull(rs!UnidadCaja), 0, rs!UnidadCaja)
        If CBool(rs!Elegido) Then
            CantCaja = CantCaja + IIf(IsNull(rs!CantCaja), 0, rs!CantCaja)
            CantEjm = CantEjm + IIf(IsNull(rs!CantEjm), 0, rs!CantEjm)
        End If
        rs.MoveNext
    Wend
    UnidadCaja = UnidadCaja / CCur(IIf(rs.RecordCount <= 0, 1, rs.RecordCount))
    tdbgIngreso.Columns("NoLicitacion").FooterText = "SALDOS" & vbCrLf & "SEL."
    tdbgIngreso.Columns("SaldoCaja").FooterText = Format(SaldoCaja, "###,###,##0") & vbCrLf & " "
    tdbgIngreso.Columns("SaldoEj").FooterText = Format(SaldoEjm, "###,###,##0") & vbCrLf & " "
    tdbgIngreso.Columns("UnidadCaja").FooterText = Format(UnidadCaja, "###,###,##0.00") & vbCrLf & " "
    tdbgIngreso.Columns("CantCaja").FooterText = vbCrLf & Format(CantCaja, "###,###,##0")
    tdbgIngreso.Columns("CantEjm").FooterText = vbCrLf & Format(CantEjm, "###,###,##0")
End Sub
Public Sub ALPrincipal(QCodArt As String, QCantCaja As Long, QCantEjm As Long)
    '--
    CodArt = QCodArt
    CantidadC = QCantCaja
    CantidadE = QCantEjm
    '--
    GlSqlAux = "DELETE FROM ALAuxEntregaManual WHERE Usuario = '" & NombreTerminal & "'"
    db.Execute GlSqlAux
    '--
    GlSqlAux = "INSERT INTO ALAuxEntregaManual(Usuario, IdIngreso, CodArt, CodProveedor, UnidadCaja, NoLicitacion, SaldoCaja, SaldoEj) " & _
               "SELECT '" & NombreTerminal & "', IdIngreso, CodArt, CodProveedor, UnidadCaja, " & _
               "(SELECT NoLicitacion FROM ALIngresoAlm  WHERE IdIngreso = Ing.IdIngreso), " & _
               "SaldoCaja = ISNULL(CantidadCaj,0) - (SELECT ISNULL(SUM(CantidadEntCaj), 0) FROM ALEntDeIng WHERE (CodArt = Ing.CodArt) AND (IdIngreso = Ing.IdIngreso)), " & _
               "SaldoEj = ISNULL(CantidadEj,0) - (SELECT ISNULL(SUM(CantidadEntEj), 0) FROM ALEntDeIng WHERE (CodArt = Ing.CodArt) AND (IdIngreso = Ing.IdIngreso)) " & _
               "FROM AlIngresoAlmDet Ing WHERE CodArt = '" & QCodArt & "'"
    db.Execute GlSqlAux
    '--
    GlSqlAux = "SELECT *, p.Descripcion_Larga As NomProv FROM ALAuxEntregaManual a INNER JOIN ac_Proveedor p ON a.CodProveedor = p.Ruc_Id WHERE Usuario = '" & NombreTerminal & "' AND (SaldoCaja > 0 AND SaldoEj > 0) ORDER BY IdIngreso"
    '--
    Set RsIngreso = New ADODB.Recordset
    RsIngreso.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
    If RsIngreso.RecordCount <= 0 Then
        MsgBox "No hay existencias en Almacen del Item '" & QCodArt & "'.", vbExclamation + vbOKOnly, "Ate3nción"
        Exit Sub
    End If
    '--
    Me.Show vbModal
End Sub
Private Function valida() As Boolean
    valida = False
    If CLng(tdbgIngreso.Columns("CantCaja").FooterText) <> CantidadC Then  'Or
       'CLng(tdbgIngreso.Columns("Cantejm").FooterText) <> CantidadE Then
       ' "; Ejemplares = " & Format(tdbgIngreso.Columns("CantEjm").FooterText, "###,###,##0") & vbCrLf & _
       '"; Ejemplares = " & Format(CantidadE, "###,###,##0") & "", vbExclamation + vbOKOnly, "Atención"
       MsgBox "La Cantidad seleccionada para la entrega no es igual a la cantidad a Entregar." & vbCrLf & _
              "Cantidad Seleccionada : " & vbCrLf & vbTab & "Cajas = " & Format(tdbgIngreso.Columns("CantCaja").FooterText, "###,###,##0") & vbCrLf & _
              "Cantidad a Entregar : " & vbCrLf & vbTab & "Cajas = " & Format(CantidadC, "###,###,##0") & vbCrLf & "", vbExclamation + vbOKOnly, "Atención"
       tdbgIngreso.SetFocus
       Exit Function
    End If
    valida = True
End Function

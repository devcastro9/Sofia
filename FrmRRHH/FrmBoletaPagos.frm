VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmBoletaPagos 
   Caption         =   "Boletas de Pagos"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   12240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnGuardar 
      Caption         =   "GUARDAR"
      Height          =   495
      Left            =   9840
      TabIndex        =   21
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "CANCELAR"
      Height          =   495
      Left            =   11040
      TabIndex        =   20
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton btnEliminar 
      Caption         =   "ELIMINAR"
      Height          =   495
      Left            =   7920
      TabIndex        =   19
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton btnEditar 
      Caption         =   "EDITAR"
      Height          =   495
      Left            =   6720
      TabIndex        =   18
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton btnNuevo 
      Caption         =   "NUEVO"
      Height          =   495
      Left            =   5520
      TabIndex        =   17
      Top             =   7560
      Width           =   975
   End
   Begin VB.ComboBox cmbMes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FrmBoletaPagos.frx":0000
      Left            =   6240
      List            =   "FrmBoletaPagos.frx":0028
      TabIndex        =   16
      Text            =   "ENERO"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtMensaje 
      Height          =   2535
      Left            =   5520
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "FrmBoletaPagos.frx":0091
      Top             =   4920
      Width           =   6495
   End
   Begin VB.TextBox txtTitulo 
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Text            =   "TITULO DEL MENSAJE"
      Top             =   3960
      Width           =   6495
   End
   Begin VB.TextBox txtGestion 
      Height          =   375
      Left            =   10080
      TabIndex        =   11
      Text            =   "2020"
      Top             =   1680
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc adoListaMensaje 
      Height          =   330
      Left            =   120
      Top             =   8520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dgListaMensaje 
      Bindings        =   "FrmBoletaPagos.frx":009C
      Height          =   6735
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   11880
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "gestion"
         Caption         =   "Gestion"
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
      BeginProperty Column01 
         DataField       =   "mes"
         Caption         =   "Mes"
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
      BeginProperty Column02 
         DataField       =   "titulo"
         Caption         =   "Titulo"
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
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3420.284
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFechaRegistroMensaje 
      Caption         =   "31/12/2020"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblUsuarioMensaje 
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   12
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblMensaje 
      Caption         =   "Mensaje:"
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
      Left            =   5520
      TabIndex        =   10
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblTituloMensaje 
      Caption         =   "Titulo del mensaje:"
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
      Left            =   5520
      TabIndex        =   9
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label lblFechaRegistro 
      Caption         =   "Fecha de registro:"
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
      Left            =   5520
      TabIndex        =   8
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblUsuario 
      Caption         =   "Usuario:"
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
      Left            =   9000
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblMes 
      Caption         =   "Mes:"
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
      Left            =   5520
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblGestion 
      Caption         =   "Gestion:"
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
      Left            =   9000
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblIdMensaje 
      Caption         =   "100000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblID 
      Caption         =   "ID:"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblListaMensaje 
      Caption         =   "Lista de Mensajes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblTitulo 
      Caption         =   "REVERSO DE BOLETAS DE PAGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "FrmBoletaPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsListaMensaje As New ADODB.Recordset
Dim rsExiste As New ADODB.Recordset
Dim esNuevo As Boolean


Private Sub adoListaMensaje_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Not rsListaMensaje.BOF And Not rsListaMensaje.EOF Then
        lblIdMensaje.Caption = adoListaMensaje.Recordset!Id
        txtGestion.Text = adoListaMensaje.Recordset!gestion
        cmbMes.Text = mesaCadena(adoListaMensaje.Recordset!mes)
        lblUsuarioMensaje.Caption = adoListaMensaje.Recordset!usrCodigo
        lblFechaRegistroMensaje = adoListaMensaje.Recordset!fechaRegistro
        txtTitulo.Text = adoListaMensaje.Recordset!Titulo
        txtMensaje.Text = adoListaMensaje.Recordset!Mensaje
    End If
End Sub

Private Sub BtnCancelar_Click()
    If MsgBox("¿Cancelar la operacion?", vbExclamation + vbYesNo, "CANCELAR") = vbYes Then
        Call deshabCampos
        rsListaMensaje.MoveNext
        rsListaMensaje.MovePrevious
    End If
End Sub

Private Sub btnEditar_Click()
    esNuevo = False
    Call habCampos
    btnGuardar.Visible = True
    btnCancelar.Visible = True
End Sub

Private Sub btnEliminar_Click()
    If MsgBox("¿Esta seguro de eliminar el registro del Mensaje?", vbExclamation + vbYesNo, "ELIMINAR MENSAJE") = vbYes Then
        db.Execute "UPDATE rcMensajeBoletaPago_JASM SET estado = 'ANL' WHERE id = '" & lblIdMensaje.Caption & "'"
        Call leerMensajes
    End If
End Sub

Private Sub btnGuardar_Click()
    'borramos posibles espacios en blanco
    txtGestion.Text = LTrim(txtGestion.Text)
    txtGestion.Text = RTrim(txtGestion.Text)
    cmbMes.Text = LTrim(cmbMes.Text)
    cmbMes.Text = RTrim(cmbMes.Text)
    txtTitulo.Text = LTrim(txtTitulo.Text)
    txtTitulo.Text = RTrim(txtTitulo.Text)
    txtMensaje.Text = LTrim(txtMensaje.Text)
    txtMensaje.Text = RTrim(txtMensaje.Text)
    
    If validar Then
        If esNuevo Then
            If existe Then
                MsgBox "El mensaje ya existe", vbCritical, "GUARDADO"
            Else
                If MsgBox("¿Esta seguro de agregar el Mensaje?", vbExclamation + vbYesNo, "AGREGAR MENSAJE") = vbYes Then
                    db.Execute "INSERT INTO rcMensajeBoletaPago_JASM (gestion, mes, titulo, mensaje, usrCodigo, fechaRegistro) VALUES ('" & txtGestion.Text & _
                    "', " & mesaEntero(cmbMes.Text) & ", '" & txtTitulo.Text & "', '" & txtMensaje.Text & "', '" & glusuario & _
                    "', '" & ObtenerFechaServidor & "')"
                    MsgBox "Guardado correctamente", vbInformation, "GUARDADO"
                    Call deshabCampos
                    Call leerMensajes
                End If
            End If
        Else
            If existe Then
                MsgBox "El mensaje ya existe", vbCritical, "GUARDADO"
            Else
                If MsgBox("¿Esta seguro de editar el Mensaje?", vbExclamation + vbYesNo, "EDITAR MENSAJE") = vbYes Then
                    db.Execute "UPDATE rcMensajeBoletaPago_JASM SET gestion = '" & txtGestion.Text & "', mes = " & mesaEntero(cmbMes.Text) & _
                    ", titulo = '" & txtTitulo.Text & "', mensaje = '" & txtMensaje.Text & "', usrCodigo = '" & glusuario & _
                    "', nroModificacion = nroModificacion + 1, fechaModificacion = '" & ObtenerFechaServidor & _
                    "' WHERE id = " & CInt(lblIdMensaje.Caption)
                    MsgBox "Editado correctamente", vbInformation, "GUARDADO"
                    Call deshabCampos
                    Call leerMensajes
                End If
            End If
        End If
    End If
End Sub

Private Sub btnNuevo_Click()
    esNuevo = True
    Call habCampos
    btnGuardar.Visible = True
    btnCancelar.Visible = True
    
    'limpiamos campos
    lblIdMensaje.Caption = "0"
    txtGestion.Text = "2020"
    cmbMes.Text = "ENERO"
    lblUsuarioMensaje.Caption = glusuario
    lblFechaRegistroMensaje.Caption = ObtenerFechaServidor
    txtTitulo.Text = ""
    txtMensaje.Text = ""
End Sub

Private Sub Form_Load()
    Call leerMensajes
    Call deshabCampos
End Sub

Private Sub leerMensajes()
    Set rsListaMensaje = New ADODB.Recordset
    If rsListaMensaje.State = 1 Then rsListaMensaje.Close
    rsListaMensaje.Open "SELECT * FROM rcMensajeBoletaPago_JASM WHERE estado = 'APR' ORDER BY gestion DESC, mes", db, adOpenStatic
    Set adoListaMensaje.Recordset = rsListaMensaje
End Sub
Private Sub deshabCampos()
    txtGestion.Enabled = False
    cmbMes.Enabled = False
    txtTitulo.Enabled = False
    txtMensaje.Enabled = False
    btnCancelar.Visible = False
    btnGuardar.Visible = False
    btnNuevo.Visible = True
    btnEditar.Visible = True
    btnEliminar.Visible = True
End Sub
Private Sub habCampos()
    txtGestion.Enabled = True
    cmbMes.Enabled = True
    txtTitulo.Enabled = True
    txtMensaje.Enabled = True
    btnNuevo.Visible = False
    btnEditar.Visible = False
    btnEliminar.Visible = False
End Sub
Private Function validar() As Boolean
    validar = False
    If Not IsNumeric(txtGestion.Text) Then
        MsgBox "La gestion debe de ser un número", vbCritical, "ERROR EN LA GESTION"
        Exit Function
    End If
    If mesaEntero(cmbMes.Text) = 0 Then    'comprueba el mes correcto
        MsgBox "El mes no corresponde a ninguno de los doce meses!!!", vbCritical, "ERROR EN EL MES"
        Exit Function
    End If
    If Len(txtTitulo.Text) > 70 Then   'titulo no tan largo
        MsgBox "El titulo sobrepasa los 70 caracteres", vbCritical, "ERROR EN EL TITULO"
        Exit Function
    End If
    If Len(txtMensaje.Text) > 700 Then 'mensaje no tan larga
        MsgBox "El mensaje sobrepasa los 700 caracteres", vbCritical, "ERROR EN EL MENSAJE"
        Exit Function
    End If
    If Len(txtTitulo.Text) < 5 Then   'titulo no tan corto
        MsgBox "El titulo es muy corto", vbCritical, "ERROR EN EL TITULO"
        Exit Function
    End If
    If Len(txtMensaje.Text) < 20 Then 'mensaje no tan corto
        MsgBox "El mensaje es muy corto", vbCritical, "ERROR EN EL MENSAJE"
        Exit Function
    End If
    validar = True
End Function
Private Function existe() As Boolean 'si existe algun registro similar
    existe = True
    Set rsExiste = New ADODB.Recordset
    If rsExiste.State = 1 Then rsExiste.Close
    rsExiste.Open "SELECT * FROM rcMensajeBoletaPago_JASM WHERE gestion = '" & txtGestion.Text & _
    "' AND mes = '" & mesaEntero(cmbMes.Text) & "' AND id <> '" & lblIdMensaje.Caption & "' AND estado = 'APR'", db, adOpenStatic
    If rsExiste.RecordCount = 0 Then
        existe = False
    End If
    rsExiste.Close
End Function
Public Function mesaCadena(numMes As Integer) As String
    Select Case numMes
        Case 1
            mesaCadena = "ENERO"
        Case 2
            mesaCadena = "FEBRERO"
        Case 3
            mesaCadena = "MARZO"
        Case 4
            mesaCadena = "ABRIL"
        Case 5
            mesaCadena = "MAYO"
        Case 6
            mesaCadena = "JUNIO"
        Case 7
            mesaCadena = "JULIO"
        Case 8
            mesaCadena = "AGOSTO"
        Case 9
            mesaCadena = "SEPTIEMBRE"
        Case 10
            mesaCadena = "OCTUBRE"
        Case 11
            mesaCadena = "NOVIEMBRE"
        Case 12
            mesaCadena = "DICIEMBRE"
        Case Else
            mesaCadena = "ERROR"
    End Select
End Function
Public Function mesaEntero(nomMes As String) As Integer
    Select Case nomMes
        Case "ENERO"
            mesaEntero = 1
        Case "FEBRERO"
            mesaEntero = 2
        Case "MARZO"
            mesaEntero = 3
        Case "ABRIL"
            mesaEntero = 4
        Case "MAYO"
            mesaEntero = 5
        Case "JUNIO"
            mesaEntero = 6
        Case "JULIO"
            mesaEntero = 7
        Case "AGOSTO"
            mesaEntero = 8
        Case "SEPTIEMBRE"
            mesaEntero = 9
        Case "OCTUBRE"
            mesaEntero = 10
        Case "NOVIEMBRE"
            mesaEntero = 11
        Case "DICIEMBRE"
            mesaEntero = 12
        Case Else
            mesaEntero = 0
    End Select
End Function

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form aw_tipo_compra_venta 
   BackColor       =   &H00000000&
   Caption         =   "Clasificadores - Tipo Compra Venta"
   ClientHeight    =   7335
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   14835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   14835
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   18
      Top             =   0
      Width           =   20280
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10800
         Picture         =   "aw_tipo_compra_venta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Picture         =   "aw_tipo_compra_venta.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "aw_tipo_compra_venta.frx":064C
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   25
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1305
         Picture         =   "aw_tipo_compra_venta.frx":0E0B
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   24
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "aw_tipo_compra_venta.frx":1720
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   23
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6960
         Picture         =   "aw_tipo_compra_venta.frx":1E6C
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   22
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4200
         Picture         =   "aw_tipo_compra_venta.frx":269F
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   21
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5520
         Picture         =   "aw_tipo_compra_venta.frx":2E54
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "aw_tipo_compra_venta.frx":3721
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   19
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRONOGRAMA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   12855
         TabIndex        =   28
         Top             =   195
         Width           =   1815
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      FillStyle       =   2  'Horizontal Line
      ForeColor       =   &H80000008&
      Height          =   676
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   20280
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         Picture         =   "aw_tipo_compra_venta.frx":3EE3
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   16
         Top             =   0
         Width           =   1335
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "aw_tipo_compra_venta.frx":46B9
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   15
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENTAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   13215
         TabIndex        =   17
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFC0&
      Height          =   6000
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   7935
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   5520
         Width           =   7665
         _ExtentX        =   13520
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
         BackColor       =   16777215
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
         Caption         =   " <-- Inicio                                        Tipo Venta Compra                                            Fin -->"
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
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   5175
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   9128
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
            DataField       =   "venta_tipo"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "venta_tipo_descripcion"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4589.858
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fra_ABM 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Left            =   8160
      TabIndex        =   8
      Top             =   720
      Width           =   6525
      Begin VB.TextBox txt_ventatipo 
         DataField       =   "venta_tipo"
         DataSource      =   "Ado_datos"
         Height          =   270
         Left            =   360
         TabIndex        =   29
         Text            =   "-"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txt_descripcion 
         DataField       =   "venta_tipo_descripcion"
         DataSource      =   "Ado_datos"
         Height          =   270
         Left            =   360
         TabIndex        =   0
         Text            =   "-"
         Top             =   1755
         Width           =   5775
      End
      Begin VB.Label lbl_inicial 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
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
         Height          =   240
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label lbl_codigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
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
         Height          =   240
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Txt_estado 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REG"
         DataField       =   "estado_codigo"
         DataSource      =   "Ado_datos"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Registro"
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
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   2520
         Width           =   1695
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   14835
      TabIndex        =   2
      Top             =   7335
      Width           =   14835
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   7
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   2400
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   6720
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
      Caption         =   "Ado_datos1"
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
End
Attribute VB_Name = "aw_tipo_compra_venta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
Dim rs_datos1 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim queryinicial As String

Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If rs_datos!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "APR"
         rs_datos!Fecha_Registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado (ERR) o Aprobado (APR) anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnBuscar_Click()
    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexión = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = dg_datos
    ClBuscaGrid.QueryUtilizado = queryinicial
    Set ClBuscaGrid.RecordsetTrabajo = rs_datos
    'ClBuscaGrid.CamposVisibles = "11010011"
    ClBuscaGrid.Ejecutar
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
       If txt_codigo.Text = "" Then
        rs_datos.Delete ' En proceso nuevo.
       Else
        rs_datos.CancelUpdate ' En proceso de edicion de registro.
       End If

        Call ABRIR_TABLA
        rs_datos.MoveFirst
        mbDataChanged = False
        Fra_ABM.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True

    End If
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
   
   If rs_datos!estado_codigo = "REG" Then
      MsgBox "No se puede Anular un registro que tenga el estado " + rs_datos!estado_codigo
      Exit Sub
   Else
        sino = MsgBox(" ¿Desea anular el registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "ERR"
         rs_datos!Fecha_Registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   End If
   
   Call ABRIR_TABLA
   Exit Sub

UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnDesAprobar_Click()
  On Error GoTo UpdateErr
   If rs_datos!estado_codigo = "APR" Or rs_datos!estado_codigo = "ERR" Or rs_datos!estado_codigo = "ANL" Then
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "REG"
         rs_datos!Fecha_Registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Elaborado (REG) ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  'Call valida_campos
  If Len(Trim((txt_ventatipo.Text))) > 1 Then
     MsgBox "La clave solo debe contener un caracter."
     Exit Sub
  End If
  
  If ExisteReg(Ado_datos.Recordset!venta_tipo) And VAR_SW = "ADD" Then
      MsgBox "El codigo '" + Ado_datos.Recordset!venta_tipo + "' ya existe."
      Exit Sub
  Else
      If VAR_VAL = "OK" Then
            If VAR_SW = "ADD" Then
                rs_datos!estado_codigo = "REG"  ' no cambia
        
             End If
         rs_datos!venta_tipo = txt_ventatipo.Text
         rs_datos!venta_tipo_descripcion = Txt_descripcion.Text
        
         rs_datos!Fecha_Registro = Date     ' no cambia
         rs_datos!usr_codigo = glusuario    ' no cambia
         rs_datos.UpdateBatch adAffectAll
        
         Call ABRIR_TABLA
         rs_datos.MoveLast
         mbDataChanged = False
          
          Fra_ABM.Enabled = False
          fraOpciones.Visible = True
          FraGrabarCancelar.Visible = False
          dg_datos.Enabled = True
         
    End If
  End If
  
  
  
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  Dim inicial As Integer
  If txt_parametroinicial.Text = "" Then
    MsgBox "Debe registrar la " + lbl_inicial.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  Else
       inicial = CInt(txt_parametroinicial.Text)
       If inicial < 0 Then
          MsgBox lbl_inicial.Caption + " debe ser mayor a cero.", vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
       End If
  End If
  
  Dim final As Integer
  If txt_parametrofinal.Text = "" Then
    MsgBox "Debe registrar la " + lbl_final.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  Else
        final = CInt(txt_parametrofinal.Text)
        If final < inicial Then
          MsgBox lbl_final.Caption + " debe ser mayor a " + lbl_inicial.Caption, vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
       End If
  End If
  
  Dim porcentaje As Integer
  If txt_antiguedad.Text = "" Then
    MsgBox "Debe registrar la " + lbl_porcentaje.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  Else
       porcentaje = CInt(txt_antiguedad.Text)
       If porcentaje < 0 Or porcentaje > 100 Then
          MsgBox lbl_porcentaje.Caption + " debe estar dentro del rango 0 - 100 ", vbCritical + vbExclamation, "Validación de datos"
          VAR_VAL = "ERR"
          Exit Sub
       End If
  End If

End Sub

Private Sub BtnImprimir_Click()
  Dim iResult As Integer
  cr01.WindowShowPrintSetupBtn = True
  cr01.WindowShowRefreshBtn = True
  cr01.ReportFileName = App.Path & "\REPORTES\clasificadores\gr_modalidad_contratacion.rpt"
  iResult = cr01.PrintReport
  If iResult <> 0 Then
      MsgBox cr01.LastErrorNumber & " : " & cr01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
  cr01.WindowState = crptMaximized
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
  If rs_datos!estado_codigo = "ERR" Or rs_datos!estado_codigo = "APR" Then
      MsgBox "No se puede MODIFICAR un registro Errado (ERR) o Aprobado (APR)...", vbExclamation, "Validación de Registro"
  Else
'  lblStatus.Caption = "Modificar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "MOD"
    
  End If
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnSalir_Click()

  Unload Me
End Sub

Private Sub DtcUE_Click(Area As Integer)
    DtcUE_Des.BoundText = DtcUE.BoundText
End Sub

Private Sub DtcUE_Des_Click(Area As Integer)
    DtcUE.BoundText = DtcUE_Des.BoundText
End Sub


Private Sub Form_Load()
    Call ABRIR_TABLA
    
    mbDataChanged = False
    Fra_ABM.Enabled = False
    dg_datos.Enabled = True

	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLA()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select * from ac_tipo_compra_venta "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
      Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
End Sub

Private Sub Ado_datos_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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

Private Sub BtnAñadir_Click()
  On Error GoTo AddErr
    Call ABRIR_TABLA
    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    rs_datos.AddNew
    'lblStatus.Caption = "Agregar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "ADD"
    txt_ventatipo.SetFocus
    txt_ventatipo.Text = ""
    Txt_descripcion.Text = ""
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  rs_datos.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub


Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txt_observacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_parametro_mayor_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_parametro_menor_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function ExisteReg(Codigo As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ac_tipo_compra_venta WHERE venta_tipo = '" & Codigo & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

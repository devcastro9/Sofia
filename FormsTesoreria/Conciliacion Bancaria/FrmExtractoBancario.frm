VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmExtractoBancario 
   Caption         =   "Actualización del Extracto Bancario"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   Icon            =   "FrmExtractoBancario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame FraBusca 
      Height          =   2085
      Left            =   2460
      TabIndex        =   37
      Top             =   3555
      Visible         =   0   'False
      Width           =   2040
      Begin VB.CommandButton CmdSalirDoc 
         Caption         =   "Salir"
         Height          =   375
         Left            =   225
         TabIndex        =   42
         Top             =   1485
         Width           =   1515
      End
      Begin VB.TextBox TxtGes 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3615
         TabIndex        =   41
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox TxtOrg 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2047
         TabIndex        =   40
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox TxtDoc 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   225
         TabIndex        =   39
         Top             =   645
         Width           =   1515
      End
      Begin VB.CommandButton CmdBuscaDoc 
         Caption         =   "Buscar"
         Height          =   390
         Left            =   225
         TabIndex        =   38
         Top             =   1095
         Width           =   1515
      End
      Begin VB.Label Label24 
         Caption         =   "Gestión"
         Height          =   165
         Left            =   3900
         TabIndex        =   45
         Top             =   645
         Width           =   795
      End
      Begin VB.Label Label23 
         Caption         =   "Organismo"
         Height          =   165
         Left            =   2310
         TabIndex        =   44
         Top             =   615
         Width           =   990
      End
      Begin VB.Label Label22 
         Caption         =   "Cmpte. Inicial"
         Height          =   165
         Left            =   450
         TabIndex        =   43
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   9120
      TabIndex        =   0
      Top             =   0
      Width           =   9180
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDAD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   60
         TabIndex        =   5
         Top             =   675
         Width           =   1110
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   4
         Top             =   705
         Width           =   2460
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9165
         TabIndex        =   3
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label LblUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   2
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXTRACTO BANCARIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   360
         Left            =   4395
         TabIndex        =   1
         Top             =   135
         Width           =   3555
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   0
         Picture         =   "FrmExtractoBancario.frx":0ECA
         Top             =   0
         Width           =   11640
      End
   End
   Begin VB.Frame FraOpciones 
      Height          =   6135
      Left            =   15
      TabIndex        =   6
      Top             =   1035
      Width           =   1545
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "Ac&tualizar"
         Height          =   780
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3435
         Width           =   990
      End
      Begin VB.CommandButton cmdAdicionar 
         Appearance      =   0  'Flat
         Caption         =   "&Adicionar"
         Height          =   780
         Left            =   240
         Picture         =   "FrmExtractoBancario.frx":27F3A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   315
         Width           =   990
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   780
         Left            =   240
         Picture         =   "FrmExtractoBancario.frx":2837C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4215
         Width           =   990
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "B&orrar"
         Height          =   780
         Left            =   240
         Picture         =   "FrmExtractoBancario.frx":289E6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1875
         Width           =   990
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "B&uscar"
         Height          =   780
         Left            =   240
         Picture         =   "FrmExtractoBancario.frx":29050
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2655
         Width           =   990
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Modificar"
         Height          =   780
         Left            =   240
         Picture         =   "FrmExtractoBancario.frx":29492
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1095
         Width           =   990
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   780
         Left            =   240
         Picture         =   "FrmExtractoBancario.frx":298D4
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4995
         Width           =   990
      End
      Begin VB.Image Image3 
         Height          =   5970
         Left            =   60
         Picture         =   "FrmExtractoBancario.frx":29D16
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1440
      End
   End
   Begin VB.Frame FraGrabarCancelar 
      Height          =   5925
      Left            =   30
      TabIndex        =   13
      Top             =   1050
      Width           =   1500
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   780
         Left            =   255
         Picture         =   "FrmExtractoBancario.frx":2BE38
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1035
         Width           =   990
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   780
         Left            =   255
         Picture         =   "FrmExtractoBancario.frx":2C142
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   255
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   1590
      TabIndex        =   16
      Top             =   1035
      Width           =   10140
      Begin MSDataGridLib.DataGrid DtGDatosBanco 
         Height          =   5925
         Left            =   45
         TabIndex        =   17
         Top             =   135
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   10451
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame FraPrincipal 
         Enabled         =   0   'False
         Height          =   5955
         Left            =   4350
         TabIndex        =   18
         Top             =   105
         Width           =   5715
         Begin MSAdodcLib.Adodc AdoCuenta 
            Height          =   330
            Left            =   3645
            Top             =   2745
            Visible         =   0   'False
            Width           =   1965
            _ExtentX        =   3466
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
            Caption         =   "Adodc1"
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
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            Height          =   1230
            Left            =   3810
            TabIndex        =   30
            Top             =   465
            Width           =   1725
            Begin VB.OptionButton OptEgresos 
               Caption         =   "Egreso"
               Height          =   360
               Left            =   195
               TabIndex        =   32
               Top             =   750
               Width           =   1335
            End
            Begin VB.OptionButton OptIngresos 
               Caption         =   "Ingreso"
               Height          =   375
               Left            =   180
               Picture         =   "FrmExtractoBancario.frx":2C584
               TabIndex        =   31
               Top             =   270
               Width           =   1365
            End
            Begin VB.Image Image4 
               Enabled         =   0   'False
               Height          =   5775
               Left            =   0
               Picture         =   "FrmExtractoBancario.frx":11279E
               Stretch         =   -1  'True
               Top             =   0
               Width           =   5640
            End
         End
         Begin VB.TextBox TxtNroDoc 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   1485
            TabIndex        =   22
            Top             =   1425
            Width           =   2220
         End
         Begin VB.TextBox TxtMonto 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   1485
            TabIndex        =   21
            Top             =   2040
            Width           =   2220
         End
         Begin VB.TextBox TxtJustificacion 
            Appearance      =   0  'Flat
            Height          =   855
            Left            =   1485
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   4260
            Width           =   4020
         End
         Begin VB.TextBox TxtBanco 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   1485
            TabIndex        =   19
            Top             =   5325
            Width           =   4020
         End
         Begin MSComCtl2.DTPicker DTPFecha 
            Height          =   375
            Left            =   1470
            TabIndex        =   33
            Top             =   750
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   248643585
            CurrentDate     =   36413
         End
         Begin MSDataListLib.DataCombo DtCCuentaOrigen 
            Bindings        =   "FrmExtractoBancario.frx":122421
            DataField       =   "cta_codigo"
            Height          =   315
            Left            =   1500
            TabIndex        =   34
            Top             =   2655
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "cta_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCDescripcion 
            Bindings        =   "FrmExtractoBancario.frx":122439
            DataField       =   "cta_codigo"
            Height          =   315
            Left            =   1500
            TabIndex        =   35
            Top             =   3360
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "Cta_descripcion_larga"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCTgn 
            Bindings        =   "FrmExtractoBancario.frx":122451
            DataField       =   "cta_codigo"
            Height          =   315
            Left            =   1500
            TabIndex        =   36
            Top             =   3015
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "Cta_codigo_tgn"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Pago"
            Height          =   360
            Left            =   255
            TabIndex        =   28
            Top             =   870
            Width           =   1470
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta "
            Height          =   360
            Left            =   225
            TabIndex        =   27
            Top             =   2700
            Width           =   1470
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Cheque"
            Height          =   360
            Left            =   255
            TabIndex        =   26
            Top             =   1425
            Width           =   1470
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Monto"
            Height          =   360
            Left            =   255
            TabIndex        =   25
            Top             =   2055
            Width           =   1470
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Justificación"
            Height          =   360
            Left            =   255
            TabIndex        =   24
            Top             =   4155
            Width           =   1470
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   " Banco"
            Height          =   360
            Left            =   240
            TabIndex        =   23
            Top             =   5250
            Width           =   1470
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   5775
            Left            =   45
            Picture         =   "FrmExtractoBancario.frx":122469
            Stretch         =   -1  'True
            Top             =   135
            Width           =   5640
         End
      End
   End
End
Attribute VB_Name = "FrmExtractoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBANCO As New ADODB.Recordset
Public Bandera_Modificar As Integer

Private Sub CmdActualizar_Click()
  Set rsBANCO = New ADODB.Recordset
    If rsBANCO.State = 1 Then rsBANCO.Close
    rsBANCO.Open "select fecha_pago,nro_doc,monto,cta_codigo,justificacion,* from fc_DatosBanco order by Nro_cmpte", db, adOpenStatic, adLockOptimistic
    If rsBANCO.RecordCount > 0 Then
     Set DtGDatosBanco.DataSource = rsBANCO
    End If
End Sub

Private Sub cmdadicionar_Click()
    rsBANCO.AddNew
    FraPrincipal.Enabled = True
    FraGrabarCancelar.Visible = True
    FraOpciones.Visible = False
    
    'colocando en blanco los controles
    dtpfecha.Value = Empty
    TxtNroDoc.Text = Empty
    txtmonto.Text = Empty
    DtCCuentaOrigen.Text = Empty
    TxtJustificacion.Text = Empty
    
    Bandera_Modificar = 0
    
End Sub

Private Sub CmdBorrar_Click()
Dim Resp As String
Resp = MsgBox("Esta seguro de eliminar Si o No", vbYesNo, "Validación de datos")
If Resp = vbYes Then
    db.Execute "delete from fc_DatosBanco where fecha_pago='" & dtpfecha.Value & "' and cta_codigo='" & DtCCuentaOrigen.Text & "' and nro_doc='" & TxtNroDoc.Text & "'"
    cmdCancelar_Click
End If
End Sub

Private Sub CmdBuscaDoc_Click()
  Set rsBANCO = New ADODB.Recordset
  rsBANCO.Open "select fecha_pago,nro_doc,monto,cta_codigo,justificacion,* from fc_DatosBanco where nro_doc='" & TxtDoc.Text & "'order by Nro_cmpte", db, adOpenStatic, adLockOptimistic
  If rsBANCO.RecordCount > 0 Then
     Set DtGDatosBanco.DataSource = rsBANCO
  End If
   
End Sub

Private Sub CmdBuscar_Click()
    FraBusca.Visible = True
End Sub

Private Sub cmdCancelar_Click()
   
    Set rsBANCO = New ADODB.Recordset
    If rsBANCO.State = 1 Then rsBANCO.Close
    rsBANCO.Open "select fecha_pago,nro_doc,monto,cta_codigo,justificacion,* from fc_DatosBanco order by Nro_cmpte", db, adOpenStatic, adLockOptimistic
    If rsBANCO.RecordCount > 0 Then
     Set DtGDatosBanco.DataSource = rsBANCO
    End If

    FraPrincipal.Enabled = False
    FraGrabarCancelar.Visible = False
    FraOpciones.Visible = True
End Sub

Private Sub Cmdeditar_Click()
    FraPrincipal.Enabled = True
    FraGrabarCancelar.Visible = True
    FraOpciones.Visible = False
    Bandera_Modificar = 1
End Sub

Private Sub CmdGrabar_Click()


   'Abriendo Tabla  de registros del Banco
   If Bandera_Modificar <> 1 Then
        Set rsVerificaBanco = New ADODB.Recordset
        rsVerificaBanco.Open "select * from fc_DatosBanco where fecha_pago='" & dtpfecha.Value & "' and nro_doc='" & TxtNroDoc.Text & "' and cta_codigo='" & DtCCuentaOrigen.Text & "' order by Nro_cmpte", db, adOpenStatic, adLockOptimistic
        If rsVerificaBanco.RecordCount > 0 Then
            MsgBox "Registro se encuentra en la base " & dtpfecha.Value & " " & TxtNroDoc.Text & " " & DtCCuentaOrigen.Text, vbCritical + vbDefaultButton1, "Validación de datos"
            Exit Sub
        End If
   End If

    rsBANCO("fecha_pago") = Format(dtpfecha.Value, "dd/mm/yyyy")
    rsBANCO("nro_doc") = TxtNroDoc.Text
    If OptIngresos.Value = True Then
        rsBANCO("monto") = CDbl(txtmonto.Text) * (-1)
    End If
    If OptEgresos.Value = True Then
        rsBANCO("monto") = CDbl(txtmonto.Text) * (-1)
    End If
    rsBANCO("cta_codigo") = DtCCuentaOrigen.Text
    rsBANCO("justificacion") = TxtJustificacion.Text
    rsBANCO("Bco_codigo") = 1 'TxtBanco.Text
    rsBANCO.Update
    'db.Execute "insert into fc_datosbanco(fecha_pago,nro_doc,monto, cta_codigo,justificacion, bco_codigo) " & _
    '           " values ('" & Format(dtpfecha.value, "dd/mm/yyyy") & "', '" & TxtNroDoc.Text & "', " & CDbl(TxtMonto.Text) & ", '" & DtCCuentaOrigen.Text & "', '" & TxtJustificacion.Text & "', 1)"
    cmdCancelar_Click
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdSalirDoc_Click()
    FraBusca.Visible = False
End Sub

Private Sub DtCCuentaOrigen_Click(Area As Integer)
    DtCDescripcion.BoundText = DtCCuentaOrigen.BoundText
    DtCTgn.BoundText = DtCCuentaOrigen.BoundText
End Sub

Private Sub DtCDescripcion_Click(Area As Integer)
   DtCTgn.BoundText = DtCDescripcion.BoundText
   DtCCuentaOrigen.BoundText = DtCDescripcion.BoundText
End Sub

Private Sub DtCTgn_Click(Area As Integer)
    DtCDescripcion.BoundText = DtCTgn.BoundText
    DtCCuentaOrigen.BoundText = DtCTgn.BoundText
End Sub

Private Sub DtGDatosBanco_Click()
Dim rsc As New ADODB.Recordset
    dtpfecha.Value = rsBANCO("fecha_pago")
    TxtNroDoc.Text = rsBANCO("Nro_Doc")
    If rsBANCO("monto") >= 0 Then
        OptIngresos.Value = True
    End If
    If rsBANCO("monto") < 0 Then
        OptEgresos.Value = True
    End If
    txtmonto.Text = rsBANCO("monto")
    
    TxtJustificacion.Text = rsBANCO("justificacion")
    If rsBANCO("bco_codigo") = 1 Then
        TxtBanco.Text = "BANCO UNION"
    End If
    If rsBANCO("bco_codigo") = 2 Then
        TxtBanco.Text = "BANCO CENTRAL DE BOLIVIA"
    End If
    DtCCuentaOrigen.Text = rsBANCO("cta_codigo")
    If rsc.State = 1 Then rsc.Close
    rsc.Open "SELECT * FROM fc_cuenta_bancaria WHERE cta_codigo='" & DtCCuentaOrigen.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsc.RecordCount > 0 Then
        DtCTgn.Text = rsc("Cta_codigo_tgn")
        DtCDescripcion.Text = rsc("Cta_descripcion_larga")
    End If
    
End Sub

Private Sub Form_Load()
'Abriendo Tabla  de registros del Banco
  Set rsBANCO = New ADODB.Recordset
  rsBANCO.Open "select fecha_pago,nro_doc,monto,cta_codigo,justificacion,* from fc_DatosBanco order by Nro_cmpte", db, adOpenStatic, adLockOptimistic
  If rsBANCO.RecordCount > 0 Then
     Set DtGDatosBanco.DataSource = rsBANCO
  End If
  
  
  'Determinar las cuentas
  Set rscuenta = New ADODB.Recordset
  rscuenta.Open "select * from fc_cuenta_bancaria order by Cta_codigo_tgn", db, adOpenKeyset, adLockOptimistic
  Set AdoCuenta.Recordset = rscuenta

End Sub



Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii < 58 And KeyAscii > 47 Or KeyAscii = 8 Then
      Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If

End Sub
Private Sub TxtNroDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii < 58 And KeyAscii > 47 Or KeyAscii = 8 Then
      Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

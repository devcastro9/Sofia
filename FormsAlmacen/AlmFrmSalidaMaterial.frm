VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Begin VB.Form AlmFrmSalidaMaterial 
   Caption         =   "Almacen"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12060
   Icon            =   "AlmFrmSalidaMaterial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12060
      TabIndex        =   31
      Top             =   7275
      Width           =   12060
      Begin VB.Frame Frame1 
         Height          =   60
         Left            =   1215
         TabIndex        =   32
         Top             =   255
         Width           =   7290
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entrega de Almacen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   2
         Left            =   8610
         TabIndex        =   33
         Top             =   90
         Width           =   3120
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entrega de Almacen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   8625
         TabIndex        =   34
         Top             =   105
         Width           =   3120
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4170
      Left            =   2520
      TabIndex        =   19
      Top             =   2880
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7355
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Items"
      TabPicture(0)   =   "AlmFrmSalidaMaterial.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBGDet"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "AdoDetalle"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tdbdMat"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tdbdIngreso"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDetalle(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdDetalle(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdDetalle(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Agregar Detalle"
         Height          =   360
         Index           =   0
         Left            =   480
         TabIndex        =   14
         Top             =   3375
         Width           =   2355
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Modificar Detalle"
         Height          =   360
         Index           =   1
         Left            =   2970
         TabIndex        =   15
         Top             =   3375
         Width           =   2355
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Eliminar Detalle"
         Height          =   360
         Index           =   2
         Left            =   5460
         TabIndex        =   16
         Top             =   3375
         Width           =   2355
      End
      Begin TrueOleDBGrid60.TDBDropDown tdbdIngreso 
         Height          =   915
         Left            =   360
         OleObjectBlob   =   "AlmFrmSalidaMaterial.frx":0EE6
         TabIndex        =   37
         Top             =   720
         Width           =   2355
      End
      Begin TrueOleDBGrid60.TDBDropDown tdbdMat 
         Height          =   915
         Left            =   375
         OleObjectBlob   =   "AlmFrmSalidaMaterial.frx":4106
         TabIndex        =   18
         Top             =   735
         Width           =   2355
      End
      Begin MSAdodcLib.Adodc AdoDetalle 
         Height          =   330
         Left            =   5655
         Top             =   2730
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
      Begin TrueOleDBGrid60.TDBGrid TDBGDet 
         Bindings        =   "AlmFrmSalidaMaterial.frx":6702
         Height          =   3225
         Left            =   75
         OleObjectBlob   =   "AlmFrmSalidaMaterial.frx":671B
         TabIndex        =   17
         Top             =   75
         Width           =   8130
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   990
      Left            =   0
      Picture         =   "AlmFrmSalidaMaterial.frx":BBFB
      ScaleHeight     =   930
      ScaleWidth      =   12000
      TabIndex        =   23
      Top             =   0
      Width           =   12060
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "Verificar Entrega"
         Height          =   855
         Left            =   10380
         Picture         =   "AlmFrmSalidaMaterial.frx":EE95
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   30
         Width           =   1425
      End
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE ENTREGA DE MATERIAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   0
         Left            =   30
         TabIndex        =   35
         Top             =   0
         Width           =   6075
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "."
         Height          =   180
         Left            =   4815
         TabIndex        =   28
         Top             =   675
         Width           =   2655
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         Height          =   195
         Left            =   1035
         TabIndex        =   27
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   75
         TabIndex        =   26
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   195
         Left            =   1035
         TabIndex        =   25
         Top             =   675
         Width           =   2310
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   75
         TabIndex        =   24
         Top             =   660
         Width           =   735
      End
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE ENTREGA DE MATERIAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   1
         Left            =   60
         TabIndex        =   36
         Top             =   15
         Width           =   6075
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   0
         Picture         =   "AlmFrmSalidaMaterial.frx":F2D7
         Top             =   0
         Width           =   11640
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TdbgEntrega 
      Height          =   5880
      Left            =   0
      OleObjectBlob   =   "AlmFrmSalidaMaterial.frx":36347
      TabIndex        =   8
      Top             =   1050
      Width           =   2475
   End
   Begin VB.Frame FraDatos 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   2520
      TabIndex        =   20
      Top             =   1080
      Width           =   8295
      Begin VB.TextBox TexTFrenSer 
         DataField       =   "NoEntrega"
         Height          =   300
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   42
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtTipoCambio 
         Alignment       =   1  'Right Justify
         DataField       =   "TipoCambio"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         Height          =   300
         Left            =   6360
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1455
         Width           =   1125
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   120
         Width           =   2010
      End
      Begin TrueOleDBList60.TDBCombo tdbcDestino 
         DataField       =   "CodDestino"
         Height          =   45
         Left            =   3360
         OleObjectBlob   =   "AlmFrmSalidaMaterial.frx":39797
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   2670
      End
      Begin VB.TextBox TxtDescripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Obs"
         Height          =   690
         Left            =   195
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   450
         Width           =   7905
      End
      Begin VB.TextBox txtNoEntrega 
         DataField       =   "NoEntrega"
         Height          =   300
         Left            =   195
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1455
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpFechaEntrega 
         DataField       =   "FechaEnt"
         Height          =   300
         Left            =   2065
         TabIndex        =   11
         Top             =   1455
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   17039361
         CurrentDate     =   36737
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   6360
         TabIndex        =   39
         Top             =   1215
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   195
         TabIndex        =   30
         Top             =   150
         Width           =   840
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Entrega"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   195
         TabIndex        =   29
         Top             =   1215
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         Height          =   195
         Left            =   3635
         TabIndex        =   22
         Top             =   1215
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   2065
         TabIndex        =   21
         Top             =   1215
         Width           =   450
      End
   End
   Begin VB.Frame FraOpciones 
      Height          =   6315
      Left            =   10800
      TabIndex        =   40
      Top             =   945
      Width           =   1140
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   855
         Left            =   120
         Picture         =   "AlmFrmSalidaMaterial.frx":3B6B3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5310
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   855
         Left            =   120
         Picture         =   "AlmFrmSalidaMaterial.frx":3B8BD
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4470
         Width           =   855
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Borrar"
         Height          =   855
         Left            =   120
         Picture         =   "AlmFrmSalidaMaterial.frx":3BAC7
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3630
         Width           =   855
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   855
         Left            =   120
         Picture         =   "AlmFrmSalidaMaterial.frx":3C1B1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2790
         Width           =   855
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   855
         Left            =   120
         Picture         =   "AlmFrmSalidaMaterial.frx":3C3BB
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1950
         Width           =   855
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "Modificar"
         Height          =   855
         Left            =   120
         Picture         =   "AlmFrmSalidaMaterial.frx":3C5C5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1110
         Width           =   855
      End
      Begin VB.CommandButton CmdAnadir 
         Caption         =   "Adicionar"
         Height          =   855
         Left            =   120
         MousePointer    =   4  'Icon
         Picture         =   "AlmFrmSalidaMaterial.frx":3C7CF
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Label lblEstadoRs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ingreso a Almacen"
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
      Height          =   300
      Left            =   0
      TabIndex        =   41
      Top             =   6930
      Width           =   2475
   End
End
Attribute VB_Name = "AlmFrmSalidaMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dll
Private Const ERROR_BAD_FORMAT = 11&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&

Private Const SW_SHOWNORMAL = 1
Private Const SE_ERR_ACCESSDENIED = 5
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_OOM = 8
Private Const SE_ERR_SHARE = 26
Private Const SE_ERR_NOASSOC = 31
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'--
'JQA 04/2008
'Dim WithEvents RsEntrega As ADODB.Recordset
Dim RsEntrega As New ADODB.Recordset
Attribute RsEntrega.VB_VarHelpID = -1
Dim RsDestino As ADODB.Recordset
Dim rsNada As ADODB.Recordset
Dim RsDet As ADODB.Recordset
Dim RsMat As ADODB.Recordset
Dim RsIngreso As ADODB.Recordset
'--------
Dim estado As Integer ' 0 navegar, 1 Agregar, 2 Editar
Dim NoEntrega As Integer
Dim CodArt As String
Dim CadenaError As String
'--
'jqa
'Dim ClBuscaGrid As  ClBuscaEnGridPropio
Public Sub ALPrincipal(QEstado As Integer)
    '
    Screen.MousePointer = vbHourglass
    estado = QEstado
    '
    Select Case estado
        Case 0
        
'            Set rstrc_personalSoli = New ADODB.Recordset
'    If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'    rstrc_personalSoli.Open "select * from rc_personal WHERE status='S' ORDER BY PATERNO ", db, adOpenKeyset, adLockReadOnly
'    Set adopuestosol.Recordset = rstrc_personalSoli
'    adopuestosol.Refresh

            Set RsEntrega = New ADODB.Recordset
            If RsEntrega.State = 1 Then RsEntrega.Close
            'RsEntrega.CursorLocation = adUseClient
            'GlSqlAux = "SELECT * FROM ALEntrega ORDER BY IdEntrega"
            'RsEntrega.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
            RsEntrega.Open "SELECT * FROM ALEntrega ORDER BY IdEntrega", db, adOpenKeyset, adLockOptimistic
            If RsEntrega.RecordCount > 0 Then
               GlHayRegs = True  'Variable global
            Else
               GlHayRegs = False
            End If
            BotonesNavegar Me
            Habilita False
            Set TdbgEntrega.DataSource = RsEntrega
        Case 1
                    
        Case 2
        
    End Select
    '
    Screen.MousePointer = vbDefault
    Me.Show
End Sub
Private Sub CmdAnadir_Click()
    estado = 1
    Set TdbgEntrega.DataSource = rsNada
    RsEntrega.AddNew
    dtpFechaEntrega.Value = Null
    BotonesEditar Me
    Habilita True
    lblEstadoRs.Caption = "Agregando Registro..."
End Sub
Private Sub CmdBuscar_Click()
'JQA
'  Set ClBuscaGrid = New  ClBuscaEnGridPropio
'  Set ClBuscaGrid.Conexión = db
'  ClBuscaGrid.FiltrosMultiples = True
'  ClBuscaGrid.QueryUtilizado = "SELECT * FROM AlEntrega"
'  ClBuscaGrid.Título = "Elija una Entrega"
'  ClBuscaGrid.OcultarPrimero = True
'  ClBuscaGrid.Ejecutar
'  If ClBuscaGrid.ElegidoCol1 <> "" Then
'    RsEntrega.Filter = adFilterNone
'    RsEntrega.MoveFirst
'    RsEntrega.Find "IdEntrega = " & ClBuscaGrid.ElegidoCol1
'  End If
'  Set ClBuscaGrid = Nothing
'JQA
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo Que_Error
    Screen.MousePointer = vbHourglass
    If RsEntrega.EditMode <> adEditNone Then RsEntrega.CancelUpdate
    BotonesNavegar Me
    Habilita False
    estado = 0
    RsEntrega.Requery
    Set TdbgEntrega.DataSource = RsEntrega
    Totales
    Screen.MousePointer = vbDefault
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox err.Number & " : " & err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub
Private Sub CmdDetalle_Click(Index As Integer)
On Error GoTo QError
    Select Case Index
        Case 0 ' Agregar
            With ALFrmEntregaDet
                .estado = 1
                .Show vbModal
                If .QResp Then
                    GlSqlAux = "INSERT INTO AlAuxEntregaDet(Usuario, CodArt, CantidadEntCaj, CantidadEntEj, DescArt, PesoKgs, PrecioSus, PesoTotalKgs, PrecioTotalSus) " & _
                               "SELECT '" & NombreTerminal & "','" & .CodItem & "'," & .CantCaja & "," & .CantEjem & ",'" & .Item & "'," & .PesoKgs & "," & .PrecioSus & "," & .PesoTotal & ", " & .PrecioTotal & " "
                    db.Execute GlSqlAux
                    '--
                    GlSqlAux = "SELECT * FROM ALAuxEntregaDet WHERE Usuario = '" & NombreTerminal & "'"
                    Set RsDet = New ADODB.Recordset
                    RsDet.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
                    Set AdoDetalle.Recordset = RsDet
                    Set TDBGDet.DataSource = RsDet
                    Totales
                End If
            End With
        Case 1 ' Modificar
            If AdoDetalle.Recordset.RecordCount <= 0 Then Beep: Exit Sub
            With ALFrmEntregaDet
                .estado = 1
                .CodItem = AdoDetalle.Recordset!CodArt
                .Item = AdoDetalle.Recordset!DescArt
                .CantCaja = AdoDetalle.Recordset!CantidadEntCaj
                .CantEjem = AdoDetalle.Recordset!CantidadEntEj
                .PesoKgs = AdoDetalle.Recordset!PesoKgs
                .PrecioSus = AdoDetalle.Recordset!PrecioSus
                .PesoTotal = AdoDetalle.Recordset!PesoTotalKgs
                .PrecioTotal = AdoDetalle.Recordset!PrecioTotalSus
                .estado = 2
                .Show vbModal
                If .QResp Then
                    GlSqlAux = "UPDATE AlAuxEntregaDet SET " & _
                               "CodArt = '" & .CodItem & "', " & _
                               "CantidadEntCaj = " & .CantCaja & ", " & _
                               "CantidadEntEj = '" & .CantEjem & "', " & _
                               "DescArt = '" & .Item & "', " & _
                               "PesoKgs = " & .PesoKgs & ", " & _
                               "PrecioSus = " & .PrecioSus & ", " & _
                               "PesoTotalKgs = " & .PesoTotal & ", " & _
                               "PrecioTotalSus = " & .PrecioTotal & " " & _
                               "WHERE Usuario = '" & NombreTerminal & "' AND CodArt = '" & AdoDetalle.Recordset!CodArt & "'"
                    db.Execute GlSqlAux
                    '--
                    GlSqlAux = "SELECT * FROM ALAuxEntregaDet WHERE Usuario = '" & NombreTerminal & "'"
                    Set RsDet = New ADODB.Recordset
                    RsDet.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
                    Set AdoDetalle.Recordset = RsDet
                    Set TDBGDet.DataSource = RsDet
                    Totales
                End If
            End With
        Case 2 ' Eliminar
            If AdoDetalle.Recordset.RecordCount <= 0 Then Beep: Exit Sub
            If MsgBox("Eliminará el Detalle seleccionado." & vbCrLf & "Esta seguro?", vbQuestion + vbYesNo, "Atención") = vbYes Then
                AdoDetalle.Recordset.Delete
                AdoDetalle.Recordset.Requery
                Totales
            End If
    End Select
    Exit Sub
QError:
    MsgBox err.Description & " : " & err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub cmdEditar_Click()
On Error GoTo Que_Error    '
    Screen.MousePointer = vbHourglass
    BotonesEditar Me
    estado = 2
    Habilita True
    lblEstadoRs.Caption = "Editando Registro..."
    Screen.MousePointer = vbDefault
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox err.Number & " : " & err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub cmdEliminar_Click()
On Error GoTo Que_Error
    If Not GlHayRegs Then
        MsgBox "No existen registro para eliminar", vbExclamation + vbOKOnly, "Atención"
        Exit Sub
    End If
    If MsgBox("¿ Está seguro que se va a borrar el registro seleccionado ?", vbExclamation + vbOKCancel, "Atención") = vbOK Then
        Screen.MousePointer = vbHourglass
        db.BeginTrans
        '-- Eliminamos el detalle del Ingreso
        GlSqlAux = "DELETE FROM AlEntregaDet WHERE IdEntrega = " & RsEntrega!IdEntrega
        db.Execute GlSqlAux
        '--
        NoEntrega = RsEntrega!IdEntrega
        db.ActualizaAlmacen CStr(NombreTerminal), NoEntrega, Format(Date, FormatoFecha), 0, 1
        '--
        RsEntrega.Delete
        db.CommitTrans
        RsEntrega.MoveNext
        If RsEntrega.EOF Then
          If RsEntrega.RecordCount > 0 Then
            RsEntrega.MoveLast
          Else
            GlHayRegs = False
            RsEntrega.Requery
          End If
        End If
        Screen.MousePointer = vbDefault
    End If
    BotonesNavegar Me
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox err.Number & " : " & err.Description, vbExclamation + vbOKOnly, "Atención"
    db.RollbackTrans
End Sub

Private Sub CmdGrabar_Click()
On Error GoTo QError
    '--
    '--
    If valida Then
        Screen.MousePointer = vbHourglass
        ' Empezar a grabar
        '*********************************
        db.BeginTrans
        ' Campos no ligados
        If estado = 1 Then
            rsPrm.Requery
            NoEntrega = rsPrm!NoEntrega + 1
            RsEntrega!IdEntrega = NoEntrega
            rsPrm!NoEntrega = NoEntrega
            rsPrm.Update
        Else
            NoEntrega = RsEntrega!IdEntrega
        End If
        '*********************************
        ' Grabar
        RsEntrega.Update
        ' Grabamos el Detalle
        '--
        ' eliminamos el detalle para almacenar el nuevo detalle
        GlSqlAux = "DELETE FROM ALEntregaDet WHERE IdEntrega = " & NoEntrega
        db.Execute GlSqlAux
        '--
        GlSqlAux = "INSERT INTO ALEntregaDet (IdEntrega, CodArt, CantidadEntCaj, CantidadEntEj, DescArt, PesoKgs, PrecioSus, PesoTotalKgs, PrecioTotalSus) " & _
                   "SELECT " & NoEntrega & ", CodArt, CantidadEntCaj, CantidadEntEj, DescArt, PesoKgs, PrecioSus, PesoTotalKgs, PrecioTotalSus " & _
                   "FROM ALAuxEntregaDet WHERE Usuario = '" & NombreTerminal & "'"
        db.Execute GlSqlAux
        '--
        '--
        db.CommitTrans
    '*********************************
        lblEstadoRs.Caption = "Registro: " & CStr(RsEntrega.AbsolutePosition) & " de " & RsEntrega.RecordCount
        ' Colocar los botones en modo navegar
        GlHayRegs = True
        BotonesNavegar Me
        Habilita False
        Screen.MousePointer = vbDefault
        estado = 0
        RsEntrega.Requery
        Set TdbgEntrega.DataSource = RsEntrega
        Totales
    End If
    Exit Sub
QError:
    Screen.MousePointer = vbDefault
    ' Manejo de errores
    MsgBox err.Number & " : " & err.Description, vbExclamation + vbOKOnly, "Atención"
    db.RollbackTrans
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub cmdVerificar_Click()
On Error GoTo QError
Dim rs As ADODB.Recordset
Dim Resp As Integer
    ' Validar
    If RsEntrega.RecordCount <= 0 Then Beep: Exit Sub
    If Not ValidarVerificar Then
        MsgBox "No se puede continuar con esta Operación debido a las siguientes causas:" & vbCrLf & CadenaError, vbInformation + vbOKOnly, "Atención"
        Exit Sub
    End If
    If Not ValidarSaldo Then
        MsgBox "No se puede continuar con esta Operación debido a las siguientes causas:" & vbCrLf & CadenaError, vbInformation + vbOKOnly, "Atención"
        Exit Sub
    End If
    '--
    '--
    If MsgBox("Esta Operación Verificará la Entrega Nro. '" & txtNoEntrega.Text & "' de Items de Almacen." & vbCrLf & _
              "Esta seguro ?", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
    '--
    '--
    Screen.MousePointer = vbHourglass
    '--
    NoEntrega = RsEntrega!IdEntrega
    '--
    Set rs = New ADODB.Recordset
    Set rs = AdoDetalle.Recordset.Clone
    '--
    'Resp = MsgBox("La Entrega de Items sera automática o manual?." & vbCrLf & "Automático [SI] - Manual [NO] - Cancelar [CANCELAR]", vbQuestion + vbYesNoCancel, "Atención")
    With GrFrmOpciones
        Screen.MousePointer = vbDefault
        .OptOpciones(1).Caption = "Entrega de Items"
        .OptOpciones(2).Caption = "Cancelar Operación"
        .Show vbModal
        Screen.MousePointer = vbHourglass
        If .POpcionElegida = 10 Then
            MsgBox "Ooops!!! No esta Habilitado.", vbInformation + vbOKOnly, "Atención"
            Exit Sub
            ' Automatica
            While Not rs.EOF
                db.ALActualizaIngreso NoEntrega, CStr(rs!CodArt), CLng(rs!CantidadEntCaj), CLng(rs!CantidadEntEj)
                rs.MoveNext
            Wend
        ElseIf .POpcionElegida = 1 Then
            ' Manual
            While Not rs.EOF
                With ALFrmEntregaManual
                    Screen.MousePointer = vbDefault
                    .ALPrincipal CStr(rs!CodArt), CLng(rs!CantidadEntCaj), CLng(rs!CantidadEntEj)
                    Screen.MousePointer = vbHourglass
                    If .QResp Then
                        GlSqlAux = "INSERT INTO ALEntDeIng(IdEntrega, CodArt, CodProveedor, IdIngreso, CantidadEntCaj, CantidadEntEj) " & _
                                   "SELECT " & NoEntrega & ", CodArt, CodProveedor, IdIngreso, CantidadEntCaj, CantidadEntEj " & _
                                   "FROM ALAuxEntDeIng " & _
                                   "WHERE Usuario = '" & NombreTerminal & "' AND CodArt = '" & CStr(rs!CodArt) & "'"
                        db.Execute GlSqlAux
                        GlSqlAux = "UPDATE AlEntregaDet " & _
                                   "SET CantidadEntEj = (SELECT SUM(CantidadEntEj) FROM ALEntDeIng WHERE IdEntrega = " & NoEntrega & " AND CodArt = '" & CStr(rs!CodArt) & "') " & _
                                   "WHERE IdEntrega = " & NoEntrega & " AND CodArt = '" & CStr(rs("CodArt")) & "'"
                        db.Execute GlSqlAux
                    Else
                        GlSqlAux = "DELETE FROM ALEntDeIng " & _
                                   "WHERE IdEntrega = " & NoEntrega & ""
                        db.Execute GlSqlAux
                        Exit Sub
                    End If
                End With
                db.ALActualizaIngreso NoEntrega, CStr(rs!CodArt), CLng(rs!CantidadEntCaj), CLng(rs!CantidadEntEj)
                rs.MoveNext
            Wend
        ElseIf .POpcionElegida = 2 Then
            Screen.MousePointer = vbDefault
            MsgBox "Operación Cancelada.", vbInformation + vbOKOnly, "Atención"
            Exit Sub
        Else
            Screen.MousePointer = vbDefault
            MsgBox "Opss!!!. Operación Cancelada.", vbInformation + vbOKOnly, "Atención"
            Exit Sub
        End If
    End With
    '--
    '--
    RsEntrega!Verificado = 1
    RsEntrega!FechaVerificado = Date
    RsEntrega.Update
    '--
    RsEntrega.Requery
    '--
    db.ALActualizaAlmacen NoEntrega, Format(Date, FormatoFecha), 0
    '--
    RsEntrega.Find "IdEntrega = " & NoEntrega
    '--
    Screen.MousePointer = vbDefault
    '--
    MsgBox "Verificación de Entrega Completada.", vbInformation + vbOKOnly, "Atención"
    Exit Sub
QError:
    MsgBox err.Number & " : " & err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Screen.MousePointer = vbHourglass
    txtTipoCambio = GlTipoCambioOficial
    '
    Set RsDestino = New ADODB.Recordset
    GlSqlAux = "SELECT * FROM ALCLDestinos ORDER BY CodDestino"
    RsDestino.Open GlSqlAux, db, adOpenStatic
    Set tdbcDestino.RowSource = RsDestino
    '--
    Set RsMat = New ADODB.Recordset
    GlSqlAux = "SELECT CodGrupo + '-' + CodDetalle AS Codigo, DescDetalle, Unidad " & _
               "FROM ALCLDetalle " & _
               "WHERE (Estado = 1)"
    RsMat.Open GlSqlAux, db, adOpenStatic
    Set tdbdMat.DataSource = RsMat
    '--
'    CodArt = ""
'    Set RsIngreso = New ADODB.Recordset
'    GlSqlAux = "SELECT ALIngresoAlm.NoLicitacion, ALIngresoAlm.FechaIng, " & _
'               "ALIngresoAlm.IdIngreso, AlIngresoAlmDet.CodArt, " & _
'               "AlIngresoAlmDet.CantidadCaj - AlIngresoAlmDet.CantidadEntCaj AS SaldoCajas, " & _
'               "AlIngresoAlmDet.CantidadEj - AlIngresoAlmDet.CantidadEntEj AS SaldoEjemp " & _
'               "FROM ALIngresoAlm INNER JOIN AlIngresoAlmDet ON ALIngresoAlm.IdIngreso = AlIngresoAlmDet.IdIngreso " & _
'               "WHERE (AlIngresoAlmDet.CodArt = '" & CodArt & "') " & _
'               "ORDER BY ALIngresoAlm.FechaIng"
'    RsIngreso.Open GlSqlAux, db, adOpenStatic
'    Set tdbdIngreso.DataSource = RsIngreso
'    TDBGDet.Columns("NroIngreso").DropDown = IIf(RsIngreso.RecordCount > 0, tdbdIngreso, "")
'    TDBGDet.Columns("NroIngreso").BackColor = IIf(RsIngreso.RecordCount > 0, RGB(255, 255, 255), RGB(219, 219, 219))
    '--
    Set txtDescripcion.DataSource = RsEntrega
    Set txtNoEntrega.DataSource = RsEntrega
    Set dtpFechaEntrega.DataSource = RsEntrega
    Set tdbcDestino.DataSource = RsEntrega
    Set txtTipoCambio.DataSource = RsEntrega
    Screen.MousePointer = vbDefault
	Call SeguridadSet(Me)
End Sub
Private Function valida() As Boolean
Dim rs As ADODB.Recordset
Dim rsAux As ADODB.Recordset
Dim CadenaError As String
    valida = False
    If Trim(txtDescripcion.Text) = "" Then
        MsgBox "Ingrese la Descripción del Ingreso.", vbExclamation + vbOKOnly, "Atención"
        txtDescripcion.SetFocus
        Exit Function
    End If
    GlSqlAux = "SELECT Count(*) As Cuantos FROM ALEntrega WHERE NoEntrega = '" & txtNoEntrega.Text & "'"
    Set rs = New ADODB.Recordset
    rs.Open GlSqlAux, db, adOpenStatic
    If rs!Cuantos > 0 And estado = 1 Then
        MsgBox "El No. de Entrega '" & txtNoEntrega.Text & "', ya se encuentra registrado..", vbExclamation + vbOKOnly, "Atención"
        txtNoEntrega.SetFocus
        Exit Function
    End If
    rs.Close
    Set rs = Nothing
    If IsNull(dtpFechaEntrega.Value) Then
        MsgBox "Ingrese la fecha de Ingreso de Materiales a Almacen.", vbExclamation + vbOKOnly, "Atención"
        dtpFechaEntrega.SetFocus
        Exit Function
    End If
    tdbcDestino.Text = TexTFrenSer
    If Trim(tdbcDestino.Text) = "" Then
        MsgBox "Ingrese el Destino de la Entrega.", vbExclamation + vbOKOnly, "Atención"
        tdbcDestino.SetFocus
        Exit Function
    End If
    If Trim(txtTipoCambio.Text) = "" Then
        MsgBox "Ingrese el Tipo de Cambio de la Entrega.", vbExclamation + vbOKOnly, "Atención"
        txtTipoCambio.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtTipoCambio.Text) Then
        MsgBox "Ingrese un Tipo de Cambio de Entrega Válido.", vbExclamation + vbOKOnly, "Atención"
        txtTipoCambio.SetFocus
        Exit Function
    End If
    If AdoDetalle.Recordset.RecordCount <= 0 Then
        MsgBox "Ingrese el detalle del Ingreso a Almacen.", vbExclamation + vbOKOnly, "Atención"
        TDBGDet.SetFocus
        Exit Function
    End If
    valida = True
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  RsEntrega.Close
  RsDestino.Close
  rsNada.Close
  RsDet.Close
  RsMat.Close
  RsIngreso.Close
End Sub




Private Sub RsEntrega_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If RsEntrega.BOF Or RsEntrega.EOF Then
        If RsEntrega.BOF And RsEntrega.EOF Then
            txtEstado.Text = ""
            txtDescripcion.Text = ""
            txtNoEntrega.Text = ""
            dtpFechaEntrega.Value = Null
            Set TDBGDet.DataSource = rsNada
            lblEstadoRs.Caption = "Registro: 0 de 0"
        Else
            Exit Sub
        End If
    Else
        TexTFrenSer = tdbcDestino.Text
        lblEstadoRs.Caption = "Registro: " & RsEntrega.AbsolutePosition & " de " & RsEntrega.RecordCount
        ' Cargamos el Detalle del Ingreso
        GlSqlAux = "DELETE ALAuxEntDeIng WHERE Usuario = '" & NombreTerminal & "'"
        db.Execute GlSqlAux
        
        GlSqlAux = "DELETE ALAuxEntregaDet WHERE Usuario = '" & NombreTerminal & "'"
        db.Execute GlSqlAux
        '--
        If estado = 1 Then
            txtEstado.Text = "SIN VERIFICAR"
        Else
            GlSqlAux = "INSERT INTO ALAuxEntregaDet (Usuario, CodArt, CantidadEntCaj, CantidadEntEj, DescArt, PesoKgs, PrecioSus, PesoTotalKgs, PrecioTotalSus, CantidadEntCajA, CantidadEntEjA) " & _
                       "SELECT '" & NombreTerminal & "', CodArt, CantidadEntCaj, CantidadEntEj, DescArt, PesoKgs, PrecioSus, PesoTotalKgs, PrecioTotalSus, CantidadEntCaj, CantidadEntEj " & _
                       "FROM ALEntregaDet WHERE IdEntrega = " & RsEntrega!IdEntrega
            db.Execute GlSqlAux
            '--
            cmdEliminar.Enabled = Not CBool(RsEntrega!Verificado)
            cmdEditar.Enabled = Not CBool(RsEntrega!Verificado)
            cmdVerificar.Enabled = Not CBool(RsEntrega!Verificado)
            '--
            txtEstado.Text = IIf(CBool(RsEntrega!Verificado), "VERIFICADO", "SIN VERIFICAR")
        End If
        '--
        GlSqlAux = "SELECT * FROM ALAuxEntregaDet WHERE Usuario = '" & NombreTerminal & "'"
        Set RsDet = New ADODB.Recordset
        RsDet.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
        Set AdoDetalle.Recordset = RsDet
        Set TDBGDet.DataSource = RsDet
        Totales
    End If
End Sub
Private Sub tdbcDestino_NotInList(NewEntry As String, Retry As Integer)
    tdbcDestino.Text = ""
End Sub
Private Sub tdbdIngreso_DropDownClose()
    TDBGDet.Columns("IdIngreso").Value = tdbdIngreso.Columns("IdIngreso").Value
End Sub
Private Sub tdbdMat_DropDownClose()
    TDBGDet.Columns("DescArt").Value = tdbdMat.Columns("DescDetalle").Value
    '--
    CodArt = tdbdMat.Columns("Codigo").Value
    Set RsIngreso = New ADODB.Recordset
    GlSqlAux = "SELECT ALIngresoAlm.NoLicitacion, ALIngresoAlm.FechaIng, " & _
               "ALIngresoAlm.IdIngreso, AlIngresoAlmDet.CodArt, " & _
               "AlIngresoAlmDet.CantidadCaj - AlIngresoAlmDet.CantidadEntCaj AS SaldoCajas, " & _
               "AlIngresoAlmDet.CantidadEj - AlIngresoAlmDet.CantidadEntEj AS SaldoEjemp " & _
               "FROM ALIngresoAlm INNER JOIN AlIngresoAlmDet ON ALIngresoAlm.IdIngreso = AlIngresoAlmDet.IdIngreso " & _
               "WHERE (AlIngresoAlmDet.CodArt = '" & CodArt & "') " & _
               "ORDER BY ALIngresoAlm.FechaIng"
    RsIngreso.Open GlSqlAux, db, adOpenStatic
    Set tdbdIngreso.DataSource = RsIngreso
    TDBGDet.Columns("NroIngreso").DropDown = IIf(RsIngreso.RecordCount > 0, tdbdIngreso, "")
    TDBGDet.Columns("NroIngreso").BackColor = IIf(RsIngreso.RecordCount > 0, RGB(255, 255, 255), RGB(219, 219, 219))
    '--
End Sub
Private Sub tdbgDet_AfterUpdate()
    Totales
End Sub

Private Sub Habilita(sw As Boolean)
    FraDatos.Enabled = sw
    cmdDetalle(0).Enabled = sw
    cmdDetalle(1).Enabled = sw
    cmdDetalle(2).Enabled = sw
    
'    With TDBGDet
'        .AllowAddNew = Sw
'        .AllowDelete = Sw
'        .AllowUpdate = Sw
'    End With
End Sub
Public Sub Totales()
Dim rs As ADODB.Recordset
Dim PrecioSus As Currency
Dim PesoKgs As Currency
Dim PrecioTotalSus As Currency
Dim PesoTotalKgs As Currency
Dim CantCaja As Integer
Dim CantEje As Integer
    Set rs = New ADODB.Recordset
    Set rs = AdoDetalle.Recordset.Clone
    PrecioSus = 0
    PesoKgs = 0
    PrecioTotalSus = 0
    PesoTotalKgs = 0
    CantCaja = 0
    CantEje = 0
    While Not rs.EOF
        PrecioSus = PrecioSus + IIf(IsNull(rs!PrecioSus), 0, rs!PrecioSus)
        PesoKgs = PesoKgs + IIf(IsNull(rs!PesoKgs), 0, rs!PesoKgs)
        PrecioTotalSus = PrecioTotalSus + IIf(IsNull(rs!PrecioTotalSus), 0, rs!PrecioTotalSus)
        PesoTotalKgs = PesoTotalKgs + IIf(IsNull(rs!PesoTotalKgs), 0, rs!PesoTotalKgs)
        CantCaja = CantCaja + IIf(IsNull(rs!CantidadEntCaj), 0, rs!CantidadEntCaj)
        CantEje = CantEje + IIf(IsNull(rs!CantidadEntEj), 0, rs!CantidadEntEj)
        rs.MoveNext
    Wend
    TDBGDet.Columns("DescArt").FooterText = "TOTALES"
'    TDBGDet.Columns("PrecioSus").FooterText = Format(PrecioSus, "###,###,##0.00") & " $us."
    TDBGDet.Columns("PesoKgs").FooterText = Format(PesoKgs, "###,###,##0.00") & " Kgs."
'    TDBGDet.Columns("PrecioTotalSus").FooterText = Format(PrecioTotalSus, "###,###,##0.00") & " $us."
    TDBGDet.Columns("PesoTotalKgs").FooterText = Format(PesoTotalKgs, "###,###,##0.00") & " Kgs"
    TDBGDet.Columns("CantidadEntCaj").FooterText = Format(CantCaja, "###,###,##0")
    TDBGDet.Columns("CantidadEntEj").FooterText = Format(CantEje, "###,###,##0")
End Sub
Private Function ValidarVerificar() As Boolean
Dim rs As ADODB.Recordset
    CadenaError = ""
    ValidarVerificar = True
    ' Validamos si ya esta Verificado
    If CBool(RsEntrega!Verificado) Then
        CadenaError = CadenaError & vbTab & "La Entrega '" & RsEntrega!NoEntrega & "' ya fue Verificado en Fecha '" & Format(RsEntrega!FechaVerificado, "dd/mm/yyyy") & "'." & vbCrLf
        ValidarVerificar = False
        Exit Function
    End If
    ' Validamos que el Detalle este Completo
    Set rs = AdoDetalle.Recordset.Clone
    While Not rs.EOF
        If Trim(rs!CodArt) = "" Then
            CadenaError = CadenaError & vbTab & "Falta Codigo del Item." & vbCrLf
            ValidarVerificar = False
        End If
        If CLng(rs!CantidadEntCaj) <= 0 Then
            CadenaError = CadenaError & vbTab & "Falta Cantidad de Cajas del Item '" & rs!CodArt & "'." & vbCrLf
            ValidarVerificar = False
        End If
'        If CLng(rs!CantidadEntEj) <= 0 Then
'            CadenaError = CadenaError & vbTab & "Falta Cantidad de Cajas del Item '" & rs!CodArt & "'." & vbCrLf
'            ValidarVerificar = False
'        End If
'        If CCur(rs!PesoKgs) <= 0 Then
'            CadenaError = CadenaError & vbTab & "Falta el Peso(Kgs) del Item '" & rs!CodArt & "'." & vbCrLf
'            ValidarVerificar = False
'        End If
'        If CCur(rs!PrecioSus) <= 0 Then
'            CadenaError = CadenaError & vbTab & "Falta el Precio($us) del Item '" & rs!CodArt & "'." & vbCrLf
'            ValidarVerificar = False
'        End If
'        If CCur(rs!PesoTotalKgs) <= 0 Then
'            CadenaError = CadenaError & vbTab & "Falta el Peso Total(Kgs) del Item '" & rs!CodArt & "'." & vbCrLf
'            ValidarVerificar = False
'        End If
'        If CCur(rs!PrecioTotalSus) <= 0 Then
'            CadenaError = CadenaError & vbTab & "Falta el Precio Total($us) del Item '" & rs!CodArt & "'." & vbCrLf
'            ValidarVerificar = False
'        End If
        CadenaError = CadenaError & "------------------------------------" & vbCrLf
        rs.MoveNext
    Wend
End Function


Private Sub TdbgEntrega_FetchRowStyle(ByVal Split As Integer, BookMark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If CBool(TdbgEntrega.Columns("Verificado").CellValue(BookMark)) Then
        RowStyle.BackColor = &H8000000D
        RowStyle.ForeColor = &HC0FFFF
        RowStyle.Font.Bold = True
    Else
        RowStyle.BackColor = &HC0FFFF
        RowStyle.ForeColor = &H8000000D
        RowStyle.Font.Bold = False
    End If
End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']", KeyAscii, 0)
End Sub
Private Function ValidarSaldo() As Boolean
Dim rs As ADODB.Recordset
Dim rsAux As ADODB.Recordset
    CadenaError = ""
    ValidarSaldo = True
    ' Validamos que el Detalle este Completo
    Set rs = AdoDetalle.Recordset.Clone
    
    While Not rs.EOF
        If Trim(rs!CodArt) = "" Then
            CadenaError = CadenaError & vbTab & "Falta Codigo del Item." & vbCrLf
            ValidarSaldo = False
            Exit Function
        End If
        Set rsAux = New ADODB.Recordset
        GlSqlAux = "SELECT CantidadCaj, CantidadEj FROM ALMaterial WHERE CodArt = '" & rs!CodArt & "'"
        rsAux.Open GlSqlAux, db, adOpenStatic
        If rs!CantidadEntCaj > rsAux!CantidadCaj And rs!CantidadEntEj > rsAux!CantidadEj Then
            CadenaError = CadenaError & vbTab & "No existe suficiente cantidad del Item '" & rs!CodArt & "' para realizar la Entrega. " & vbCrLf & vbTab & vbTab & "Existe en Almacen : " & Format(rsAux!CantidadCaj, "###,###,##0") & " Cajas y " & Format(rsAux!CantidadEj, "###,###,##0") & " Ejemp." & _
                          vbCrLf & vbTab & vbTab & "Cantidad a Entregar : " & Format(rs!CantidadEntCaj, "###,###,##0") & " Cajas y " & Format(rs!CantidadEntEj, "###,###,##0") & " Ejemp." & vbCrLf
            ValidarSaldo = False
        End If
        rsAux.Close
        Set rsAux = Nothing
        CadenaError = CadenaError & "------------------------------------" & vbCrLf
        rs.MoveNext
    Wend
End Function

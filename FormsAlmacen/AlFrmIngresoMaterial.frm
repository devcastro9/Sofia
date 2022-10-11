VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form AlFrmIngresoMaterial 
   Caption         =   "Ingreso de Material"
   ClientHeight    =   7950
   ClientLeft      =   -90
   ClientTop       =   1425
   ClientWidth     =   11910
   Icon            =   "AlFrmIngresoMaterial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11910
      TabIndex        =   30
      Top             =   7455
      Width           =   11910
      Begin VB.Frame Frame4 
         Height          =   60
         Left            =   1215
         TabIndex        =   31
         Top             =   255
         Width           =   7290
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso a Almacen"
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
         Left            =   8715
         TabIndex        =   32
         Top             =   90
         Width           =   2925
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso a Almacen"
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
         Left            =   8730
         TabIndex        =   33
         Top             =   105
         Width           =   2925
      End
   End
   Begin TabDlg.SSTab sstDetalle 
      Height          =   4335
      Left            =   3120
      TabIndex        =   29
      Top             =   2865
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   7646
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Item"
      TabPicture(0)   =   "AlFrmIngresoMaterial.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "AdoDetalle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdbgDet"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tdbdProv"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tdbdMat"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDetalle(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdDetalle(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdDetalle(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Eliminar Detalle"
         Height          =   360
         Index           =   2
         Left            =   5085
         TabIndex        =   15
         Top             =   3450
         Width           =   2355
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Modificar Detalle"
         Height          =   360
         Index           =   1
         Left            =   2595
         TabIndex        =   14
         Top             =   3450
         Width           =   2355
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Agregar Detalle"
         Height          =   360
         Index           =   0
         Left            =   105
         TabIndex        =   13
         Top             =   3450
         Width           =   2355
      End
      Begin TrueOleDBGrid60.TDBDropDown tdbdMat 
         Height          =   705
         Left            =   1575
         OleObjectBlob   =   "AlFrmIngresoMaterial.frx":0EE6
         TabIndex        =   34
         Top             =   105
         Width           =   1125
      End
      Begin TrueOleDBGrid60.TDBDropDown tdbdProv 
         Height          =   945
         Left            =   1560
         OleObjectBlob   =   "AlFrmIngresoMaterial.frx":34E2
         TabIndex        =   35
         Top             =   105
         Width           =   1215
      End
      Begin TrueOleDBGrid60.TDBGrid tdbgDet 
         Bindings        =   "AlFrmIngresoMaterial.frx":54EF
         Height          =   3255
         Left            =   75
         OleObjectBlob   =   "AlFrmIngresoMaterial.frx":5508
         TabIndex        =   16
         Top             =   90
         Width           =   7350
      End
      Begin MSAdodcLib.Adodc AdoDetalle 
         Height          =   330
         Left            =   645
         Top             =   3915
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
   End
   Begin TrueOleDBGrid60.TDBGrid TdbgIngreso 
      Height          =   5850
      Left            =   0
      OleObjectBlob   =   "AlFrmIngresoMaterial.frx":C490
      TabIndex        =   7
      Top             =   960
      Width           =   3090
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   990
      Left            =   0
      Picture         =   "AlFrmIngresoMaterial.frx":FEE0
      ScaleHeight     =   930
      ScaleWidth      =   11850
      TabIndex        =   21
      Top             =   0
      Width           =   11910
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "Verificar Ingreso"
         Height          =   855
         Left            =   10275
         Picture         =   "AlFrmIngresoMaterial.frx":1317A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         Width           =   1425
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
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
         Left            =   90
         TabIndex        =   27
         Top             =   675
         Width           =   735
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   195
         Left            =   1035
         TabIndex        =   26
         Top             =   705
         Width           =   2310
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
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
         Left            =   90
         TabIndex        =   25
         Top             =   435
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   1035
         TabIndex        =   24
         Top             =   465
         Width           =   540
      End
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE INGRESO DE MATERIAL"
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
         Left            =   75
         TabIndex        =   23
         Top             =   15
         Width           =   6045
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         Caption         =   "."
         ForeColor       =   &H0000C000&
         Height          =   180
         Left            =   4815
         TabIndex        =   22
         Top             =   675
         Width           =   2655
      End
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE INGRESO DE MATERIAL"
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
         Left            =   105
         TabIndex        =   36
         Top             =   30
         Width           =   6045
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   0
         Picture         =   "AlFrmIngresoMaterial.frx":135BC
         Top             =   0
         Width           =   11640
      End
   End
   Begin VB.Frame FraDatos 
      BorderStyle     =   0  'None
      Height          =   1890
      Left            =   3120
      TabIndex        =   17
      Top             =   960
      Width           =   7530
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
         Left            =   5445
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   105
         Width           =   2010
      End
      Begin VB.TextBox txtNoIngreso 
         DataField       =   "NroIngreso"
         Height          =   300
         Left            =   150
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1440
         Width           =   1200
      End
      Begin VB.TextBox TxtDescripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Obs"
         Height          =   645
         Left            =   150
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   465
         Width           =   7305
      End
      Begin VB.TextBox txtNoLicitacion 
         DataField       =   "NoLicitacion"
         Height          =   300
         Left            =   1455
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1440
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker dtpFechIngreso 
         DataField       =   "FechaIng"
         Height          =   300
         Left            =   2775
         TabIndex        =   12
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   51314689
         CurrentDate     =   36737
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Ingreso"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   1215
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ingreso"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   2775
         TabIndex        =   20
         Top             =   1215
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   225
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Licitación"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   1455
         TabIndex        =   18
         Top             =   1215
         Width           =   1020
      End
   End
   Begin VB.Frame FraOpciones 
      Height          =   6315
      Left            =   10680
      TabIndex        =   38
      Top             =   960
      Width           =   1140
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   855
         Left            =   120
         Picture         =   "AlFrmIngresoMaterial.frx":3A62C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5310
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   855
         Left            =   120
         Picture         =   "AlFrmIngresoMaterial.frx":3A836
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4470
         Width           =   855
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Borrar"
         Height          =   855
         Left            =   120
         Picture         =   "AlFrmIngresoMaterial.frx":3AA40
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3630
         Width           =   855
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   855
         Left            =   120
         Picture         =   "AlFrmIngresoMaterial.frx":3B12A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2790
         Width           =   855
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   855
         Left            =   120
         Picture         =   "AlFrmIngresoMaterial.frx":3B334
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1950
         Width           =   855
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "Modificar"
         Height          =   855
         Left            =   120
         Picture         =   "AlFrmIngresoMaterial.frx":3B53E
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
         Picture         =   "AlFrmIngresoMaterial.frx":3B748
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
      TabIndex        =   39
      Top             =   6840
      Width           =   3090
   End
End
Attribute VB_Name = "AlFrmIngresoMaterial"
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
Dim WithEvents RsIngreso As ADODB.Recordset
Attribute RsIngreso.VB_VarHelpID = -1
Dim rsProv As ADODB.Recordset
Dim rsNada As ADODB.Recordset
Dim RsDet As ADODB.Recordset
Dim RsMat As ADODB.Recordset
'--------
Dim estado As Integer ' 0 navegar, 1 Agregar, 2 Editar
Dim NoIngreso As Integer
Dim CadenaError As String
'--
'JQA
'Dim ClBuscaGrid As  ClBuscaEnGridPropio
Public Sub ALPrincipal(QEstado As Integer, Optional IdIngreso As Integer = 0)
    '
    Screen.MousePointer = vbHourglass
    estado = QEstado
    '
    Select Case estado
        Case 0 ' Navegar
            Set RsIngreso = New ADODB.Recordset
            RsIngreso.CursorLocation = adUseClient
            GlSqlAux = "SELECT * FROM ALIngresoAlm ORDER BY IdIngreso"
            RsIngreso.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
            If RsIngreso.RecordCount > 0 Then
               GlHayRegs = True  'Variable global
            Else
               GlHayRegs = False
            End If
            BotonesNavegar Me
            Habilita False
            Set tdbgIngreso.DataSource = RsIngreso
        Case 1 ' Agregar
            Set RsIngreso = New ADODB.Recordset
            RsIngreso.CursorLocation = adUseClient
            GlSqlAux = "SELECT * FROM ALIngresoAlm ORDER BY IdIngreso"
            RsIngreso.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
            If RsIngreso.RecordCount > 0 Then
               GlHayRegs = True  'Variable global
            Else
               GlHayRegs = False
            End If
            Set tdbgIngreso.DataSource = RsIngreso
            CmdAnadir_Click
        Case 2 ' Editar
            Set RsIngreso = New ADODB.Recordset
            RsIngreso.CursorLocation = adUseClient
            GlSqlAux = "SELECT * FROM ALIngresoAlm ORDER BY IdIngreso"
            RsIngreso.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
            If RsIngreso.RecordCount > 0 Then
               GlHayRegs = True  'Variable global
            Else
               GlHayRegs = False
            End If
            Set tdbgIngreso.DataSource = RsIngreso
            RsIngreso.Find "IdIngreso = " & IdIngreso
            cmdEditar_Click
    End Select
    '--
    Screen.MousePointer = vbDefault
    Me.Show
End Sub

Private Sub CmdAnadir_Click()
    estado = 1
    Set tdbgIngreso.DataSource = rsNada
    RsIngreso.AddNew
    RsIngreso!TipoIng = 1
    BotonesEditar Me
    Habilita True
    lblEstadoRs.Caption = "Agregando Registro..."
End Sub

Private Sub CmdBuscar_Click()
'JQA
'  Set ClBuscaGrid = New  ClBuscaEnGridPropio
'  Set ClBuscaGrid.Conexión = db
'  ClBuscaGrid.FiltrosMultiples = True
'  ClBuscaGrid.QueryUtilizado = "SELECT * FROM AlIngresoAlm"
'  ClBuscaGrid.Título = "Elija un Ingreso"
'  ClBuscaGrid.OcultarPrimero = True
'  ClBuscaGrid.Ejecutar
'  If ClBuscaGrid.ElegidoCol1 <> "" Then
'    RsIngreso.Filter = adFilterNone
'    RsIngreso.MoveFirst
'    RsIngreso.Find "IdIngreso = " & ClBuscaGrid.ElegidoCol1
'  End If
'  Set ClBuscaGrid = Nothing
'JQA
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo Que_Error
    Screen.MousePointer = vbHourglass
    If RsIngreso.EditMode <> adEditNone Then RsIngreso.CancelUpdate
    BotonesNavegar Me
    Habilita False
    estado = 0
    RsIngreso.Requery
    Set tdbgIngreso.DataSource = RsIngreso
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
            With ALFrmIngresoDet
                .estado = 1
                .Show vbModal
                If .QResp Then
                    GlSqlAux = "INSERT INTO AlAuxIngresoAlmDet(Usuario, CodProveedor, CodArt, NombProv, CantidadCaj, CantidadEj, UnidadCaja, DescArt, PesoKgs, PrecioSus, PesoTotalKgs, PrecioTotalSus, TipoCambio) " & _
                               "SELECT '" & NombreTerminal & "','" & .CodProv & "','" & .CodItem & "','" & .NomProv & "'," & .CantCaja & "," & .CantEjem & "," & .CantEjem / .CantCaja & ",'" & .Item & "'," & .PesoKgs & "," & .PrecioSus & "," & .PesoTotal & ", " & .PrecioTotal & "," & .TipoCambio & " "
                    db.Execute GlSqlAux
                    '--
                    GlSqlAux = "SELECT * FROM ALAuxIngresoAlmDet WHERE Usuario = '" & NombreTerminal & "'"
                    Set RsDet = New ADODB.Recordset
                    RsDet.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
                    Set AdoDetalle.Recordset = RsDet
                    Set TDBGDet.DataSource = RsDet
                    Totales
                End If
            End With
        Case 1 ' Modificar
            If AdoDetalle.Recordset.RecordCount <= 0 Then Beep: Exit Sub
            With ALFrmIngresoDet
                .estado = 1
                .CodProv = AdoDetalle.Recordset!CodProveedor
                .CodItem = AdoDetalle.Recordset!CodArt
                .NomProv = AdoDetalle.Recordset!NombProv
                .Item = AdoDetalle.Recordset!DescArt
                .CantCaja = AdoDetalle.Recordset!CantidadCaj
                .CantEjem = AdoDetalle.Recordset!CantidadEj
                .PesoKgs = AdoDetalle.Recordset!PesoKgs
                .PrecioSus = AdoDetalle.Recordset!PrecioSus
                .PesoTotal = AdoDetalle.Recordset!PesoTotalKgs
                .PrecioTotal = AdoDetalle.Recordset!PrecioTotalSus
                .TipoCambio = AdoDetalle.Recordset!TipoCambio
                .estado = 2
                .Show vbModal
                If .QResp Then
                    GlSqlAux = "UPDATE AlAuxIngresoAlmDet SET " & _
                               "CodProveedor = '" & .CodProv & "', " & _
                               "NombProv = '" & .NomProv & "', " & _
                               "CodArt = '" & .CodItem & "', " & _
                               "CantidadCaj = " & .CantCaja & ", " & _
                               "CantidadEj = '" & .CantEjem & "', " & _
                               "UnidadCaja = " & .CantEjem / .CantCaja & ", " & _
                               "DescArt = '" & .Item & "', " & _
                               "PesoKgs = " & .PesoKgs & ", " & _
                               "PrecioSus = " & .PrecioSus & ", " & _
                               "PesoTotalKgs = " & .PesoTotal & ", " & _
                               "PrecioTotalSus = " & .PrecioTotal & ", " & _
                               "TipoCambio = " & .TipoCambio & " " & _
                               "WHERE Usuario = '" & NombreTerminal & "' AND CodProveedor = '" & AdoDetalle.Recordset!CodProveedor & "' AND CodArt = '" & AdoDetalle.Recordset!CodArt & "'"
                    db.Execute GlSqlAux
                    '--
                    GlSqlAux = "SELECT * FROM ALAuxIngresoAlmDet WHERE Usuario = '" & NombreTerminal & "'"
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
On Error GoTo Que_Error
    '
    Screen.MousePointer = vbHourglass
    BotonesEditar Me
    estado = 2
    Habilita True
    If RsIngreso!Verificado = 1 Then ' Se Verifico una entrega
        cmdDetalle(0).Enabled = True
        cmdDetalle(1).Enabled = False
        cmdDetalle(2).Enabled = False
    Else
        cmdDetalle(0).Enabled = True
        cmdDetalle(1).Enabled = True
        cmdDetalle(2).Enabled = True
    End If
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
Dim Msj As String
    If Not GlHayRegs Then
        MsgBox "No existen registro para eliminar", vbExclamation + vbOKOnly, "Atención"
        Exit Sub
    End If
    If MsgBox("¿ Está seguro que se va a borrar el registro seleccionado ?", vbExclamation + vbOKCancel, "Atención") = vbOK Then
        Screen.MousePointer = vbHourglass
        db.BeginTrans
        '-- Eliminamos el detalle del Ingreso
        GlSqlAux = "DELETE FROM AlIngresoAlmDet WHERE IdIngreso = " & RsIngreso!IdIngreso
        db.Execute GlSqlAux
        '--
        NoIngreso = RsIngreso!IdIngreso
        '--
        '--
        RsIngreso.Delete
        db.CommitTrans
        RsIngreso.MoveNext
        If RsIngreso.EOF Then
          If RsIngreso.RecordCount > 0 Then
            RsIngreso.MoveLast
          Else
            GlHayRegs = False
            RsIngreso.Requery
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
            RsIngreso!EstadoIng = 1 ' Ingreso Sin Verificar
            rsPrm.Requery
            NoIngreso = rsPrm!NoIngreso + 1
            RsIngreso!IdIngreso = NoIngreso
            rsPrm!NoIngreso = NoIngreso
            rsPrm.Update
        Else
            NoIngreso = RsIngreso!IdIngreso
        End If
        '*********************************
        ' Grabar
        RsIngreso.Update
        ' Grabamos el Detalle
        '--
        ' eliminamos el detalle para almacenar el nuevo detalle
        GlSqlAux = "DELETE FROM ALIngresoAlmDet WHERE IdIngreso = " & NoIngreso
        db.Execute GlSqlAux
        '--
        GlSqlAux = "INSERT INTO ALIngresoAlmDet (IdIngreso, CodProveedor, CantidadCaj, CantidadEj, UnidadCaja, CodArt, DescArt, PesoKgs, PrecioSus, PesoTotalKgs, PrecioTotalSus, TipoCambio) " & _
                   "SELECT " & NoIngreso & ", CodProveedor, CantidadCaj, CantidadEj, UnidadCaja, CodArt, DescArt, PesoKgs, PrecioSus, PesoTotalKgs, PrecioTotalSus, TipoCambio " & _
                   "FROM ALAuxIngresoAlmDet WHERE Usuario = '" & NombreTerminal & "'"
        db.Execute GlSqlAux
        '--
        '--
        db.CommitTrans
    '*********************************
        lblEstadoRs.Caption = "Registro: " & CStr(RsIngreso.AbsolutePosition) & " de " & RsIngreso.RecordCount
        ' Colocar los botones en modo navegar
        GlHayRegs = True
        BotonesNavegar Me
        Habilita False
        Screen.MousePointer = vbDefault
        estado = 0
        'AdoIngreso.Refresh
        RsIngreso.Requery
        Set tdbgIngreso.DataSource = RsIngreso 'AdoIngreso
        Totales
    End If
    Exit Sub
QError:
    ' Manejo de errores
    MsgBox err.Number & " : " & err.Description, vbExclamation + vbOKOnly, "Atención"
    db.RollbackTrans
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVerificar_Click()
On Error GoTo QError
    ' Validar
    If RsIngreso.RecordCount <= 0 Then Beep: Exit Sub
    If Not ValidarVerificar Then
        MsgBox "No se puede continuar con esta Operación debido a las siguientes causas:" & vbCrLf & CadenaError, vbInformation + vbOKOnly, "Atención"
        Exit Sub
    End If
    '--
    If MsgBox("Esta Operación Verificará el Ingreso de Licitación Nro. '" & txtNoLicitacion.Text & "' a Almacen." & vbCrLf & _
              "Esta seguro ?", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
    '--
    NoIngreso = RsIngreso!IdIngreso
    RsIngreso!Verificado = 1
    RsIngreso!FechaVerificado = Date
    RsIngreso.Update
    '--
    db.ALActualizaAlmacen NoIngreso, Format(Date, FormatoFecha), 1
    '--
    RsIngreso.Find "IdIngreso = " & NoIngreso
    '--
    MsgBox "Verificación de Ingreso Completada.", vbInformation + vbOKOnly, "Atención"
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
    '
    Set rsProv = New ADODB.Recordset
    GlSqlAux = "SELECT * FROM ac_Proveedor ORDER BY Ruc_Id"
    rsProv.Open GlSqlAux, db, adOpenStatic
    tdbdProv.DataSource = rsProv
    '--
    Set RsMat = New ADODB.Recordset
    GlSqlAux = "SELECT CodGrupo + '-' + CodDetalle AS Codigo, DescDetalle, Unidad " & _
               "FROM ALCLDetalle " & _
               "WHERE (Estado = 1)"
    RsMat.Open GlSqlAux, db, adOpenStatic
    tdbdMat.DataSource = RsMat
    '--
    Set txtDescripcion.DataSource = RsIngreso
    Set txtNoIngreso.DataSource = RsIngreso
    Set txtNoLicitacion.DataSource = RsIngreso
    Set dtpFechIngreso.DataSource = RsIngreso
    '--
    Screen.MousePointer = vbDefault
End Sub

Private Function valida() As Boolean
Dim rs As ADODB.Recordset
    valida = False
    If Trim(txtDescripcion.Text) = "" Then
        MsgBox "Ingrese la Descripción del Ingreso.", vbExclamation + vbOKOnly, "Atención"
        txtDescripcion.SetFocus
        Exit Function
    End If
    If Trim(txtNoLicitacion.Text) = "" Then
        MsgBox "Ingrese el No. de Licitación de la cual se registra el Ingreso.", vbExclamation + vbOKOnly, "Atención"
        txtNoLicitacion.SetFocus
        Exit Function
    End If
    GlSqlAux = "SELECT Count(*) As Cuantos FROM ALIngresoAlm WHERE NoLicitacion = '" & txtNoLicitacion.Text & "'"
    Set rs = New ADODB.Recordset
    rs.Open GlSqlAux, db, adOpenStatic
    If rs!Cuantos > 0 And estado = 1 Then
        MsgBox "El No. de Licitación '" & txtNoLicitacion.Text & "', ya se encuentra registrado..", vbExclamation + vbOKOnly, "Atención"
        txtNoLicitacion.SetFocus
        Exit Function
    End If
    rs.Close
    Set rs = Nothing
    If IsNull(dtpFechIngreso.Value) Then
        MsgBox "Ingrese la fecha de Ingreso de Materiales a Almacen.", vbExclamation + vbOKOnly, "Atención"
        dtpFechIngreso.SetFocus
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
  RsIngreso.Close
  rsProv.Close
  rsNada.Close
  RsDet.Close
  RsMat.Close
End Sub


Private Sub RsIngreso_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If RsIngreso.BOF Or RsIngreso.EOF Then
        If RsIngreso.BOF And RsIngreso.EOF Then
            txtEstado.Text = ""
            txtDescripcion.Text = ""
            txtNoIngreso.Text = ""
            txtNoLicitacion.Text = ""
            dtpFechIngreso.Value = Null
            Set TDBGDet.DataSource = rsNada
            lblEstadoRs.Caption = "Registro: 0 de 0"
        Else
            Exit Sub
        End If
    Else
        
        lblEstadoRs.Caption = "Registro: " & RsIngreso.AbsolutePosition & " de " & RsIngreso.RecordCount
        ' Cargamos el Detalle del Ingreso
        GlSqlAux = "DELETE ALAuxIngresoAlmDet WHERE Usuario = '" & NombreTerminal & "'"
        db.Execute GlSqlAux
        '--
        If estado = 1 Then
            txtEstado.Text = "SIN VERIFICAR"
        Else
            GlSqlAux = "INSERT INTO ALAuxIngresoAlmDet (Usuario, CodProveedor, NombProv, CantidadCaj, CantidadEj, UnidadCaja, CodArt, DescArt, PesoKgs, PrecioSus, PesoTotalKgs, PrecioTotalSus, TipoCambio) " & _
                       "SELECT '" & NombreTerminal & "',CodProveedor, Descripcion_Larga, CantidadCaj, CantidadEj, UnidadCaja, CodArt, DescArt, PesoKgs, PrecioSus, PesoTotalKgs, PrecioTotalSus, TipoCambio " & _
                       "FROM ALIngresoAlmDet INNER JOIN ac_Proveedor ON ALIngresoAlmDet.CodProveedor = ac_Proveedor.Ruc_Id WHERE IdIngreso = " & RsIngreso!IdIngreso
            db.Execute GlSqlAux
            '--
            cmdEliminar.Enabled = Not CBool(RsIngreso!Verificado)
            cmdEditar.Enabled = Not CBool(RsIngreso!Verificado)
            cmdVerificar.Enabled = Not CBool(RsIngreso!Verificado)
            '--
            txtEstado.Text = IIf(CBool(RsIngreso!Verificado), "VERIFICADO", "SIN VERIFICAR")
        End If
        '--
        GlSqlAux = "SELECT * FROM ALAuxIngresoAlmDet WHERE Usuario = '" & NombreTerminal & "'"
        Set RsDet = New ADODB.Recordset
        RsDet.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
        Set AdoDetalle.Recordset = RsDet
        Set TDBGDet.DataSource = RsDet
        Totales
    End If
End Sub
Private Sub tdbdMat_DropDownClose()
    TDBGDet.Columns("DescArt").Value = tdbdMat.Columns("DescDetalle").Value
End Sub

Private Sub tdbdProv_DropDownClose()
    TDBGDet.Columns("NombProv").Value = tdbdProv.Columns("Descripcion_Larga").Value
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
        CantCaja = CantCaja + IIf(IsNull(rs!CantidadCaj), 0, rs!CantidadCaj)
        CantEje = CantEje + IIf(IsNull(rs!CantidadEj), 0, rs!CantidadEj)
        rs.MoveNext
    Wend
    TDBGDet.Columns("DescArt").FooterText = "TOTALES"
    TDBGDet.Columns("PrecioSus").FooterText = Format(PrecioSus, "###,###,##0.00") & " $us."
    TDBGDet.Columns("PesoKgs").FooterText = Format(PesoKgs, "###,###,##0.00") & " Kgs."
    TDBGDet.Columns("PrecioTotalSus").FooterText = Format(PrecioTotalSus, "###,###,##0.00") & " $us."
    TDBGDet.Columns("PesoTotalKgs").FooterText = Format(PesoTotalKgs, "###,###,##0.00") & " Kgs"
    TDBGDet.Columns("CantidadCaj").FooterText = Format(CantCaja, "###,###,##0")
    TDBGDet.Columns("CantidadEj").FooterText = Format(CantEje, "###,###,##0")
End Sub
Private Function ValidarVerificar() As Boolean
Dim rs As ADODB.Recordset
    CadenaError = ""
    ValidarVerificar = True
    ' Validamos que el Detalle este Completo
    Set rs = AdoDetalle.Recordset.Clone
    While Not rs.EOF
        If Trim(rs!CodArt) = "" Then
            CadenaError = CadenaError & vbTab & "Falta Codigo del Item." & vbCrLf
            ValidarVerificar = False
        End If
        If CLng(rs!CantidadCaj) <= 0 Then
            CadenaError = CadenaError & vbTab & "Falta Cantidad de Cajas del Item '" & rs!CodArt & "'." & vbCrLf
            ValidarVerificar = False
        End If
        If CLng(rs!CantidadEj) <= 0 Then
            CadenaError = CadenaError & vbTab & "Falta Cantidad de Cajas del Item '" & rs!CodArt & "'." & vbCrLf
            ValidarVerificar = False
        End If
        If rs!CantidadCaj > 0 Then
            If (rs!CantidadEj Mod rs!CantidadCaj) > 0 Then
                CadenaError = CadenaError & vbTab & "La Cantidad de Cajas del Item '" & rs!CodArt & "', no tiene relación con el Número de Unidades." & vbCrLf & vbTab & vbTab & "1 Caja = " & rs!CantidadEj / CCur(rs!CantidadCaj) & " Unidades ?"
                ValidarVerificar = False
            End If
        End If
        If CCur(rs!PesoKgs) <= 0 Then
            CadenaError = CadenaError & vbTab & "Falta el Peso(Kgs) del Item '" & rs!CodArt & "'." & vbCrLf
            ValidarVerificar = False
        End If
        If CCur(rs!PrecioSus) <= 0 Then
            CadenaError = CadenaError & vbTab & "Falta el Precio($us) del Item '" & rs!CodArt & "'." & vbCrLf
            ValidarVerificar = False
        End If
        If CCur(rs!PesoTotalKgs) <= 0 Then
            CadenaError = CadenaError & vbTab & "Falta el Peso Total(Kgs) del Item '" & rs!CodArt & "'." & vbCrLf
            ValidarVerificar = False
        End If
        If CCur(rs!PrecioTotalSus) <= 0 Then
            CadenaError = CadenaError & vbTab & "Falta el Precio Total($us) del Item '" & rs!CodArt & "'." & vbCrLf
            ValidarVerificar = False
        End If
        If CCur(rs!TipoCambio) <= 0 Then
            CadenaError = CadenaError & vbTab & "Falta el Tipo de Cambio del Item '" & rs!CodArt & "'." & vbCrLf
            ValidarVerificar = False
        End If
        CadenaError = CadenaError & "------------------------------------" & vbCrLf
        rs.MoveNext
    Wend
End Function

Private Sub TdbgIngreso_FetchRowStyle(ByVal Split As Integer, BookMark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If CBool(tdbgIngreso.Columns("Verificado").CellValue(BookMark)) Then
        RowStyle.BackColor = &H8000000D
        RowStyle.ForeColor = &HC0FFFF
        RowStyle.Font.Bold = True
    Else
        RowStyle.BackColor = &HC0FFFF
        RowStyle.ForeColor = &H8000000D
        RowStyle.Font.Bold = False
    End If
End Sub

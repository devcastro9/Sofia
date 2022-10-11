VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form AlFrm_montador 
   Caption         =   "Crea Material"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   120
   ClientWidth     =   12105
   Icon            =   "AlFrm_montador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   12105
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid tdbgArt 
      Height          =   5775
      Left            =   0
      TabIndex        =   33
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   10186
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
         DataField       =   "codgrupo"
         Caption         =   "Grupo"
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
         DataField       =   "codDetalle"
         Caption         =   "SubGrupo"
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
         DataField       =   "descDetalle"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraOpciones 
      Height          =   6315
      Left            =   10800
      TabIndex        =   16
      Top             =   960
      Width           =   1140
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   855
         Left            =   120
         Picture         =   "AlFrm_montador.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   5310
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   855
         Left            =   120
         Picture         =   "AlFrm_montador.frx":10D4
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4470
         Width           =   855
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Borrar"
         Height          =   855
         Left            =   120
         Picture         =   "AlFrm_montador.frx":12DE
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3630
         Width           =   855
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   855
         Left            =   120
         Picture         =   "AlFrm_montador.frx":19C8
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2790
         Width           =   855
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   855
         Left            =   120
         Picture         =   "AlFrm_montador.frx":1BD2
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1950
         Width           =   855
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "Modificar"
         Height          =   855
         Left            =   120
         Picture         =   "AlFrm_montador.frx":1DDC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1110
         Width           =   855
      End
      Begin VB.CommandButton CmdAnadir 
         Caption         =   "Adicionar"
         Height          =   855
         Left            =   120
         MousePointer    =   4  'Icon
         Picture         =   "AlFrm_montador.frx":1FE6
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Frame FraArticulos 
      BorderStyle     =   0  'None
      Height          =   6090
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Width           =   7980
      Begin VB.TextBox TxtUnidad 
         DataField       =   "Unidad"
         DataSource      =   "AdoArt"
         Height          =   300
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   15
         Top             =   5160
         Width           =   1365
      End
      Begin VB.TextBox TxtDescripcion 
         DataField       =   "DescDetalle"
         DataSource      =   "AdoArt"
         Height          =   645
         Left            =   2160
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   3915
         Width           =   5745
      End
      Begin VB.CheckBox chkEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Activo"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   510
         TabIndex        =   12
         Top             =   4875
         Width           =   1185
      End
      Begin VB.TextBox TxtGrupo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         DataField       =   "CodGrupo"
         DataSource      =   "AdoArt"
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
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "11"
         Top             =   3915
         Width           =   615
      End
      Begin VB.TextBox TxtDetalle 
         BackColor       =   &H00FFFFFF&
         DataField       =   "CodDetalle"
         DataSource      =   "AdoArt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   705
         MaxLength       =   5
         TabIndex        =   10
         Top             =   3915
         Width           =   1275
      End
      Begin VB.TextBox txtUnidadCaja 
         Alignment       =   1  'Right Justify
         DataField       =   "UnidadCaja"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         DataSource      =   "AdoArt"
         Height          =   300
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   9
         Top             =   5160
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txtStockMin 
         Alignment       =   1  'Right Justify
         DataField       =   "StockMin"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         DataSource      =   "AdoArt"
         Height          =   300
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "0"
         Top             =   5160
         Visible         =   0   'False
         Width           =   1365
      End
      Begin MSComctlLib.TreeView trv 
         Height          =   2265
         Left            =   90
         TabIndex        =   13
         Top             =   570
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   3995
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imlMaterial"
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0080FFFF&
         Height          =   135
         Left            =   90
         TabIndex        =   32
         Top             =   5880
         Width           =   7815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código Grupo"
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   90
         TabIndex        =   31
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label TDBFrame3D3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      Código      Detalle"
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   840
         TabIndex        =   30
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label TDBFrame3D8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stock Mínimo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5040
         TabIndex        =   29
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label TDBFrame3D7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidades x Caja"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3600
         TabIndex        =   28
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label TDBFrame3D5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidad"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label TDBFrame3D4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DESCRIPCION"
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   2160
         TabIndex        =   26
         Top             =   3360
         Width           =   5745
      End
      Begin VB.Label TDBFrame3D1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DETALLE"
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   90
         TabIndex        =   25
         Top             =   3000
         Width           =   7815
      End
      Begin VB.Label TDBFrame3D2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GRUPOS"
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   90
         TabIndex        =   24
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12105
      TabIndex        =   3
      Top             =   7920
      Width           =   12105
      Begin VB.Frame Frame4 
         Height          =   60
         Left            =   1215
         TabIndex        =   4
         Top             =   255
         Width           =   8370
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clasificador"
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
         Left            =   9675
         TabIndex        =   6
         Top             =   90
         Width           =   1845
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clasificador"
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
         Left            =   9660
         TabIndex        =   5
         Top             =   75
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   990
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   12045
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12105
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE SUBGRUPO"
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
         Left            =   15
         TabIndex        =   1
         Top             =   90
         Width           =   4095
      End
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE MONTADOR"
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
         Left            =   -45
         TabIndex        =   2
         Top             =   120
         Width           =   4185
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   0
         Picture         =   "AlFrm_montador.frx":22F0
         Top             =   -120
         Width           =   11640
      End
   End
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   375
      Left            =   60
      Top             =   6870
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
   Begin MSComctlLib.ImageList imlMaterial 
      Left            =   3975
      Top             =   3180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AlFrm_montador.frx":29360
            Key             =   "Grupos"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AlFrm_montador.frx":29C3A
            Key             =   "NoElegido"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AlFrm_montador.frx":2A514
            Key             =   "Elegido"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "AlFrm_montador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsgrupo As ADODB.Recordset
Dim RsArt As ADODB.Recordset
Dim rsNada As ADODB.Recordset
'--------
Dim estado As Integer ' 0 navegar, 1 Agregar, 2 Editar
'--
Dim ClBuscaGrid As ClBuscaEnGridExterno
Public Sub ALPrincipal(QEstado As Integer)
    '
    Screen.MousePointer = vbHourglass
    estado = QEstado
    '
    Select Case estado
        Case 0
            Set RsArt = New ADODB.Recordset
            GlSqlAux = "SELECT * FROM ALCLDetalle WHERE coddetalle = ISNULL(coddetalle, NULL)"
            RsArt.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
            If RsArt.RecordCount > 0 Then
               GlHayRegs = True  'Variable global
            Else
               GlHayRegs = False
            End If
            BotonesNavegar Me
            FraArticulos.Enabled = False
            Set AdoArt.Recordset = RsArt
        Case 1
                    
        Case 2
        
    End Select
    '
    Screen.MousePointer = vbDefault
    Me.Show
End Sub

Private Sub AdoArt_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If AdoArt.Recordset.BOF Or AdoArt.Recordset.EOF Then
        If AdoArt.Recordset.BOF And AdoArt.Recordset.EOF Then
            TxtGrupo.Text = ""
            TxtDetalle.Text = ""
            Txtdescripcion.Text = ""
            TxtUnidad.Text = ""
            chkEstado.Value = vbUnchecked
            AdoArt.Caption = "Registro: 0 de 0"
            BuscaNodo "rupo"
        Else
            Exit Sub
        End If
    Else
        AdoArt.Caption = "Registro: " & AdoArt.Recordset.AbsolutePosition & " de " & AdoArt.Recordset.RecordCount
        chkEstado.Value = IIf(CBool(AdoArt.Recordset!estado), vbChecked, vbUnchecked)
        BuscaNodo AdoArt.Recordset!CodGrupo
    End If
End Sub
Private Sub CmdAnadir_Click()
    Set tdbgArt.DataSource = rsNada
    AdoArt.Recordset.AddNew
    estado = 1
    BotonesEditar Me
    FraArticulos.Enabled = True
    trv.SetFocus
    BuscaNodo "rupo"
    txtStockMin.Text = 0
    txtUnidadCaja.Text = 0
End Sub

Private Sub CmdBuscar_Click()
  Set ClBuscaGrid = New ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.QueryUtilizado = GlSqlAux
  ClBuscaGrid.Título = "Elija un Detalle"
  ClBuscaGrid.EsTdbGrid = True
  Set ClBuscaGrid.GridTrabajo = tdbgArt
  Set ClBuscaGrid.RecordsetTrabajo = AdoArt.Recordset
  ClBuscaGrid.Ejecutar
'  If ClBuscaGrid.ElegidoCol1 <> "" Then
'    AdoArt.Recordset.Filter = adFilterNone
'    AdoArt.Recordset.MoveFirst
'    AdoArt.Recordset.Find "CodGrupo + '-' + CodDetalle   = " & ClBuscaGrid.ElegidoCol1 & " - " & ClBuscaGrid.ElegidoCol2 & ""
'  End If

End Sub
Private Sub cmdCancelar_Click()
On Error GoTo Que_Error
    Screen.MousePointer = vbHourglass
    If AdoArt.Recordset.EditMode <> adEditNone Then AdoArt.Recordset.CancelUpdate
    AdoArt.Recordset.Requery
    AdoArt.Caption = "Registro: " & CStr(AdoArt.Recordset.AbsolutePosition) & " de " & AdoArt.Recordset.RecordCount
    BotonesNavegar Me
    FraArticulos.Enabled = False
    Set tdbgArt.DataSource = AdoArt
    Screen.MousePointer = vbDefault
    estado = 0
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub
Private Sub Cmdeditar_Click()
On Error GoTo Que_Error
    '
    Screen.MousePointer = vbHourglass
    BotonesEditar Me
    estado = 2
    FraArticulos.Enabled = True
    AdoArt.Caption = "Editando Registro..."
    Screen.MousePointer = vbDefault
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub
Private Sub cmdEliminar_Click()
On Error GoTo Que_Error
    If Not GlHayRegs Then
        MsgBox "No existen registro para eliminar", vbExclamation + vbOKOnly, "Atención"
        Exit Sub
    End If
    If ExisteDetalle(AdoArt.Recordset!CodGrupo & "-" & AdoArt.Recordset!codDetalle) Then MsgBox "No se puede eliminar el Detalle seleccionado ya que se tiene registro de Movimientos en Almacen.", vbInformation + vbOKOnly, "Atención": Exit Sub
    If MsgBox("¿ Está seguro que se va a borrar el registro visualizado ?", vbExclamation + vbOKCancel, "Atención") = vbOK Then
        Screen.MousePointer = vbHourglass
        AdoArt.Recordset.Delete
        AdoArt.Recordset.MoveNext
        If AdoArt.Recordset.EOF Then
          If AdoArt.Recordset.RecordCount > 0 Then
            AdoArt.Recordset.MoveLast
          Else
            GlHayRegs = False
            AdoArt.Refresh
          End If
        End If
        Screen.MousePointer = vbDefault
    End If
    BotonesNavegar Me
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub
Private Sub CmdGrabar_Click()
On Error GoTo QError
    If valida Then
        Screen.MousePointer = vbHourglass
        ' Empezar a grabar
        '*********************************
        db.BeginTrans
        ' Campos no ligados
        AdoArt.Recordset!estado = IIf(chkEstado.Value = vbChecked, 1, 0)
        '*********************************
        ' Grabar
        AdoArt.Recordset.Update
        db.CommitTrans
    '*********************************
        AdoArt.Caption = "Registro: " & CStr(AdoArt.Recordset.AbsolutePosition) & " de " & AdoArt.Recordset.RecordCount
        ' Colocar los botones en modo navegar
        GlHayRegs = True
        BotonesNavegar Me
        FraArticulos.Enabled = False
        Screen.MousePointer = vbDefault
        AdoArt.Refresh
        Set tdbgArt.DataSource = AdoArt
        estado = 0
    End If
    Exit Sub
QError:
    ' Manejo de errores
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
    db.RollbackTrans
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdRefrescar_Click()
On Error GoTo Que_Error
    Screen.MousePointer = vbHourglass
    AdoArt.Recordset.Requery
    Screen.MousePointer = vbDefault
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
Dim Nodo As Node
    Me.Top = 0
    Me.Left = 0
    Screen.MousePointer = vbHourglass
    ' Cargamos el Arbol
    ' Nodo Principal
    Set Nodo = trv.Nodes.Add(, , "Grupo", "Grupos", "Grupos")
    
    Nodo.Expanded = True
    Nodo.Bold = True
    Set rsgrupo = New ADODB.Recordset
    rsgrupo.Open "SELECT * FROM ALClGrupo ORDER BY CodGrupo", db, adOpenStatic
    If rsgrupo.RecordCount > 0 Then
      rsgrupo.MoveFirst
      While Not rsgrupo.EOF
        Set Nodo = trv.Nodes.Add("Grupo", tvwChild, "D" & Trim(rsgrupo!CodGrupo), rsgrupo!descgrupo, "NoElegido", "Elegido")
        rsgrupo.MoveNext
      Wend
    Else
        trv.Nodes(1).Text = "No Existen Grupos Creados..."
    End If
    '
    Screen.MousePointer = vbDefault
End Sub
Private Function valida() As Boolean
    valida = False
    If Trim(TxtGrupo.Text) = "" Then
        MsgBox "Elija el Grupo al Cual pertenece el Detalle.", vbExclamation + vbOKOnly, "Atención"
        trv.SetFocus
        Exit Function
    End If
    If Trim(TxtDetalle.Text) = "" Then
        MsgBox "Ingrese el Codigo del Detalle.", vbExclamation + vbOKOnly, "Atención"
        TxtDetalle.SetFocus
        Exit Function
    End If
    If Trim(Txtdescripcion.Text) = "" Then
        MsgBox "Ingrese la Descripción del Detalle.", vbExclamation + vbOKOnly, "Atención"
        Txtdescripcion.SetFocus
        Exit Function
    End If
    If Trim(TxtUnidad.Text) = "" Then
        MsgBox "Ingrese la Unidad del Detalle.", vbExclamation + vbOKOnly, "Atención"
        TxtUnidad.SetFocus
        Exit Function
    End If
    If Trim(txtUnidadCaja.Text) = "" Then
        MsgBox "Ingrese la Unidad por Caja del Detalle.", vbExclamation + vbOKOnly, "Atención"
        TxtUnidad.SetFocus
        Exit Function
    End If
    If Trim(txtStockMin.Text) = "" Then
        MsgBox "Ingrese el Stock Mínimo del Detalle.", vbExclamation + vbOKOnly, "Atención"
        TxtUnidad.SetFocus
        Exit Function
    End If
    valida = True
End Function
Private Sub Form_Unload(Cancel As Integer)
  Set ClBuscaGrid = Nothing
End Sub

Private Sub trv_NodeClick(ByVal Node As MSComctlLib.Node)
    If InStr(Node.Key, "G") = 0 Then
        TxtGrupo.Text = Mid(Node.Key, 2)
    Else
        TxtGrupo.Text = ""
    End If
End Sub
Private Sub BuscaNodo(QNodo As String)
Dim Nodo As Node
Dim Indice As Integer
    For Indice = 1 To trv.Nodes.Count
        If Mid(trv.Nodes(Indice).Key, 2) = QNodo Then
            trv.Nodes(Indice).Selected = True
            Exit For
        End If
    Next
End Sub
Private Sub txtStockMin_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]", KeyAscii, 0)
End Sub
Private Sub txtUnidadCaja_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]", KeyAscii, 0)
End Sub
Private Function ExisteDetalle(codDetalle As String) As Boolean
Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ALIngresoAlmDet WHERE CodArt = '" & codDetalle & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteDetalle = rs!Cuantos > 0
End Function


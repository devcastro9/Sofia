VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmHistoricoTributosFiscales 
   Caption         =   "Reporte de Históricos de Tributos Fiscales"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmHistoricoTributosFiscales.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CryHistorico 
      Left            =   7215
      Top             =   2925
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame FraOpciones 
      Height          =   7365
      Left            =   90
      TabIndex        =   10
      Top             =   900
      Width           =   1290
      Begin VB.CommandButton CmdRestaurar 
         Caption         =   "Restaurar"
         Height          =   885
         Left            =   75
         TabIndex        =   22
         Top             =   3585
         Width           =   1170
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   885
         Left            =   75
         TabIndex        =   21
         Top             =   2700
         Width           =   1170
      End
      Begin VB.CommandButton CmdBuscarFecha 
         Caption         =   "Buscar por fecha"
         Height          =   885
         Left            =   75
         TabIndex        =   20
         Top             =   1815
         Width           =   1170
      End
      Begin VB.CommandButton CmdImprimirSel 
         Caption         =   "Imprimir Seleccionados"
         Height          =   795
         Left            =   75
         Picture         =   "FrmHistoricoTributosFiscales.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1005
         Width           =   1170
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   150
         Picture         =   "FrmHistoricoTributosFiscales.frx":1534
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6330
         Width           =   945
      End
      Begin VB.CommandButton CmdInforme 
         Caption         =   "Imprimir General"
         Height          =   795
         Left            =   75
         Picture         =   "FrmHistoricoTributosFiscales.frx":1976
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   210
         Width           =   1170
      End
   End
   Begin VB.CommandButton CmdElegir 
      Caption         =   ">>"
      Height          =   555
      Left            =   7095
      TabIndex        =   9
      Top             =   1440
      Width           =   660
   End
   Begin VB.CommandButton CmdDevolver 
      Caption         =   "<<"
      Height          =   525
      Left            =   7095
      TabIndex        =   8
      Top             =   2100
      Width           =   660
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   4620
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.Label Label8 
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
         Top             =   480
         Width           =   1110
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   4
         Top             =   525
         Width           =   2460
      End
      Begin VB.Label LblUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   2
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HISTORICO DE TRIBUTOS FISCALES"
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
         Left            =   3390
         TabIndex        =   1
         Top             =   120
         Width           =   5535
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   9210
         TabIndex        =   3
         Top             =   480
         Width           =   1275
      End
      Begin VB.Image Image3 
         Height          =   960
         Left            =   0
         Picture         =   "FrmHistoricoTributosFiscales.frx":1FE0
         Top             =   0
         Width           =   12600
      End
   End
   Begin MSDataGridLib.DataGrid DtGTributosFiscales 
      Height          =   6105
      Left            =   1380
      TabIndex        =   6
      Top             =   1380
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   10769
      _Version        =   393216
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin MSDataGridLib.DataGrid DtGSeleccionados 
      Height          =   6105
      Left            =   7965
      TabIndex        =   7
      Top             =   1425
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   10769
      _Version        =   393216
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin MSComCtl2.DTPicker DTPFechaInicio 
      Height          =   375
      Left            =   1425
      TabIndex        =   16
      Top             =   7905
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   248643585
      CurrentDate     =   36413
   End
   Begin MSComCtl2.DTPicker DTPFechaFin 
      Height          =   375
      Left            =   3390
      TabIndex        =   17
      Top             =   7905
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   248643585
      CurrentDate     =   36413
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha Inicio"
      Height          =   240
      Left            =   1470
      TabIndex        =   19
      Top             =   7680
      Width           =   1590
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Fin"
      Height          =   240
      Left            =   3435
      TabIndex        =   18
      Top             =   7695
      Width           =   1590
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccionados"
      Height          =   165
      Left            =   8025
      TabIndex        =   14
      Top             =   1080
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Histórico General"
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   1050
      Width           =   3990
   End
End
Attribute VB_Name = "FrmHistoricoTributosFiscales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsNada As ADODB.Recordset
Private Sub CmdBuscarFecha_Click()
    Set rsHistorico = New ADODB.Recordset
    If rsHistorico.State = 1 Then rsHistorico.Close
    rsHistorico.Open "SELECT * FROM to_CajaHistoricoPagos_fechas WHERE Fecha_Impresion>= '" & Str(DTPFechaInicio.Value) & "'  and Fecha_Impresion <= '" & Str(DTPFechaFin.Value) & "'", db, adOpenKeyset, adLockOptimistic
    If rsHistorico.RecordCount > 0 Then
        Set DtGTributosFiscales.DataSource = rsHistorico
        While Not rsHistorico.EOF
                db.Execute "insert into to_CajaHistoricoPagos (T_beneficiario, codigo_pago, org_codigo, Nro_doc, monto_bolivianos, porcentaje_1, porcentaje_2, fecha_impresion) " & _
                           "values  ('" & rsHistorico!T_beneficiario & "','" & rsHistorico!codigo_pago & "', '" & rsHistorico!org_codigo & "','" & rsHistorico!Nro_doc & "'," & rsHistorico!monto_Bolivianos & " , " & rsHistorico!Porcentaje_1 & ", " & rsHistorico!Porcentaje_2 & ", '" & Date & "') "
                rsHistorico.MoveNext
        Wend
    Else
       MsgBox "No existen registros", vbInformation + vbCritical, "Validación"
       Set DtGTributosFiscales.DataSource = rsNada
    End If
End Sub

Private Sub cmdElegir_Click()
Dim rsSel As New ADODB.Recordset
    
    db.Execute "insert into to_CajaHistoricoPagos (T_beneficiario, codigo_pago, org_codigo, Nro_doc, monto_bolivianos, par_codigo, Porcentaje_1, Porcentaje_2 )" & _
                 "values  ('" & DtGTributosFiscales.Columns(0) & "','" & DtGTributosFiscales.Columns(1) & "', '" & DtGTributosFiscales.Columns(2) & "','" & DtGTributosFiscales.Columns(3) & "'," & DtGTributosFiscales.Columns(4) & ", '" & DtGTributosFiscales.Columns(5) & "', " & DtGTributosFiscales.Columns(6) & ", " & DtGTributosFiscales.Columns(7) & " ) "
    Set rsSel = New ADODB.Recordset
    If rsSel.State = 1 Then rsSel.Close
    rsSel.Open "SELECT * FROM to_CajaHistoricoPagos ", db, adOpenKeyset, adLockOptimistic
    If rsSel.RecordCount > 0 Then
        Set DtGSeleccionados.DataSource = rsSel
    Else
        Set DtGSeleccionados.DataSource = rsNada
    End If
End Sub

Private Sub CmdImprimirSel_Click()
'RepSoloTributos.Show
 CryHistorico.ReportFileName = App.Path & "\FormsTesoreria\EntregaCheques\Impresiones\Rpt_SoloTributos.rpt"
 iResult = CryHistorico.PrintReport
 If iResult <> 0 Then
    MsgBox CryHistorico.LastErrorNumber & " : " & CryHistorico.LastErrorString, vbCritical + vbOKOnly, "Error..."
 End If
End Sub

Private Sub CmdInforme_Click()
'RepSoloTributos.Show
 CryHistorico.ReportFileName = App.Path & "\FormsTesoreria\EntregaCheques\Impresiones\Rpt_SoloTributos.rpt"
 iResult = CryHistorico.PrintReport
 If iResult <> 0 Then
    MsgBox CryHistorico.LastErrorNumber & " : " & CryHistorico.LastErrorString, vbCritical + vbOKOnly, "Error..."
 End If
End Sub

Private Sub CmdLimpiar_Click()
Dim rsSel As ADODB.Recordset
    db.Execute "delete from to_CajaHistoricoPagos"
    Set rsSel = New ADODB.Recordset
    If rsSel.State = 1 Then rsSel.Close
    rsSel.Open "SELECT * FROM to_CajaHistoricoPagos ", db, adOpenKeyset, adLockOptimistic
    If rsSel.RecordCount > 0 Then
        Set DtGSeleccionados.DataSource = rsSel
    Else
        Set DtGSeleccionados.DataSource = rsNada
    End If
End Sub

Private Sub CmdRestaurar_Click()
Dim rsHistorico As New ADODB.Recordset
    Set rsHistorico = New ADODB.Recordset
    If rsHistorico.State = 1 Then rsHistorico.Close
    rsHistorico.Open "SELECT * FROM to_CajaHistoricoPagos_Fechas", db, adOpenKeyset, adLockOptimistic
    If rsHistorico.RecordCount > 0 Then
       Set DtGTributosFiscales.DataSource = rsHistorico
    Else
       Set DtGTributosFiscales.DataSource = rsNada
       MsgBox "No existen registros", vbCritical + vbDefaultButton1, "VALIDACION DE DATOS"
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Load()
db.Execute "delete from to_CajaHistoricoPagos "
'Abriendo histórico por fechas
Set rsHistorico = New ADODB.Recordset
If rsHistorico.State = 1 Then rsHistorico.Close
rsHistorico.Open "SELECT * FROM to_CajaHistoricoPagos_Fechas", db, adOpenKeyset, adLockOptimistic
If rsHistorico.RecordCount > 0 Then
   Set DtGTributosFiscales.DataSource = rsHistorico
Else
   Set DtGTributosFiscales.DataSource = rsNada
End If
'Colocando en la fecha actual
DTPFechaInicio.Value = Date
DTPFechaFin.Value = Date
End Sub


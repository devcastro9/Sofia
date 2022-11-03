VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmDesplegado 
   Caption         =   "Desplegando Cheques"
   ClientHeight    =   8535
   ClientLeft      =   180
   ClientTop       =   1815
   ClientWidth     =   11400
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "CRITERIO DE BUSQUEDA"
      Height          =   1110
      Left            =   1845
      TabIndex        =   8
      Top             =   1050
      Width           =   11280
      Begin VB.TextBox TxtValor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   14
         Text            =   "S"
         Top             =   630
         Width           =   1575
      End
      Begin VB.ComboBox CmbOperador 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmDesplegado.frx":0000
         Left            =   2025
         List            =   "FrmDesplegado.frx":0013
         Style           =   1  'Simple Combo
         TabIndex        =   13
         Text            =   "="
         Top             =   630
         Width           =   1065
      End
      Begin VB.ComboBox CmbCampo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   630
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPFechaInicio 
         Height          =   375
         Left            =   5340
         TabIndex        =   20
         Top             =   540
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   24903681
         CurrentDate     =   36413
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   375
         Left            =   7275
         TabIndex        =   21
         Top             =   555
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   24903681
         CurrentDate     =   36413
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio"
         Height          =   240
         Left            =   5355
         TabIndex        =   23
         Top             =   330
         Width           =   1590
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Fin"
         Height          =   240
         Left            =   7320
         TabIndex        =   22
         Top             =   345
         Width           =   1590
      End
      Begin VB.Label LblValor 
         Caption         =   "Valor"
         Height          =   285
         Left            =   3270
         TabIndex        =   11
         Top             =   255
         Width           =   675
      End
      Begin VB.Label LblOperador 
         Caption         =   "Operador"
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   315
         Width           =   885
      End
      Begin VB.Label LblCampo 
         Caption         =   "Campo"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   255
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Height          =   4905
      Left            =   1845
      TabIndex        =   25
      Top             =   2145
      Width           =   2475
      Begin VB.OptionButton Option10 
         Caption         =   "Option1"
         Height          =   345
         Left            =   2010
         TabIndex        =   40
         Top             =   1725
         Width           =   300
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Option1"
         Height          =   345
         Left            =   1440
         TabIndex        =   39
         Top             =   1710
         Width           =   300
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Option1"
         Height          =   345
         Left            =   2025
         TabIndex        =   38
         Top             =   1440
         Width           =   300
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Option1"
         Height          =   345
         Left            =   1455
         TabIndex        =   37
         Top             =   1425
         Width           =   300
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option1"
         Height          =   345
         Left            =   2025
         TabIndex        =   36
         Top             =   1095
         Width           =   300
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option1"
         Height          =   345
         Left            =   1455
         TabIndex        =   35
         Top             =   1080
         Width           =   300
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option1"
         Height          =   345
         Left            =   2025
         TabIndex        =   34
         Top             =   795
         Width           =   300
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option1"
         Height          =   345
         Left            =   1455
         TabIndex        =   33
         Top             =   780
         Width           =   300
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option1"
         Height          =   345
         Left            =   2025
         TabIndex        =   32
         Top             =   495
         Width           =   300
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   345
         Left            =   1455
         TabIndex        =   31
         Top             =   495
         Width           =   300
      End
      Begin VB.Label Label9 
         Caption         =   "No"
         Height          =   210
         Left            =   2025
         TabIndex        =   42
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label8 
         Caption         =   "Si"
         Height          =   210
         Left            =   1455
         TabIndex        =   41
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label7 
         Caption         =   "Devuelto"
         Height          =   270
         Left            =   135
         TabIndex        =   30
         Top             =   1785
         Width           =   1470
      End
      Begin VB.Label Label6 
         Caption         =   "Anulado"
         Height          =   270
         Left            =   165
         TabIndex        =   29
         Top             =   1440
         Width           =   1470
      End
      Begin VB.Label Label5 
         Caption         =   "Cobrado"
         Height          =   270
         Left            =   150
         TabIndex        =   28
         Top             =   1125
         Width           =   1470
      End
      Begin VB.Label Label4 
         Caption         =   "Entregado"
         Height          =   270
         Left            =   165
         TabIndex        =   27
         Top             =   810
         Width           =   1110
      End
      Begin VB.Label Label2 
         Caption         =   "Impreso"
         Height          =   270
         Left            =   165
         TabIndex        =   26
         Top             =   525
         Width           =   1470
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   930
      Left            =   2895
      TabIndex        =   24
      Top             =   7290
      Visible         =   0   'False
      Width           =   2145
   End
   Begin Crystal.CrystalReport CryHis 
      Left            =   1020
      Top             =   7230
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
   Begin VB.Frame Frame3 
      Height          =   6015
      Left            =   345
      TabIndex        =   15
      Top             =   1020
      Width           =   1320
      Begin VB.CommandButton CmdFiltCriterio 
         Caption         =   "Filtra CRITERIO"
         Height          =   975
         Left            =   165
         Picture         =   "FrmDesplegado.frx":002A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1185
         Width           =   1005
      End
      Begin VB.CommandButton CmdAnulaFCriterio 
         Caption         =   "Restaurar"
         Height          =   975
         Left            =   150
         Picture         =   "FrmDesplegado.frx":046C
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2160
         Width           =   1020
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprime Cheque"
         Height          =   975
         Left            =   165
         Picture         =   "FrmDesplegado.frx":08AE
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   210
         Width           =   1005
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   975
         Left            =   165
         Picture         =   "FrmDesplegado.frx":0F18
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4455
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   690
      Left            =   365
      ScaleHeight     =   630
      ScaleWidth      =   12735
      TabIndex        =   1
      Top             =   210
      Width           =   12795
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "IMPRESION HISTORICA DE CHEQUES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   3120
         TabIndex        =   2
         Top             =   135
         Width           =   7320
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4785
      Left            =   4440
      TabIndex        =   0
      Top             =   2265
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   8440
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
   Begin VB.Frame Frame1 
      Caption         =   "BUSQUEDA"
      Height          =   870
      Left            =   1890
      TabIndex        =   3
      Top             =   5985
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton CmdUltimo 
         Caption         =   "Ultimo"
         Height          =   510
         Left            =   3515
         Picture         =   "FrmDesplegado.frx":135A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1095
      End
      Begin VB.CommandButton CmdAnterior 
         Caption         =   "Anterior"
         Height          =   510
         Left            =   1225
         Picture         =   "FrmDesplegado.frx":145C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1095
      End
      Begin VB.CommandButton CmdSiguiente 
         Caption         =   "Siguiente"
         Height          =   510
         Left            =   2370
         Picture         =   "FrmDesplegado.frx":1556
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1095
      End
      Begin VB.CommandButton CmdPrimero 
         Caption         =   "Primero"
         Height          =   510
         Left            =   80
         Picture         =   "FrmDesplegado.frx":1658
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1845
      Top             =   6555
      Visible         =   0   'False
      Width           =   11340
      _ExtentX        =   20003
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\MVB5\Labs\Neptuno.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\MVB5\Labs\Neptuno.mdb"
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
   Begin MSAdodcLib.Adodc AdoAuxiliar 
      Height          =   330
      Left            =   1845
      Top             =   6540
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
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
      Caption         =   "Adodc2"
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
Attribute VB_Name = "FrmDesplegado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsbusca As New ADODB.Recordset
Dim rsauxiliar As New ADODB.Recordset
Dim rsCheque As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
Dim rsEntregados As New ADODB.Recordset

Dim CAMPOS As ADODB.Field
Dim Vquery As String
Public NoRegistros As Long

Private Sub CmdAnulaF_Click()
    rsbusca.Filter = adFilterNone
End Sub
Private Sub CmdAnulaFCriterio_Click()
Dim errCriterio As Variant
If FrmActivacionCheques.OptCheques = True Then
    Vquery = "select * from cns_cheques where  cheque_o_trf= 'C'"
End If
If FrmActivacionCheques.OptTransferencias = True Then
    Vquery = "select * from cns_cheques where  cheque_o_trf= 'T'"
End If

    errCriterio = "Existe Error en el Criterio"
    If rsbusca.State = 1 Then rsbusca.Close
    On Error GoTo Error:
        rsbusca.Open Vquery, db, adOpenStatic, adLockReadOnly
        If rsbusca.RecordCount > 0 Then
            NoRegistros = rsbusca.RecordCount
        Else
            NoRegistros = 0
        End If
        
        Set Adodc1.Recordset = rsbusca
        Set DataGrid1.DataSource = Adodc1
        For Each CAMPOS In rsbusca.Fields
            CmbCampo.AddItem CAMPOS.Name
        Next CAMPOS
        
        DTPFechaInicio.Value = Date
        DTPFechaFin.Value = Date
    Exit Sub
Error:
    MsgBox "Existe error de sintaxis", vbDefaultButton2, "ERROR"
    MsgBox Vquery
End Sub
Private Sub CmdFiltCriterio_Click()
Dim errCriterio As Variant
Dim Sw As Integer


'Validación de fechas
If DTPFechaInicio.Value > DTPFechaFin.Value Or DTPFechaFin.Value < DTPFechaInicio.Value Then
     MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
     Exit Sub
End If

db.Execute "DELETE FROM to_operaciones_cheques"
    Sw = 0
    If rsbusca.State = 1 Then rsbusca.Close
    'esta bien
    Set DataGrid1.DataSource = rsauxiliar
    If CmbCampo.Text <> "" And CmbOperador.Text <> "" And "'" & TxtValor.Text & "'" <> "" Then
        rsbusca.Open Vquery & " and " & CmbCampo.Text & CmbOperador.Text & "'" & TxtValor.Text & "'", db, adOpenStatic, adLockReadOnly
                        'Vaciando los datos al temporal to_operaciones_cheques
                            If CmbCampo.Text = "estado_entregado" Then
                                  While Not rsbusca.EOF
                                      If rsbusca.RecordCount > 0 Then
                                          If rsbusca("Fecha_Entregado") >= DTPFechaInicio.Value And rsbusca("Fecha_Entregado") <= DTPFechaFin.Value Then
                                              db.Execute "insert into to_operaciones_cheques(numero_cheque,fecha_registro,monto_bolivianos,denominacion_beneficiario,cta_descripcion_larga,cta_codigo, codigo_pago,estado_impreso,estado_entregado,estado_cobrado,estado_anulado, estado_devuelto) " & _
                                                         "values ('" & rsbusca!numero_cheque & "','" & rsbusca!fecha_entregado & "','" & rsbusca!monto_bolivianos & "','" & rsbusca!denominacion_beneficiario & "','" & rsbusca!cta_descripcion_larga & "','" & rsbusca!cta_codigo & "', '" & rsbusca!codigo_pago & "','" & rsbusca!estado_impreso & "','" & rsbusca!estado_entregado & "','" & rsbusca!estado_cobrado & "','" & rsbusca!estado_anulado & "', '" & rsbusca!estado_devuelto & "') "
                                              Sw = 1
                                          End If
                                          rsbusca.MoveNext
                                      End If
                                   Wend
                             End If
                             If CmbCampo.Text = "estado_impreso" Then
                                   While Not rsbusca.EOF
                                      If rsbusca.RecordCount > 0 Then
                                          If rsbusca("Fecha_impreso") >= DTPFechaInicio.Value And rsbusca("Fecha_impreso") <= DTPFechaFin.Value Then
                                              db.Execute "insert into to_operaciones_cheques(numero_cheque,fecha_registro,monto_bolivianos,denominacion_beneficiario,cta_descripcion_larga,cta_codigo, codigo_pago,estado_impreso,estado_entregado,estado_cobrado,estado_anulado, estado_devuelto ) " & _
                                                         "values ('" & rsbusca!numero_cheque & "','" & rsbusca!fecha_impreso & "','" & rsbusca!monto_bolivianos & "','" & rsbusca!denominacion_beneficiario & "','" & rsbusca!cta_descripcion_larga & "','" & rsbusca!cta_codigo & "', '" & rsbusca!codigo_pago & "','" & rsbusca!estado_impreso & "','" & rsbusca!estado_entregado & "','" & rsbusca!estado_cobrado & "','" & rsbusca!estado_anulado & "','" & rsbusca!estado_devuelto & "') "
                                              Sw = 1
                                          End If
                                          rsbusca.MoveNext
                                      End If
                                  Wend
                             End If
                             If CmbCampo.Text = "estado_cobrado" Then
                                  While Not rsbusca.EOF
                                      If rsbusca.RecordCount > 0 Then
                                          If rsbusca("Fecha_Cobrado") >= DTPFechaInicio.Value And rsbusca("Fecha_Cobrado") <= DTPFechaFin.Value Then
                                              db.Execute "insert into to_operaciones_cheques(numero_cheque,fecha_registro,monto_bolivianos,denominacion_beneficiario,cta_descripcion_larga,cta_codigo, codigo_pago,estado_impreso,estado_entregado,estado_cobrado,estado_anulado, estado_devuelto) " & _
                                                         "values ('" & rsbusca!numero_cheque & "','" & rsbusca!fecha_cobrado & "','" & rsbusca!monto_bolivianos & "','" & rsbusca!denominacion_beneficiario & "','" & rsbusca!cta_descripcion_larga & "','" & rsbusca!cta_codigo & "', '" & rsbusca!codigo_pago & "','" & rsbusca!estado_impreso & "','" & rsbusca!estado_entregado & "','" & rsbusca!estado_cobrado & "','" & rsbusca!estado_anulado & "', '" & rsbusca!estado_devuelto & "') "
                                              Sw = 1
                                          End If
                                          rsbusca.MoveNext
                                      End If
                                  Wend
                              End If
                              If CmbCampo.Text = "estado_anulado" Then
                                  While Not rsbusca.EOF
                                      If rsbusca.RecordCount > 0 Then
                                          If rsbusca("Fecha_Anulado") >= DTPFechaInicio.Value And rsbusca("Fecha_Anulado") <= DTPFechaFin.Value Then
                                              db.Execute "insert into to_operaciones_cheques(numero_cheque,fecha_registro,monto_bolivianos,denominacion_beneficiario,cta_descripcion_larga,cta_codigo, codigo_pago,estado_impreso,estado_entregado,estado_cobrado,estado_anulado, estado_devuelto) " & _
                                                         "values ('" & rsbusca!numero_cheque & "','" & rsbusca!fecha_anulado & "','" & rsbusca!monto_bolivianos & "','" & rsbusca!denominacion_beneficiario & "','" & rsbusca!cta_descripcion_larga & "','" & rsbusca!cta_codigo & "', '" & rsbusca!codigo_pago & "','" & rsbusca!estado_impreso & "','" & rsbusca!estado_entregado & "','" & rsbusca!estado_cobrado & "','" & rsbusca!estado_anulado & "', '" & rsbusca!estado_devuelto & "') "
                                              Sw = 1
                                          End If
                                          rsbusca.MoveNext
                                      End If
                                  Wend
                              End If
                              If CmbCampo.Text = "estado_devuelto" Then
                                  While Not rsbusca.EOF
                                      If rsbusca.RecordCount > 0 Then
                                          If rsbusca("Fecha_Devuelto") >= DTPFechaInicio.Value And rsbusca("Fecha_Devuelto") <= DTPFechaFin.Value Then
                                              db.Execute "insert into to_operaciones_cheques(numero_cheque,fecha_registro,monto_bolivianos,denominacion_beneficiario,cta_descripcion_larga,cta_codigo, codigo_pago,estado_impreso,estado_entregado,estado_cobrado,estado_anulado, estado_devuelto) " & _
                                                         "values ('" & rsbusca!numero_cheque & "','" & rsbusca!fecha_devuelto & "','" & rsbusca!monto_bolivianos & "','" & rsbusca!denominacion_beneficiario & "','" & rsbusca!cta_descripcion_larga & "','" & rsbusca!cta_codigo & "', '" & rsbusca!codigo_pago & "','" & rsbusca!estado_impreso & "','" & rsbusca!estado_entregado & "','" & rsbusca!estado_cobrado & "','" & rsbusca!estado_anulado & "', '" & rsbusca!estado_devuelto & "') "
                                              Sw = 1
                                          End If
                                          rsbusca.MoveNext
                                      End If
                                  Wend
                               End If
                                  If Sw <> 1 And rsbusca.EOF Then
                                         Set DataGrid1.DataSource = rsNada
                                         NoRegistros = 0
                                         Sw = 0
                                  Else
                                        If rsEntregados.State = 1 Then rsEntregados.Close
                                        rsEntregados.Open "SELECT * from to_operaciones_cheques", db, adOpenKeyset, adLockOptimistic
                                        If rsEntregados.RecordCount > 0 Then
                                               Set Adodc1.Recordset = rsEntregados
                                               Set DataGrid1.DataSource = rsEntregados
                                               NoRegistros = rsbusca.RecordCount
                                        Else
                                               Set Adodc1.Recordset = rsNada
                                               Set DataGrid1.DataSource = rsEntregados
                                        End If
                                  End If
  
                        
                    Else
                        MsgBox "Coloque datos para realizar la búsqueda de registros", vbCritical, "Validación de datos"
                        rsbusca.Open Vquery, db, adOpenStatic, adLockReadOnly
                        If rsbusca.RecordCount > 0 Then
                           NoRegistros = rsbusca.RecordCount
                        Else
                           NoRegistros = 0
                        End If
                        Set Adodc1.Recordset = rsbusca
                        Set DataGrid1.DataSource = rsbusca
           End If
           
End Sub
Private Sub CmdImprimir_Click()
  Dim iResult As String
  Dim Cadena As String
  
  Set rsCheque = New ADODB.Recordset
  If rsCheque.State = 1 Then rsCheque.Close
  rsCheque.Open "SELECT * FROM to_operaciones_cheques", db, adOpenDynamic, adLockOptimistic
  If rsCheque.RecordCount <= 0 Then
    MsgBox "No existen registros para imprimir ", vbCritical + vbDefaultButton1, "Validación de datos"
    Exit Sub
  End If
  If Option1.Value = True And FrmActivacionCheques.OptCheques.Value = True Then
      Cadena = "CHEQUES IMPRESOS"
  End If
  If Option1.Value = True And FrmActivacionCheques.OptTransferencias.Value = True Then
      Cadena = "TRANSFERENCIAS IMPRESAS"
  End If
  If Option2.Value = True And FrmActivacionCheques.OptCheques.Value = True Then
      Cadena = "CHEQUES NO IMPRESOS"
  End If
  If Option2.Value = True And FrmActivacionCheques.OptTransferencias.Value = True Then
      Cadena = "TRANSFERENCIAS NO IMPRESAS"
  End If
  If Option3.Value = True And FrmActivacionCheques.OptCheques.Value = True Then
      Cadena = "CHEQUES ENTREGADOS"
  End If
  If Option3.Value = True And FrmActivacionCheques.OptTransferencias.Value = True Then
      Cadena = "TRANSFERENCIAS ENTREGADAS"
  End If
  If Option4.Value = True And FrmActivacionCheques.OptCheques.Value = True Then
      Cadena = "CHEQUES NO ENTREGADOS"
  End If
  If Option4.Value = True And FrmActivacionCheques.OptTransferencias.Value = True Then
      Cadena = "TRANSFERENCIAS NO ENTREGADAS"
  End If
  If Option5.Value = True And FrmActivacionCheques.OptCheques.Value = True Then
      Cadena = "CHEQUES COBRADOS"
  End If
  If Option5.Value = True And FrmActivacionCheques.OptTransferencias.Value = True Then
      Cadena = "TRANSFERENCIAS COBRADAS"
  End If
  If Option6.Value = True And FrmActivacionCheques.OptCheques.Value = True Then
      Cadena = "CHEQUES NO COBRADOS"
  End If
  If Option6.Value = True And FrmActivacionCheques.OptTransferencias.Value = True Then
      Cadena = "TRANSFERENCIAS NO COBRADAS"
  End If
  If Option7.Value = True And FrmActivacionCheques.OptCheques.Value = True Then
      Cadena = "CHEQUES ANULADOS"
  End If
  If Option7.Value = True And FrmActivacionCheques.OptTransferencias.Value = True Then
      Cadena = "TRANSFERENCIAS ANULADAS"
  End If
  If Option8.Value = True And FrmActivacionCheques.OptCheques.Value = True Then
      Cadena = "CHEQUES NO ANULADOS"
  End If
  If Option8.Value = True And FrmActivacionCheques.OptTransferencias.Value = True Then
      Cadena = "TRANSFERENCIAS NO ANULADAS"
  End If
  If Option9.Value = True And FrmActivacionCheques.OptCheques.Value = True Then
      Cadena = "CHEQUES DEVUELTOS"
  End If
  If Option9.Value = True And FrmActivacionCheques.OptTransferencias.Value = True Then
      Cadena = "TRANSFERENCIAS DEVUELTAS"
  End If
  If Option10.Value = True And FrmActivacionCheques.OptCheques.Value = True Then
      Cadena = "CHEQUES NO DEVUELTOS"
  End If
  If Option10.Value = True And FrmActivacionCheques.OptTransferencias.Value = True Then
      Cadena = "TRANSFERENCIAS NO DEVUELTAS"
  End If
  CryHis.Formulas(1) = "Fecha_Ini='" & FrmDesplegado.DTPFechaInicio.Value & "'"
  CryHis.Formulas(2) = "Fecha_Fin='" & FrmDesplegado.DTPFechaFin.Value & "'"
  CryHis.Formulas(3) = "FTitulo='" & Cadena & "'"
  CryHis.ReportFileName = App.path & "\FormsTesoreria\Operacion de Cheques\Rpt_SatusCheques.rpt"   'gvi "C:\Saf-2000\FormsTesoreria\Operacion de Cheques\Rpt_SatusCheques.rpt"
  iResult = CryHis.PrintReport
  If iResult <> 0 Then
     MsgBox CryHis.LastErrorNumber & " : " & CryHis.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub
Private Sub CmdSalir_Click()
    FrmActivacionCheques.Show
    Unload Me
End Sub

Private Sub CmdSalirPagos_Click()
    Unload Me
End Sub


Private Sub Command1_Click()
Dim x As Variant
    Dim rscheques As New ADODB.Recordset
    rscheques.Open "select * from to_cheques_operaciones where fecha_registro='05/08/2000' ", db, adOpenKeyset, adLockOptimistic
    If rscheques.RecordCount > 0 Then
      While Not rscheques.EOF
        x = rscheques("fecha_Registro")
        rscheques("fecha_Entregado") = x
        rscheques.Update
        rscheques.MoveNext
      Wend
    End If
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim i As Integer
Dim SwOrden As Variant
    
    Print ColIndex
    If rsbusca.State = 1 Then rsbusca.Close
    If SwOrden = "" Then
        Exit Sub
    End If
    If SwOrden Then
        rsbusca.Open Vquery & " order by " & CmbCampo.List(ColIndex) & " asc ", db, adOpenStatic, adLockReadOnly
        Set Adodc1.Recordset = rsbusca
        Set DataGrid1.DataSource = Adodc1
        SwOrden = False
    Else
        rsbusca.Open Vquery & " order by " & CmbCampo.List(ColIndex) & " desc ", db, adOpenStatic, adLockReadOnly
        Set Adodc1.Recordset = rsbusca
        Set DataGrid1.DataSource = Adodc1
        SwOrden = True
    End If
 
End Sub
Private Sub Form_Load()
Dim errCriterio As Variant

db.Execute "delete from to_operaciones_cheques"
If FrmActivacionCheques.OptCheques = True Then
    Vquery = "select * from cns_cheques where  cheque_o_trf= 'C'"
End If
If FrmActivacionCheques.OptTransferencias = True Then
    Vquery = "select * from cns_cheques where  cheque_o_trf= 'T'"
End If

    errCriterio = "Existe Error en el Criterio"
    If rsbusca.State = 1 Then rsbusca.Close
    On Error GoTo Error:
        rsbusca.Open Vquery, db, adOpenStatic, adLockReadOnly
        If rsbusca.RecordCount > 0 Then
            NoRegistros = rsbusca.RecordCount
        Else
            NoRegistros = 0
        End If
        
        Set Adodc1.Recordset = rsbusca
        Set DataGrid1.DataSource = Adodc1
        For Each CAMPOS In rsbusca.Fields
          If Mid(CAMPOS.Name, 1, 6) = "estado" Then
            CmbCampo.AddItem CAMPOS.Name
          End If
        Next CAMPOS

        
        DTPFechaInicio.Value = Date
        DTPFechaFin.Value = Date
    Exit Sub
Error:
    MsgBox "Existe error de sintaxis", vbDefaultButton2, "ERROR"
    MsgBox Vquery
	Call SeguridadSet(Me)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Option1_Click()
    CmbCampo.Text = "estado_impreso"
    TxtValor.Text = "S"
End Sub

Private Sub Option10_Click()
    CmbCampo.Text = "estado_devuelto"
    TxtValor.Text = "N"
End Sub

Private Sub Option2_Click()
    CmbCampo.Text = "estado_impreso"
    TxtValor.Text = "N"
End Sub

Private Sub Option3_Click()
    CmbCampo.Text = "estado_entregado"
    TxtValor.Text = "S"
End Sub

Private Sub Option4_Click()
    CmbCampo.Text = "estado_entregado"
    TxtValor.Text = "N"
End Sub

Private Sub Option5_Click()
    CmbCampo.Text = "estado_cobrado"
    TxtValor.Text = "S"
End Sub

Private Sub Option6_Click()
    CmbCampo.Text = "estado_cobrado"
    TxtValor.Text = "N"
End Sub

Private Sub Option7_Click()
    CmbCampo.Text = "estado_anulado"
    TxtValor.Text = "S"
End Sub

Private Sub Option8_Click()
    CmbCampo.Text = "estado_anulado"
    TxtValor.Text = "N"
End Sub

Private Sub Option9_Click()
    CmbCampo.Text = "estado_devuelto"
    TxtValor.Text = "S"
End Sub

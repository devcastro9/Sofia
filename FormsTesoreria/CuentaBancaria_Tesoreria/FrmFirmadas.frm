VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmFirmadas 
   Caption         =   "Cheques y / o Transferencias  Firmadas"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   15180
      TabIndex        =   2
      Top             =   0
      Width           =   15240
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "CHEQUES Y TRANSFERENCIAS FIRMADAS Y SIN FIRMAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1875
         TabIndex        =   7
         Top             =   135
         Width           =   8625
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   6
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label6 
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
         Left            =   9210
         TabIndex        =   5
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   4
         Top             =   690
         Width           =   2460
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   60
         TabIndex        =   3
         Top             =   675
         Width           =   1110
      End
   End
   Begin VB.Frame Frame2 
      Height          =   825
      Left            =   1290
      TabIndex        =   27
      Top             =   945
      Width           =   12345
      Begin VB.Frame Frame4 
         Height          =   525
         Left            =   3450
         TabIndex        =   35
         Top             =   150
         Width           =   2895
         Begin VB.OptionButton OptTodasCuentas 
            Caption         =   "X Todas las Cuentas"
            Height          =   315
            Left            =   120
            TabIndex        =   36
            Top             =   135
            Value           =   -1  'True
            Width           =   2010
         End
      End
      Begin VB.Frame Frame3 
         Height          =   540
         Left            =   315
         TabIndex        =   32
         Top             =   135
         Width           =   2880
         Begin VB.OptionButton OptFechaImpresion 
            Caption         =   "Fecha de Impresión"
            Height          =   315
            Left            =   90
            TabIndex        =   33
            Top             =   150
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin MSComCtl2.DTPicker DTPFechaInicio 
         Height          =   375
         Left            =   6750
         TabIndex        =   28
         Top             =   360
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   24707073
         CurrentDate     =   36413
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   375
         Left            =   8880
         TabIndex        =   29
         Top             =   375
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   24707073
         CurrentDate     =   36413
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha Fin"
         Height          =   240
         Left            =   8910
         TabIndex        =   31
         Top             =   180
         Width           =   1590
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha Inicio"
         Height          =   240
         Left            =   6765
         TabIndex        =   30
         Top             =   150
         Width           =   1590
      End
   End
   Begin Crystal.CrystalReport CryNoFir 
      Left            =   6195
      Top             =   4980
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
      Height          =   7035
      Left            =   15
      TabIndex        =   21
      Top             =   990
      Width           =   1245
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   120
         Picture         =   "FrmFirmadas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2895
         Width           =   945
      End
      Begin VB.CommandButton CmdBusqueda 
         Caption         =   "Busqueda"
         Height          =   855
         Left            =   120
         Picture         =   "FrmFirmadas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2040
         Width           =   945
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Impresión Firmados"
         Height          =   885
         Left            =   105
         Picture         =   "FrmFirmadas.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1155
         Width           =   960
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Impresión Sin Firmar"
         Height          =   885
         Left            =   105
         Picture         =   "FrmFirmadas.frx":0BAE
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.Frame FraBusca 
      Height          =   2115
      Left            =   1740
      TabIndex        =   12
      Top             =   3705
      Visible         =   0   'False
      Width           =   2040
      Begin VB.CommandButton CmdSalirBusqueda 
         Caption         =   "Salir"
         Height          =   390
         Left            =   225
         TabIndex        =   17
         Top             =   1650
         Width           =   1515
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   390
         Left            =   225
         TabIndex        =   16
         Top             =   1245
         Width           =   1515
      End
      Begin VB.TextBox TxtCmpte 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   225
         TabIndex        =   15
         Top             =   780
         Width           =   1515
      End
      Begin VB.TextBox TxtOrg 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2047
         TabIndex        =   14
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox TxtGes 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3615
         TabIndex        =   13
         Top             =   915
         Width           =   1515
      End
      Begin VB.Label Label21 
         Caption         =   "Cmpte. Inicial"
         Height          =   165
         Left            =   450
         TabIndex        =   20
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Organismo"
         Height          =   165
         Left            =   2310
         TabIndex        =   19
         Top             =   615
         Width           =   990
      End
      Begin VB.Label Label20 
         Caption         =   "Gestión"
         Height          =   165
         Left            =   3900
         TabIndex        =   18
         Top             =   645
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   1320
      TabIndex        =   8
      Top             =   1695
      Width           =   12315
      Begin VB.Label Label4 
         Caption         =   "No Firmados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   135
         TabIndex        =   10
         Top             =   255
         Width           =   2445
      End
      Begin VB.Label Label7 
         Caption         =   "Firmados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5955
         TabIndex        =   9
         Top             =   240
         Width           =   2445
      End
   End
   Begin VB.CommandButton Seleccionar 
      Caption         =   ">>"
      Height          =   750
      Left            =   6150
      TabIndex        =   1
      Top             =   2400
      Width           =   1020
   End
   Begin VB.CommandButton Retornar 
      Caption         =   "<<"
      Height          =   750
      Left            =   6150
      TabIndex        =   0
      Top             =   3195
      Width           =   1020
   End
   Begin Crystal.CrystalReport CryFir 
      Left            =   6180
      Top             =   4470
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
   Begin MSDataGridLib.DataGrid DtGFir 
      Height          =   5670
      Left            =   7215
      TabIndex        =   11
      Top             =   2355
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   10001
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
   Begin MSDataGridLib.DataGrid DtGNoFir 
      Height          =   5670
      Left            =   1290
      TabIndex        =   23
      Top             =   2370
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   10001
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
            Type            =   1
            Format          =   "0.00"
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
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   75
      Left            =   10875
      TabIndex        =   24
      Top             =   1725
      Width           =   45
   End
End
Attribute VB_Name = "FrmFirmadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsNoFirmadas As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset

Public Sub Reporte_TodasCtas()
Dim Fecha1 As New ADODB.Parameter
Dim Fecha2 As New ADODB.Parameter
Dim op1 As New ADODB.Parameter
    
Dim De As String
Dim A As String
Dim TC As String
Dim Cadena As String
If FrmFirmadas.OptFechaImpresion.Value = True Then
     Cadena = "REPORTE POR FECHA DE IMPRESION"
End If


If FrmFirmadas.OptFechaImpresion.Value = True Then CryNoFir.Formulas(4) = "Fecha_Pago_Impresion='" & Cadena & "'"

If FrmFirmadas.OptTodasCuentas.Value = True Then
    TC = "Todas las cuentas"
    CryNoFir.Formulas(9) = "FTodasCuentas='" & TC & "'"
End If

    De = "De"
    A = "A"
    CryNoFir.Formulas(6) = "FFechaInicio='" & FrmListadoPagos.DTPFechaInicio.Value & "'"
    CryNoFir.Formulas(5) = "FFechaFin='" & FrmListadoPagos.DTPFechaFin.Value & "'"
    CryNoFir.Formulas(2) = "FDe='" & De & "' "
    CryNoFir.Formulas(0) = "Fa='" & A & "' "
    
    If OptFechaImpresion.Value = True Then
        op1 = "I"
    End If
    Set comPagos = New ADODB.Command
    With comPagos
        .CommandText = "Cel_Listado_Pagos_Todos"
        .CommandType = adCmdStoredProc
        Set Fecha1 = .CreateParameter("FechaIni", adVarChar, adParamInput, 10, DTPFechaInicio.Value)
        .Parameters.Append Fecha1
        Set Fecha2 = .CreateParameter("FechaFin", adVarChar, adParamInput, 10, DTPFechaFin.Value)
        .Parameters.Append Fecha2
        Set op1 = .CreateParameter("Opcion", adVarChar, adParamInput, 1, op1)
        .Parameters.Append op1
        
        .ActiveConnection = db
        .Execute
    End With
    

End Sub

Private Sub CmdBusqueda_Click()
    FraBusca.Visible = True
End Sub

Private Sub CmdImprimir_Click()
        CryFir.ReportFileName = "C:\SAF-2000\FormsTesoreria\CuentaBancaria_Tesoreria\Impresiones\Rpt_CtaBancaria.rpt"
        iResult = CryFir.PrintReport
        If iResult <> 0 Then
           MsgBox CryFir.LastErrorNumber & " : " & CryFir.LastErrorString, vbCritical + vbOKOnly, "Error..."
         End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdSalirBusqueda_Click()
    FraBusca.Visible = False
End Sub

Private Sub Form_Load()

'Validación de fechas
    If DTPFechaInicio.Value > DTPFechaFin.Value Or DTPFechaFin.Value < DTPFechaInicio.Value Then
        MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
        Exit Sub
     End If
      
     DTPFechaInicio.Value = Date
     DTPFechaFin.Value = Date
    
    If OptTodasCuentas.Value = True Then
        Reporte_TodasCtas
        Set rsNoFirmadas = New ADODB.Recordset
        rsNoFirmadas.Open "select codigo_pago,org_codigo, monto_bolivianos, numero_cheque_trf, chq_trf, denominacion_beneficiario   from to_Movimiento", db, adOpenKeyset, adLockOptimistic
        If rsNoFirmadas.RecordCount > 0 Then
            Set DtGNoFir.DataSource = rsNoFirmadas
        Else
            Set DtGNoFir.DataSource = rsNada
        End If
    End If
    
    
	Call SeguridadSet(Me)
End Sub

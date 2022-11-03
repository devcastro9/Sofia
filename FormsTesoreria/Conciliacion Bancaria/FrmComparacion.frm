VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmComparacion 
   Caption         =   "Comparación de Datos de GTZ y Banco"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   Icon            =   "FrmComparacion.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CryConBan 
      Left            =   5460
      Top             =   8910
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
   Begin Crystal.CrystalReport CryConGTZ 
      Left            =   4995
      Top             =   8910
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
   Begin VB.Frame Frame5 
      Height          =   480
      Left            =   7815
      TabIndex        =   20
      Top             =   1905
      Width           =   6435
      Begin VB.Label Label2 
         Caption         =   "G  T  Z     -     U  D  A  P  R  E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   345
         TabIndex        =   21
         Top             =   165
         Width           =   2670
      End
   End
   Begin VB.Frame Frame4 
      Height          =   480
      Left            =   1365
      TabIndex        =   18
      Top             =   1905
      Width           =   6405
      Begin VB.Label LblBanco 
         Caption         =   "BANCO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   23
         Top             =   150
         Width           =   5025
      End
      Begin VB.Label Label1 
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   19
         Top             =   180
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      Height          =   870
      Left            =   1380
      TabIndex        =   11
      Top             =   1050
      Width           =   12840
      Begin VB.TextBox TxtCuentaBancaria 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9495
         TabIndex        =   24
         Top             =   225
         Width           =   2250
      End
      Begin VB.TextBox TxtCompFFin 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5010
         TabIndex        =   17
         Top             =   195
         Width           =   2145
      End
      Begin VB.TextBox TxtCompFIni 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   195
         Width           =   2145
      End
      Begin MSComCtl2.Animation AVI 
         Height          =   630
         Left            =   11865
         TabIndex        =   26
         Top             =   150
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1111
         _Version        =   393216
         FullWidth       =   59
         FullHeight      =   42
      End
      Begin VB.Label Label10 
         Caption         =   "Cuenta Bancaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7860
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Fin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3945
         TabIndex        =   16
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Inicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   14
         Top             =   225
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   7215
      TabIndex        =   3
      Top             =   0
      Width           =   7275
      Begin VB.Label LblTitulo 
         BackColor       =   &H8000000A&
         Caption         =   "CONCILIACION BANCARIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   450
         Left            =   4605
         TabIndex        =   12
         Top             =   135
         Width           =   8850
      End
      Begin VB.Label Label8 
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
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   675
         Width           =   1110
      End
      Begin VB.Label Label7 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Index           =   0
         Left            =   1245
         TabIndex        =   6
         Top             =   690
         Width           =   2460
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   9210
         TabIndex        =   5
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   4
         Top             =   660
         Width           =   1305
      End
   End
   Begin VB.Frame FraOpciones 
      Height          =   7875
      Left            =   15
      TabIndex        =   8
      Top             =   1020
      Width           =   1320
      Begin VB.CommandButton CmdImprimirGTZ 
         Caption         =   "Imprimir Conciliados GTZ"
         Height          =   1155
         Left            =   105
         Picture         =   "FrmComparacion.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1410
         Width           =   1095
      End
      Begin VB.CommandButton CmdImprimirBanco 
         Caption         =   "Imprimir Conciliados Banco"
         Height          =   1080
         Left            =   105
         Picture         =   "FrmComparacion.frx":1534
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   315
         Width           =   1095
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "Coloca status de conciliacion en N"
         Height          =   1155
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   945
         Left            =   90
         Picture         =   "FrmComparacion.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5460
         Width           =   1110
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6600
      Left            =   7830
      TabIndex        =   2
      Top             =   2280
      Width           =   6420
      Begin MSDataGridLib.DataGrid DtGConciliacionUDAPRE 
         Height          =   6105
         Left            =   150
         TabIndex        =   27
         Top             =   255
         Width           =   6060
         _ExtentX        =   10689
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
      Begin MSAdodcLib.Adodc AdoConciliacionUDAPRE 
         Height          =   435
         Left            =   165
         Top             =   6015
         Visible         =   0   'False
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   767
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
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   1350
      TabIndex        =   0
      Top             =   2280
      Width           =   6420
      Begin MSDataGridLib.DataGrid DtGConciliacionBanco 
         Height          =   6120
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   10795
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
               LCID            =   3082
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
               LCID            =   3082
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
   End
End
Attribute VB_Name = "FrmComparacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsGTZ As New ADODB.Recordset
Dim rsBANCO As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset

Private Sub CmdImprimirBanco_Click()
Dim Titulo As String
        If FrmConciliacion.OptConciliados.Value = True Then
            Titulo = "REGISTROS CONCILIADOS - BANCO"
        End If
        If FrmConciliacion.optNoConciliados.Value = True Then
            Titulo = "REGISTROS NO CONCILIADOS - BANCO"
        End If

        CryConBan.Formulas(0) = "FechaIni='" & FrmComparacion.TxtCompFIni.Text & "'"
        CryConBan.Formulas(1) = "FechaFin='" & FrmComparacion.TxtCompFFin.Text & "'"
        CryConBan.Formulas(2) = "FTitulo='" & Titulo & "'"
        
        CryConBan.ReportFileName = App.Path & "\FormsTesoreria\Conciliacion Bancaria\REPORTES\Rep_ConciliadosBANCO.rpt"
        iResult = CryConBan.PrintReport
        If iResult <> 0 Then
           MsgBox CryConBan.LastErrorNumber & " : " & CryConGTZ.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
End Sub

Private Sub CmdImprimirGTZ_Click()
        
Dim Titulo As String
        If FrmConciliacion.OptConciliados.Value = True Then
            Titulo = "REGISTROS CONCILIADOS - "
        End If
        If FrmConciliacion.optNoConciliados.Value = True Then
            Titulo = "REGISTROS NO CONCILIADOS - "
        End If

        CryConGTZ.Formulas(0) = "FechaIni='" & FrmComparacion.TxtCompFIni.Text & "'"
        CryConGTZ.Formulas(1) = "FechaFin='" & FrmComparacion.TxtCompFFin.Text & "'"
        CryConGTZ.Formulas(2) = "FTitulo='" & Titulo & "'"
        
        
        CryConGTZ.ReportFileName = App.Path & "\FormsTesoreria\Conciliacion Bancaria\REPORTES\Rep_ConciliadosGTZ.rpt"
        iResult = CryConGTZ.PrintReport
        If iResult <> 0 Then
           MsgBox CryConGTZ.LastErrorNumber & " : " & CryConGTZ.LastErrorString, vbCritical + vbOKOnly, "Error..."
         End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
If FrmConciliacion.OptConciliados.Value = True Then
         Conciliados_Proc
End If
If FrmConciliacion.optNoConciliados.Value = True Then
         NoConciliados_Proc
End If

    Set rsGTZ = New ADODB.Recordset
    rsGTZ.Open "SELECT * FROM to_DatosGTZ ", db, adOpenKeyset, adLockOptimistic
    If rsGTZ.RecordCount > 0 Then
        Set DtGConciliacionUDAPRE.DataSource = rsGTZ
    Else
        Set DtGConciliacionUDAPRE.DataSource = rsNada
    End If
    
    Set rsGTZ = New ADODB.Recordset
    rsGTZ.Open "SELECT * FROM to_DatosBanco", db, adOpenKeyset, adLockOptimistic
    If rsGTZ.RecordCount > 0 Then
        Set DtGConciliacionBanco.DataSource = rsGTZ
    Else
        Set DtGConciliacionBanco.DataSource = rsNada
    End If
	Call SeguridadSet(Me)
End Sub

Public Sub Conciliados_Proc()



If FrmConciliacion.OptEgresos.Value = True Then op1 = "EGR"
If FrmConciliacion.OptIngresos.Value = True Then op1 = "ING"
If FrmConciliacion.OptTraspasos.Value = True Then op1 = "TRP"
If FrmConciliacion.opttodos.Value = True Then op1 = "TDS"

  Set comConciliados = New ADODB.Command
  With comConciliados
        .CommandText = "Cel_Conciliacion_FechaGTZ"
        .CommandType = adCmdStoredProc
        Set fecha1 = .CreateParameter("FechaIni", adVarChar, adParamInput, 10, FrmConciliacion.DTPInicio.Value)
        .Parameters.Append fecha1
        Set fecha2 = .CreateParameter("FechaFin", adVarChar, adParamInput, 10, FrmConciliacion.DTPFin.Value)
        .Parameters.Append fecha2
        Set op1 = .CreateParameter("Opcion", adVarChar, adParamInput, 3, op1)
        .Parameters.Append op1
        Set Cta = .CreateParameter("Cuenta", adVarChar, adParamInput, 40, FrmConciliacion.DtCCuentaOrigen.Text)
        .Parameters.Append Cta
        .ActiveConnection = db
        .Execute
    End With

End Sub

Public Sub NoConciliados_Proc()


If FrmConciliacion.OptEgresos.Value = True Then op1 = "EGR"
If FrmConciliacion.OptIngresos.Value = True Then op1 = "ING"
If FrmConciliacion.OptTraspasos.Value = True Then op1 = "TRP"
If FrmConciliacion.opttodos.Value = True Then op1 = "TDS"

  Set comConciliados = New ADODB.Command
  With comConciliados
        .CommandText = "Cel_NoConciliacion_FechaGTZ"
        .CommandType = adCmdStoredProc
        Set fecha1 = .CreateParameter("FechaIni", adVarChar, adParamInput, 10, FrmConciliacion.DTPInicio.Value)
        .Parameters.Append fecha1
        Set fecha2 = .CreateParameter("FechaFin", adVarChar, adParamInput, 10, FrmConciliacion.DTPFin.Value)
        .Parameters.Append fecha2
        Set op1 = .CreateParameter("Opcion", adVarChar, adParamInput, 3, op1)
        .Parameters.Append op1
        Set Cta = .CreateParameter("Cuenta", adVarChar, adParamInput, 40, FrmConciliacion.DtCCuentaOrigen.Text)
        .Parameters.Append Cta
        .ActiveConnection = db
        .Execute
    End With

End Sub

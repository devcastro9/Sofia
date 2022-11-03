VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmActivacionCheques 
   Caption         =   "Activación de Cheques Impresos"
   ClientHeight    =   8595
   ClientLeft      =   885
   ClientTop       =   930
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ListBox LstCuenta 
      Height          =   5325
      Left            =   10590
      TabIndex        =   50
      Top             =   1545
      Width           =   1155
   End
   Begin VB.Frame FraBusqueda 
      Height          =   1845
      Left            =   1680
      TabIndex        =   37
      Top             =   4200
      Visible         =   0   'False
      Width           =   6060
      Begin VB.Frame Frame1 
         Height          =   1065
         Left            =   135
         TabIndex        =   41
         Top             =   150
         Width           =   5820
         Begin VB.TextBox TxtValor 
            Height          =   285
            Left            =   3165
            TabIndex        =   44
            Top             =   645
            Width           =   2505
         End
         Begin VB.ComboBox CmbOperador 
            Height          =   315
            ItemData        =   "FrmActivacionCheques.frx":0000
            Left            =   1965
            List            =   "FrmActivacionCheques.frx":0013
            TabIndex        =   43
            Top             =   630
            Width           =   1065
         End
         Begin VB.ComboBox CmbCampo 
            Height          =   315
            Left            =   45
            TabIndex        =   42
            Top             =   630
            Width           =   1815
         End
         Begin VB.Label LblValor 
            Caption         =   "Valor"
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
            Left            =   3315
            TabIndex        =   47
            Top             =   255
            Width           =   675
         End
         Begin VB.Label LblOperador 
            Caption         =   "Operador"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1965
            TabIndex        =   46
            Top             =   255
            Width           =   885
         End
         Begin VB.Label LblCampo 
            Caption         =   "Campo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   45
            TabIndex        =   45
            Top             =   255
            Width           =   615
         End
      End
      Begin VB.CommandButton CmdImprimirBusqueda 
         Caption         =   "Imprimir"
         Height          =   390
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1320
         Width           =   1140
      End
      Begin VB.CommandButton CmdCancelarBusqueda 
         Caption         =   "Cancelar"
         Height          =   390
         Left            =   2370
         TabIndex        =   39
         Top             =   1320
         Width           =   1140
      End
      Begin VB.CommandButton CmdEjecutarBusqueda 
         Caption         =   "Ejecutar"
         Height          =   390
         Left            =   1245
         TabIndex        =   38
         Top             =   1320
         Width           =   1140
      End
   End
   Begin VB.TextBox TxtCheques 
      Appearance      =   0  'Flat
      Height          =   465
      Left            =   1305
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   7950
      Width           =   9465
   End
   Begin VB.OptionButton OptTransferencias 
      Caption         =   "Transferencias"
      Height          =   255
      Left            =   3300
      TabIndex        =   28
      Top             =   1230
      Width           =   1500
   End
   Begin VB.OptionButton OptCheques 
      Caption         =   "Cheques"
      Height          =   255
      Left            =   1335
      TabIndex        =   27
      Top             =   1245
      Value           =   -1  'True
      Width           =   1785
   End
   Begin VB.Frame FraBusca 
      Height          =   1335
      Left            =   5730
      TabIndex        =   22
      Top             =   4605
      Visible         =   0   'False
      Width           =   3750
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   390
         Left            =   2445
         TabIndex        =   29
         Top             =   855
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   390
         Left            =   1275
         TabIndex        =   26
         Top             =   855
         Width           =   1170
      End
      Begin VB.CommandButton CmdEjecutar 
         Caption         =   "Ejecutar"
         Height          =   390
         Left            =   165
         TabIndex        =   24
         Top             =   855
         Width           =   1110
      End
      Begin VB.TextBox TxtBusca 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   165
         TabIndex        =   23
         Top             =   480
         Width           =   3420
      End
      Begin VB.Label Label22 
         Caption         =   "Buscar"
         Height          =   180
         Left            =   180
         TabIndex        =   25
         Top             =   210
         Width           =   525
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   930
      Left            =   0
      Picture         =   "FrmActivacionCheques.frx":002A
      ScaleHeight     =   870
      ScaleWidth      =   11820
      TabIndex        =   16
      Top             =   0
      Width           =   11880
      Begin VB.Label Label8 
         Caption         =   "UNIDAD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   60
         TabIndex        =   21
         Top             =   495
         Width           =   1110
      End
      Begin VB.Label Label7 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   20
         Top             =   525
         Width           =   2460
      End
      Begin VB.Label Label6 
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9210
         TabIndex        =   19
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   18
         Top             =   495
         Width           =   1305
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "OPERACION CHEQUES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   4350
         TabIndex        =   17
         Top             =   135
         Width           =   3615
      End
   End
   Begin VB.TextBox TxtFinal 
      Height          =   390
      Left            =   13350
      TabIndex        =   12
      Top             =   2985
      Width           =   1380
   End
   Begin VB.TextBox TxtInicial 
      Height          =   390
      Left            =   13380
      TabIndex        =   10
      Top             =   1950
      Width           =   1380
   End
   Begin VB.ListBox LstCheques 
      Height          =   5325
      Left            =   9465
      TabIndex        =   7
      Top             =   1545
      Width           =   1110
   End
   Begin MSAdodcLib.Adodc AdoPagos 
      Height          =   420
      Left            =   1380
      Top             =   10350
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   741
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
      Caption         =   "Cheques"
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
   Begin VB.Frame FraOpciones 
      Height          =   7500
      Left            =   15
      TabIndex        =   1
      Top             =   915
      Width           =   1245
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   150
         TabIndex        =   49
         Top             =   3465
         Width           =   945
      End
      Begin VB.CommandButton CmdRestaurar 
         Caption         =   "Restaurar"
         Height          =   510
         Left            =   150
         TabIndex        =   48
         Top             =   3960
         Width           =   945
      End
      Begin VB.CommandButton CmdCobrado 
         Caption         =   "Cobrado"
         Height          =   480
         Left            =   165
         TabIndex        =   15
         Top             =   2505
         Width           =   930
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   480
         Left            =   165
         TabIndex        =   14
         Top             =   2985
         Width           =   930
      End
      Begin VB.CommandButton CmdAnulado 
         Caption         =   "Anulado"
         Height          =   495
         Left            =   165
         TabIndex        =   6
         Top             =   2010
         Width           =   930
      End
      Begin VB.CommandButton CmdDevuelto 
         Caption         =   "Devuelto"
         Height          =   495
         Left            =   165
         TabIndex        =   5
         Top             =   1515
         Width           =   930
      End
      Begin VB.CommandButton CmdEntregado 
         Caption         =   "Entregado"
         Height          =   510
         Left            =   180
         TabIndex        =   4
         Top             =   1005
         Width           =   915
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   150
         Picture         =   "FrmActivacionCheques.frx":1BD0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5145
         Width           =   945
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprime Cheque"
         Height          =   795
         Left            =   180
         Picture         =   "FrmActivacionCheques.frx":2012
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   915
      End
   End
   Begin MSDataGridLib.DataGrid DtGPagos 
      Height          =   5340
      Left            =   1350
      TabIndex        =   0
      Top             =   1575
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   9419
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
   Begin MSDataListLib.DataCombo DtCCuentaOrigen 
      Bindings        =   "FrmActivacionCheques.frx":267C
      DataField       =   "cta_codigo"
      DataSource      =   "AdoCuenta"
      Height          =   315
      Left            =   1350
      TabIndex        =   33
      Top             =   7245
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "cta_codigo"
      BoundColumn     =   "cta_codigo"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DtCCuentaOrigenDes 
      Bindings        =   "FrmActivacionCheques.frx":2694
      DataField       =   "cta_codigo"
      DataSource      =   "AdoCuenta"
      Height          =   315
      Left            =   5115
      TabIndex        =   34
      Top             =   7245
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Cta_descripcion_larga"
      BoundColumn     =   "cta_codigo"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DtcCtaTGN 
      Bindings        =   "FrmActivacionCheques.frx":26AC
      DataField       =   "cta_codigo"
      DataSource      =   "AdoCuenta"
      Height          =   315
      Left            =   3495
      TabIndex        =   35
      Top             =   7245
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Cta_codigo_tgn"
      BoundColumn     =   "cta_codigo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc AdoCuenta 
      Height          =   390
      Left            =   9480
      Top             =   7200
      Visible         =   0   'False
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   688
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
      Caption         =   "Cuenta Bancaria"
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
   Begin MSComCtl2.DTPicker DTPFechaRegistro 
      Height          =   345
      Left            =   7530
      TabIndex        =   51
      Top             =   1110
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   609
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   64356353
      CurrentDate     =   36413
   End
   Begin VB.Label Label12 
      Caption         =   "Fecha de Registro"
      Height          =   240
      Left            =   6075
      TabIndex        =   54
      Top             =   1230
      Width           =   1320
   End
   Begin VB.Label Label11 
      Caption         =   "Cuenta"
      Height          =   270
      Left            =   10560
      TabIndex        =   53
      Top             =   1230
      Width           =   1125
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha Inicio"
      Height          =   240
      Left            =   30
      TabIndex        =   52
      Top             =   0
      Width           =   1590
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "No. Cta. "
      Height          =   195
      Left            =   1350
      TabIndex        =   36
      Top             =   7050
      Width           =   630
   End
   Begin VB.Label Label9 
      Caption         =   "SELECCIONE EL CHEQUE A OPERARSE"
      Height          =   195
      Left            =   1350
      TabIndex        =   32
      Top             =   1005
      Width           =   2085
   End
   Begin VB.Label Label10 
      Caption         =   "DIGITE EL NUMERO DE CHEQUE A OPERARSE (ESPECIFICOS  00122,00345, etc.)  (RANGOS 00122-00129)"
      Height          =   240
      Left            =   1350
      TabIndex        =   31
      Top             =   7650
      Width           =   8265
   End
   Begin VB.Label Label4 
      Caption         =   "Cheques Consecutivos"
      Height          =   360
      Left            =   13350
      TabIndex        =   13
      Top             =   1215
      Width           =   1800
   End
   Begin VB.Label Label3 
      Caption         =   "Nro. Cheque Final"
      Height          =   450
      Left            =   13320
      TabIndex        =   11
      Top             =   2685
      Width           =   1785
   End
   Begin VB.Label Label2 
      Caption         =   "Nro. Cheque Inicial"
      Height          =   450
      Left            =   13365
      TabIndex        =   9
      Top             =   1665
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Cheques"
      Height          =   270
      Left            =   9450
      TabIndex        =   8
      Top             =   1245
      Width           =   1125
   End
End
Attribute VB_Name = "FrmActivacionCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' Sistema:                  SAF-2000
' Módulo:                   Operaciones sobre cheques y transferencias
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmActivaciónCheques
' Descipción :              Control de los status de Cheq/Trans de entrgado, pagado, anulado,cobrado
' Formularios relacionados: Main.frm (Padre)
'                           CryopCheques
' Autor:                    Celia Elena Tarquino Peralta
' Fecha de creación         14/Abril/ 2000
' Fecha última modificación 01/May/ 2000
' Versión:                  2.0
'========================================================================================

Dim rscorrelativo As New ADODB.Recordset
Dim rsComprobante As New ADODB.Recordset
Dim rsCheques As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
Dim NrosChequeImprimir As String

Private Sub CmdAnulado_Click()
Dim i As Integer
Dim Cheque_Inicial As Long
Dim Cheque_Final As Long
Dim s As String
Dim k As Long

Dim sino As Variant

s = ""

sino = MsgBox("Está seguro de colocar este status de anulado?", vbYesNo + vbQuestion, "Atenciòn")
If sino = vbYes Then

        If TxtCheques.Text <> "" Then
          Anulado_Lista
          Exit Sub
'        Else
'          MsgBox "No existen Nros. de cheque", vbInformation + vbCritical
'          Exit Sub
        End If

        If TxtInicial.Text <> "" Then
                Cheque_Inicial = Val(TxtInicial.Text)
                Cheque_Final = Val(TxtFinal.Text)
                For k = Cheque_Inicial To Cheque_Final Step 1
                    s = ""
                    T = CStr(k)
                    Select Case Len(T)
                           Case 1
                                s = "0000" + CStr(k)
                           Case 2
                                s = "000" + CStr(k)
                           Case 3
                                s = "00" + CStr(k)
                           Case 4
                                s = "0" + CStr(k)
                           Case 5
                                s = CStr(k)
                    End Select
                    
                     LstCheques.AddItem s
                Next k
                
                For i = 0 To LstCheques.ListCount - 1
                    LstCheques.ListIndex = i
                    If rsCheques.State = 1 Then rsCheques.Close
                    rsCheques.Open "SELECT * FROM to_cheques_operaciones where numero_cheque='" & LstCheques.Text & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
                    If rsCheques.RecordCount > 0 Then
                            rsCheques("estado_anulado") = "S"
                    Else
                            rsCheques.AddNew
                            rsCheques("numero_cheque") = LstCheques.Text
                            rsCheques("estado_anulado") = "S"
                    End If
                    rsCheques("usr_usuario") = LblUsuario.Caption
                    rsCheques("fecha_registro") = DTPFechaRegistro.Value 'Date
                    rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
                    rsCheques.Update
                Next i
        End If
        
        LstCheques.ListIndex = 0
        If LstCheques.Text <> "" Then
           If rsCheques.State = 1 Then rsCheques.Close
           For i = 0 To LstCheques.ListCount - 1
                    LstCheques.ListIndex = i
                    LstCuenta.ListIndex = i
                    If DtGPagos.Columns(0) = "" Then
                        MsgBox "No existre cheque ", vbInformation + vbCritical, "Validación de datos"
                        Exit Sub
                    End If
                    If rsCheques.State = 1 Then rsCheques.Close
                    rsCheques.Open "SELECT * FROM to_cheques_operaciones WHERE  numero_cheque= '" & LstCheques.Text & "' and cta_codigo= '" & LstCuenta.Text & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
                    If rsCheques.RecordCount > 0 Then
                        rsCheques("estado_anulado") = "S"
                    Else
                        rsCheques.AddNew
                        'rsCheques("numero_cheque") = AdoPagos.Recordset("numero_cheque_trf")
                        rsCheques("numero_cheque") = LstCheques.Text
                        rsCheques("cta_codigo") = LstCuenta.Text
                        rsCheques("estado_anulado") = "S"
                    End If
                    rsCheques("usr_usuario") = LblUsuario.Caption
                    rsCheques("fecha_registro") = DTPFechaRegistro.Value 'Date
                    rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
                    rsCheques.Update
            Next i
        End If
        
        
        '    If DtGPagos.Columns(0) = "" Then
        '        MsgBox "No existre cheque ", vbInformation + vbCritical, "Validación de datos"
        '        Exit Sub
        '    End If
        '    MsgBox "Cheque " + AdoPagos.Recordset("numero_cheque_trf") + " Devuelto"
        '    Set rsCheques = New ADODB.Recordset
        '    If rsCheques.State = 1 Then rsCheque.Close
        '    rsCheques.Open "SELECT * FROM to_cheques WHERE numero_cheque = '" & AdoPagos.Recordset("numero_cheque_trf") & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
        '    If rsCheques.RecordCount > 0 Then
        '        rsCheques("estado_anulado") = "S"
        '    Else
        '        rsCheques.AddNew
        '        rsCheques("numero_cheque") = AdoPagos.Recordset("numero_cheque_trf")
        '        rsCheques("estado_anulado") = "S"
        '    End If
        '        rsCheques.Update
End If
End Sub

Private Sub CmdBuscar_Click()
''    FraBusca.Visible = True
' GrFrmFiltraBusca.GrPrincipal rsComprobante, QueryInicial, DtGPagos
FraBusqueda.Visible = True
On Error GoTo Error:
        For Each CAMPOS In rsComprobante.Fields
            CmbCAMPO.AddItem CAMPOS.Name
        Next CAMPOS
        FraBusqueda.Visible = True
Exit Sub
Error:
    MsgBox "Existe error de sintaxis", vbDefaultButton2, "ERROR"
End Sub

Private Sub CmdCancelar_Click()
    FraBusca.Visible = False
End Sub

Private Sub CmdCancelarBusqueda_Click()
    FraBusqueda.Visible = False
End Sub

Private Sub CmdCobrado_Click()
Dim i As Integer
Dim Cheque_Inicial As Long
Dim Cheque_Final As Long
Dim s As String
Dim k As Long
Dim sino As Variant

s = ""

sino = MsgBox("Está seguro de colocar este status de cobrado?", vbYesNo + vbQuestion, "Atenciòn")
If sino = vbYes Then

       If TxtCheques.Text <> "" Then
          Cobrado_Lista
          Exit Sub
'        Else
'          MsgBox "No existen Nros. de cheque", vbInformation + vbCritical
'          Exit Sub
        End If

        LstCheques.ListIndex = 0
        If LstCheques.Text <> "" Then
           If rsCheques.State = 1 Then rsCheques.Close
           For i = 0 To LstCheques.ListCount - 1
                    LstCheques.ListIndex = i
                    If DtGPagos.Columns(0) = "" Then
                        MsgBox "No existre cheque ", vbInformation + vbCritical, "Validación de datos"
                        Exit Sub
                    End If
                    If rsCheques.State = 1 Then rsCheques.Close
                    rsCheques.Open "SELECT * FROM to_cheques_operaciones WHERE  numero_cheque= '" & LstCheques.Text & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
                    If rsCheques.RecordCount > 0 Then
                        rsCheques("estado_cobrado") = "S"
                    Else
                        rsCheques.AddNew
                        'rsCheques("numero_cheque") = AdoPagos.Recordset("numero_cheque_trf")
                        rsCheques("numero_cheque") = LstCheques.Text
                        rsCheques("estado_cobrado") = "S"
                    End If
                    rsCheques("usr_usuario") = LblUsuario.Caption
                    rsCheques("fecha_registro") = DTPFechaRegistro.Value 'Date
                    rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
                    rsCheques.Update
            Next i
        End If
        
        s = ""
        If TxtInicial.Text <> "" Then
                Cheque_Inicial = Val(TxtInicial.Text)
                Cheque_Final = Val(TxtFinal.Text)
                For k = Cheque_Inicial To Cheque_Final Step 1
                    s = ""
                    T = CStr(k)
                    Select Case Len(T)
                           Case 1
                                s = "0000" + CStr(k)
                           Case 2
                                s = "000" + CStr(k)
                           Case 3
                                s = "00" + CStr(k)
                           Case 4
                                s = "0" + CStr(k)
                           Case 5
                                s = CStr(k)
                    End Select
                    
                     LstCheques.AddItem s
                Next k
                
                For i = 0 To LstCheques.ListCount - 1
                    LstCheques.ListIndex = i
                    LstCuenta.ListIndex = i
                    If rsCheques.State = 1 Then rsCheques.Close
                    rsCheques.Open "SELECT * FROM to_cheques_operaciones where numero_cheque='" & LstCheques.Text & "'order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
                    If rsCheques.RecordCount > 0 Then
                            rsCheques("estado_cobrado") = "S"
                    Else
                            rsCheques.AddNew
                            rsCheques("numero_cheque") = LstCheques.Text
                            rsCheques("cta_codigo") = LstCuenta.Text
                            rsCheques("estado_cobrado") = "S"
                    End If
                    rsCheques("usr_usuario") = LblUsuario.Caption
                    rsCheques("fecha_registro") = DTPFechaRegistro.Value 'Date
                    rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
                    rsCheques.Update
                Next i
        End If
End If
End Sub

Private Sub CmdDevuelto_Click()
Dim i As Integer
Dim Cheque_Inicial As Long
Dim Cheque_Final As Long
Dim s As String
Dim k As Long
Dim sino As Variant

s = ""
sino = MsgBox("Está seguro de colocar este status de devuelto?", vbYesNo + vbQuestion, "Atenciòn")
If sino = vbYes Then
      If TxtCheques.Text <> "" Then
        Devuelto_Lista
        Exit Sub
'      Else
'        MsgBox "No existen Nros. de cheque", vbInformation + vbCritical
'        Exit Sub
      End If
        LstCheques.ListIndex = 0
        If LstCheques.Text <> "" Then
           If rsCheques.State = 1 Then rsCheques.Close
           For i = 0 To LstCheques.ListCount - 1
                    LstCheques.ListIndex = i
                    If DtGPagos.Columns(0) = "" Then
                        MsgBox "No existre cheque ", vbInformation + vbCritical, "Validación de datos"
                        Exit Sub
                    End If
                    If rsCheques.State = 1 Then rsCheques.Close
                    rsCheques.Open "SELECT * FROM to_cheques_operaciones WHERE  numero_cheque= '" & LstCheques.Text & "'   order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
                    If rsCheques.RecordCount > 0 Then
                        rsCheques("estado_devuelto") = "S"
                    Else
                        rsCheques.AddNew
                        'rsCheques("numero_cheque") = AdoPagos.Recordset("numero_cheque_trf")
                        rsCheques("numero_cheque") = LstCheques.Text
                        rsCheques("estado_devuelto") = "S"
                    End If
                    rsCheques("usr_usuario") = LblUsuario.Caption
                    rsCheques("fecha_registro") = DTPFechaRegistro.Value 'Date
                    rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
                    
                    rsCheques.Update
            Next i
        End If
        
        s = ""
        If TxtInicial.Text <> "" Then
                Cheque_Inicial = Val(TxtInicial.Text)
                Cheque_Final = Val(TxtFinal.Text)
                For k = Cheque_Inicial To Cheque_Final Step 1
                    s = ""
                    T = CStr(k)
                    Select Case Len(T)
                           Case 1
                                s = "0000" + CStr(k)
                           Case 2
                                s = "000" + CStr(k)
                           Case 3
                                s = "00" + CStr(k)
                           Case 4
                                s = "0" + CStr(k)
                           Case 5
                                s = CStr(k)
                    End Select
                    
                     LstCheques.AddItem s
                Next k
                
                For i = 0 To LstCheques.ListCount - 1
                    LstCheques.ListIndex = i
                    LstCuenta.ListIndex = i
                    If rsCheques.State = 1 Then rsCheques.Close
                    rsCheques.Open "SELECT * FROM to_cheques_operaciones where numero_cheque='" & LstCheques.Text & "' and cta_codigo= '" & LstCuenta.Text & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
                    If rsCheques.RecordCount > 0 Then
                            rsCheques("estado_devuelto") = "S"
                    Else
                            rsCheques.AddNew
                            rsCheques("numero_cheque") = LstCheques.Text
                            rsCheques("cta_codigo") = LstCuenta.Text
                            rsCheques("estado_devuelto") = "S"
                    End If
                    
                    rsCheques("usr_usuario") = LblUsuario.Caption
                    rsCheques("fecha_registro") = DTPFechaRegistro.Value 'Date
                    rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
                    
                    rsCheques.Update
                Next i
        End If
        
        '    If DtGPagos.Columns(0) = "" Then
        '        MsgBox "No existre cheque ", vbInformation + vbCritical, "Validación de datos"
        '        Exit Sub
        '    End If
        '    MsgBox "Cheque " + AdoPagos.Recordset("numero_cheque_trf") + " Devuelto"
        '    Set rsCheques = New ADODB.Recordset
        '    If rsCheques.State = 1 Then rsCheque.Close
        '    rsCheques.Open "SELECT * FROM to_cheques WHERE numero_cheque = '" & AdoPagos.Recordset("numero_cheque_trf") & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
        '    If rsCheques.RecordCount > 0 Then
        '        rsCheques("estado_devuelto") = "S"
        '    Else
        '        rsCheques.AddNew
        '        rsCheques("numero_cheque") = AdoPagos.Recordset("numero_cheque_trf")
        '        rsCheques("estado_devuelto") = "S"
        '    End If
        '        rsCheques.Update
End If
End Sub

Private Sub CmdEjecutar_Click()
    Set rsComprobante = New ADODB.Recordset
    rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo", db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        Set DtGPagos.DataSource = rsComprobante
        Set AdoPagos.Recordset = rsComprobante
    End If
End Sub

Private Sub CmdEjecutarBusqueda_Click()
Dim cadena_busqueda As String
   cadena_busqueda = ""
    If CmbCAMPO = "codigo_pago" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + CmbOPERADOR + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "org_codigo" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + CmbOPERADOR + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "denominacion_beneficiario" Then
        cadena_busqueda = "fc_beneficiario." + CmbCAMPO.Text + " like " + "'%" + TxtValor + "%'"
    End If
    If CmbCAMPO = "fecha_pago" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " = " + "#" + TxtValor + "#"
    End If
    If CmbCAMPO = "monto_bolivianos" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "tipo_cambio" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "codigo_beneficiario" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "justificacion" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'%" + TxtValor + "%'"
    End If
    If CmbCAMPO = "cheque_o_trf" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "numero_cheque_trf" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "cta_codigo" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "Bco_descripcion_larga" Then
        cadena_busqueda = "fc_bancos." + CmbCAMPO.Text + " like " + "'%" + TxtValor + "%'"
    End If
    If CmbCAMPO = "literal" Then
        cadena_busqueda = "pago_detalle." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "cta_descripcion_larga" Then
        cadena_busqueda = "fc_cuenta_bancaria." + CmbCAMPO.Text + " like " + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "Org_descripcion" Then
        cadena_busqueda = "fc_organismo_financiamiento." + CmbCAMPO.Text + " like " + "'%" + TxtValor + "%'"
    End If
    
    If CmbCAMPO = "codigo_solicitud" Then
        cadena_busqueda = "pagos." + CmbCAMPO.Text + CmbOPERADOR + "'" + TxtValor + "'"
    End If
    If CmbCAMPO = "codigo_orden" Then
        cadena_busqueda = "pagos." + CmbCAMPO.Text + CmbOPERADOR + "'" + TxtValor + "'"
    End If
    
    'Realizar la busqueda dado un criterio
    Set rsComprobante = New ADODB.Recordset
    If cadena_busqueda <> "" Then
        Set rsComprobante = New ADODB.Recordset
        If OptTransferencias.Value = True Then
        rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf='C' and  " & cadena_busqueda & " order by  numero_cheque_trf", db, adOpenKeyset, adLockOptimistic
        End If
        If OptCheques.Value = True Then
        rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf='T' and  " & cadena_busqueda & " order by  numero_cheque_trf", db, adOpenKeyset, adLockOptimistic
        End If
'        rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf,Pagos.codigo_pago,pago_detalle.monto_Bolivianos,pago_detalle.fecha_pago,fc_beneficiario.denominacion_beneficiario, pago_detalle.cta_codigo,pagos.org_codigo,pago_detalle.literal, fc_bancos.Bco_descripcion_larga " & _
'               "FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.ges_gestion = pago_detalle.Ges_gestion) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.codigo_pago = pago_detalle.codigo_pago)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo) AND (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo WHERE " & cadena_busqueda & " order by  numero_cheque_trf", db, adOpenKeyset, adLockOptimistic
        If rsComprobante.RecordCount > 0 Then
            NoRegistros = rsComprobante.RecordCount
            'AdoCuenta.Caption = NoRegistros
            Set DtGPagos.DataSource = rsComprobante
            Set AdoPagos.Recordset = rsComprobante
        Else
            Set DtGPagos.DataSource = rsNada
            'Set AdoPagos.Recordset = rsNada
        End If
    Else
        MsgBox "Coloque datos"
    End If
    FraBusqueda.Visible = False
End Sub

Private Sub CmdEntregado_Click()
Dim i As Integer
Dim Cheque_Inicial As Long
Dim Cheque_Final As Long
Dim s As String
Dim k As Long
Dim sino As Variant

s = ""

sino = MsgBox("Está seguro de colocar este status de entregado?", vbYesNo + vbQuestion, "Atenciòn")
If sino = vbYes Then

If TxtCheques.Text <> "" Then
  Entregado_Lista
  Exit Sub
'Else
'  MsgBox "No existen Nros. de cheque", vbInformation + vbCritical
'  Exit Sub
End If
    If TxtInicial.Text <> "" Then
            Cheque_Inicial = Val(TxtInicial.Text)
            Cheque_Final = Val(TxtFinal.Text)
            For k = Cheque_Inicial To Cheque_Final Step 1
                s = ""
                T = CStr(k)
                Select Case Len(T)
                       Case 1
                            s = "0000" + CStr(k)
                       Case 2
                            s = "000" + CStr(k)
                       Case 3
                            s = "00" + CStr(k)
                       Case 4
                            s = "0" + CStr(k)
                       Case 5
                            s = CStr(k)
                End Select
                
                 LstCheques.AddItem s
            Next k
            For i = 0 To LstCheques.ListCount - 1
                LstCheques.ListIndex = i
                If rsCheques.State = 1 Then rsCheques.Close
                rsCheques.Open "SELECT * FROM to_cheques_operaciones where numero_cheque='" & LstCheques.Text & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
                If rsCheques.RecordCount > 0 Then
                        rsCheques("estado_entregado") = "S"
                        If OptCheques.Value = True Then
                             rsCheques("Cheq_Transf") = "C"
                        End If
                        If OptTransferencias.Value = True Then
                            rsCheques("Cheq_Transf") = "T"
                        End If
                
                Else
                        rsCheques.AddNew
                        rsCheques("numero_cheque") = LstCheques.Text
                        rsCheques("estado_entregado") = "S"
                        If OptCheques.Value = True Then
                             rsCheques("Cheq_Transf") = "C"
                        End If
                        If OptTransferencias.Value = True Then
                            rsCheques("Cheq_Transf") = "T"
                        End If
                        
                End If
                MsgBox "Satisfactorio", vbInformation + vbCritical
                rsCheques("usr_usuario") = LblUsuario.Caption
                rsCheques("fecha_registro") = DTPFechaRegistro.Value 'Date
                rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
                
                rsCheques.Update
            Next i
    End If
    
    
    LstCheques.ListIndex = 0
    If LstCheques.Text <> "" Then
       If rsCheques.State = 1 Then rsCheques.Close
       For i = 0 To LstCheques.ListCount - 1
                LstCheques.ListIndex = i
                LstCuenta.ListIndex = i
                If DtGPagos.Columns(0) = "" Then
                    MsgBox "No existe cheque ", vbInformation + vbCritical, "Validación de datos"
                    Exit Sub
                End If
                If rsCheques.State = 1 Then rsCheques.Close
                rsCheques.Open "SELECT * FROM to_cheques_operaciones WHERE  numero_cheque= '" & LstCheques.Text & "' and cta_codigo= '" & LstCuenta.Text & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
                If rsCheques.RecordCount > 0 Then
                
                    rsCheques("estado_entregado") = "S"
                    If OptCheques.Value = True Then
                         rsCheques("Cheq_Transf") = "C"
                    End If
                    If OptTransferencias.Value = True Then
                         rsCheques("Cheq_Transf") = "T"
                    End If
                    
                Else
                    rsCheques.AddNew
                    'rsCheques("numero_cheque") = AdoPagos.Recordset("numero_cheque_trf")
                    rsCheques("numero_cheque") = LstCheques.Text
                    rsCheques("cta_codigo") = LstCuenta.Text
                    rsCheques("estado_entregado") = "S"
                    If OptCheques.Value = True Then
                         rsCheques("Cheq_Transf") = "C"
                    End If
                    If OptTransferencias.Value = True Then
                         rsCheques("Cheq_Transf") = "T"
                    End If
                
                End If
                rsCheques("usr_usuario") = LblUsuario.Caption
                rsCheques("fecha_registro") = DTPFechaRegistro.Value 'Date
                rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
                rsCheques.Update
        Next i
        MsgBox "Satisfactorio", vbInformation + vbCritical
    End If
    
        
        'MsgBox "Cheque " + AdoPagos.Recordset("numero_cheque_trf") + " Entregado"
    '    Set rsCheques = New ADODB.Recordset
    '    If rsCheques.State = 1 Then rsCheques.Close
    '    rsCheques.Open "SELECT * FROM to_cheques WHERE numero_cheque = '" & AdoPagos.Recordset("numero_cheque_trf") & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
    '    If rsCheques.RecordCount > 0 Then
    '        rsCheques("estado_entregado") = "S"
    '    Else
    '        rsCheques.AddNew
    '        rsCheques("numero_cheque") = AdoPagos.Recordset("numero_cheque_trf")
    '        rsCheques("estado_entregado") = "S"
    '    End If
    '        rsCheques.Update
End If
End Sub

Private Sub Cmdimprimir_Click()
    FrmDesplegado.Show
End Sub

Private Sub CmdLimpiar_Click()
    LstCheques.Clear
    LstCuenta.Clear
End Sub

Private Sub CmdRestaurar_Click()
       Set rsComprobante = New ADODB.Recordset
    rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf='C'", db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        Set DtGPagos.DataSource = rsComprobante
        Set AdoPagos.Recordset = rsComprobante
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub DtcCtaTGN_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtcCtaTGN.BoundText
    DtCCuentaOrigen.BoundText = DtcCtaTGN.BoundText
End Sub

Private Sub DtCCuentaOrigen_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtCCuentaOrigen.BoundText
    DtcCtaTGN.BoundText = DtCCuentaOrigen.BoundText
End Sub

Private Sub DtCCuentaOrigenDes_Click(Area As Integer)
   DtcCtaTGN.BoundText = DtCCuentaOrigenDes.BoundText
   DtCCuentaOrigen.BoundText = DtCCuentaOrigenDes.BoundText
End Sub

Private Sub DtGPagos_Click()
Dim bandera As Integer
Dim i As Integer

If DtGPagos.Columns(0) = "" Then
    MsgBox "No existe cheque", vbInformation + vbCritical, "Validación de datos"
    Exit Sub
End If
    TxtInicial.Text = ""
    TxtFinal.Text = ""
    bandera = 0
    For i = 0 To LstCheques.ListCount - 1
         LstCheques.ListIndex = i
         If LstCheques.Text = DtGPagos.Columns(0) Then
              bandera = 1
         End If
    Next i
    If bandera = 0 Then
        LstCheques.AddItem DtGPagos.Columns(0)
        LstCuenta.AddItem DtGPagos.Columns(7)
    End If
FraBusca.Visible = False
End Sub

Private Sub DtGPagos_HeadClick(ByVal ColIndex As Integer)
    FraBusca.Visible = True
    Set rsComprobante = New ADODB.Recordset
    If rsComprobante.State = 1 Then rsComprobante.Close
    Select Case ColIndex
        Case 0
                rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo,pago_detalle.fecha_pago  " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.numero_cheque_trf", db, adOpenKeyset, adLockOptimistic
        Case 1
                rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo,pago_detalle.fecha_pago  " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by fc_beneficiario.denominacion_beneficiario", db, adOpenKeyset, adLockOptimistic
        Case 2
                rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo,pago_detalle.fecha_pago  " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.monto_bolivianos", db, adOpenKeyset, adLockOptimistic
        Case 3
                rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo,pago_detalle.fecha_pago  " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.codigo_pago", db, adOpenKeyset, adLockOptimistic
        Case 3
                rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo,pago_detalle.fecha_pago  " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.monto_dolares", db, adOpenKeyset, adLockOptimistic
    End Select
    Set DtGPagos.DataSource = rsComprobante

End Sub

Private Sub Form_Load()
    'SELECT Pagos.codigo_pago, Pagos.org_codigo, fc_organismo_financiamiento.Org_descripcion, pago_detalle.fecha_pago, pago_detalle.saldo_bolivianos, pago_detalle.monto_Bolivianos, pago_detalle.tipo_cambio, pago_detalle.codigo_beneficiario, fc_beneficiario.denominacion_beneficiario, Pagos.justificacion, pago_detalle.cheque_o_trf, pago_detalle.numero_cheque_trf, pago_detalle.cta_codigo, fc_bancos.Bco_descripcion_larga, pago_detalle.literal, fc_cuenta_bancaria.Cta_descripcion_larga
    'FROM (((Pagos INNER JOIN pago_detalle ON (Pagos.codigo_pago = pago_detalle.codigo_pago) AND (Pagos.org_codigo = pago_detalle.org_codigo) AND (Pagos.ges_gestion = pago_detalle.Ges_gestion)) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_organismo_financiamiento ON Pagos.org_codigo = fc_organismo_financiamiento.Org_codigo) INNER JOIN (fc_bancos INNER JOIN fc_cuenta_bancaria ON (fc_bancos.Ges_gestion = fc_cuenta_bancaria.Ges_gestion) AND (fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo)) ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo;
    DTPFechaRegistro.Value = Date
    Set rsComprobante = New ADODB.Recordset
    'rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf='C'", db, adOpenKeyset, adLockOptimistic
                       
                       
    rsComprobante.Open "SELECT     dbo.pago_detalle.numero_cheque_trf, dbo.FC_BENEFICIARIO.denominacion_beneficiario, dbo.pago_detalle.monto_Bolivianos, dbo.pago_detalle.codigo_pago, dbo.pago_detalle.monto_Dolares, dbo.pago_detalle.tipo_cambio, dbo.fc_cuenta_bancaria.Cta_descripcion_larga, dbo.fc_cuenta_bancaria.Cta_codigo, dbo.pago_detalle.Org_codigo, dbo.pago_detalle.fecha_pago AS [ & _ ] FROM         dbo.pago_detalle INNER JOIN dbo.FC_BENEFICIARIO ON dbo.pago_detalle.codigo_beneficiario COLLATE Modern_Spanish_CI_AS = dbo.FC_BENEFICIARIO.codigo_beneficiario INNER JOIN dbo.fc_cuenta_bancaria ON dbo.pago_detalle.cta_codigo = dbo.fc_cuenta_bancaria.Cta_codigo WHERE (dbo.pago_detalle.cheque_o_trf = 'C')", db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        Set DtGPagos.DataSource = rsComprobante
        Set AdoPagos.Recordset = rsComprobante
    End If
    
    'Abriendo cuenta bancaria
    Set rscuenta = New ADODB.Recordset
    rscuenta.Open "select * from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
    Set AdoCuenta.Recordset = rscuenta
    DtCCuentaOrigenDes.BoundText = DtCCuentaOrigen.BoundText
    DtcCtaTGN.BoundText = DtCCuentaOrigen.BoundText

	Call SeguridadSet(Me)
End Sub
Public Sub Determina_Cheques()
Dim i As Integer
    NrosChequeImprimir = " "
    For i = 0 To LstCheques.ListCount - 2
        LstCheques.ListIndex = i
        NrosChequeImprimir = NrosChequeImprimir & "numero_cheque= " & "'" & LstCheques.Text & "'" & " Or "
    Next i
    LstCheques.ListIndex = i
    NrosChequeImprimir = NrosChequeImprimir + "numero_cheque = " & "'" & LstCheques.Text & "'"
End Sub

Private Sub LstCheques_Click()
    'LstComprobante.RemoveItem ListIndex
End Sub

Private Sub LstCheques_DblClick()
    'LstCheques.RemoveItem ListIndex
End Sub

Private Sub OptCheques_Click()
    lblTitulo.Caption = "OPERACIONES CHEQUES"
    CmdEntregado.Enabled = True
    CmdDevuelto.Enabled = True
    CmdAnulado.Enabled = False
    CmdCobrado.Enabled = True
    
    '
    If rsComprobante.State = 1 Then rsComprobante.Close
    rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf='C'", db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        Set DtGPagos.DataSource = rsComprobante
        Set AdoPagos.Recordset = rsComprobante
    End If
End Sub

Private Sub OptTransferencias_Click()
    lblTitulo.Caption = "OPERACIONES TRANSFERENCIAS"
    CmdEntregado.Enabled = False
    CmdDevuelto.Enabled = False
    CmdAnulado.Enabled = True
    CmdCobrado.Enabled = False
    '
    If rsComprobante.State = 1 Then rsComprobante.Close
    rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
                       "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.cheque_o_trf='T' order by numero_cheque_trf", db, adOpenKeyset, adLockOptimistic
    If rsComprobante.RecordCount > 0 Then
        Set DtGPagos.DataSource = rsComprobante
        Set AdoPagos.Recordset = rsComprobante
    End If
    
End Sub

Private Sub TxtCheques_Change()
  LstCheques.Clear
End Sub

Public Sub Entregado_Lista()

'========================================================================================
' Módulo:                   Entregado_Lista
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmActivacionCheques.frm
' Descipción :              Se coloca el status de entregado
'                           de acuerdo a una lista y en el caso de cheques
'                           de acuerdo a la cuenta bancaria
'                           si se trata de cheques
' Autor:                    Celia Elena Tarquino Peralta
' Versión:                  2.0
'========================================================================================

Dim AUX, NUMERO As String
Dim Car As String
Dim i As Integer
Dim LONGITUD As Integer

NUMERO = ""
AUX = TxtCheques.Text
LONGITUD = Len(AUX)
  While (LONGITUD + 1 > 0)
      i = i + 1
      Car = Mid(AUX, i, 1)
      LONGITUD = LONGITUD - 1
      If Car <> "," And Car <> "" Then
         NUMERO = NUMERO + Car
      Else
                MsgBox NUMERO
                T = CStr(NUMERO)
                Select Case Len(T)
                       Case 1
                            s = "0000" + CStr(NUMERO)
                       Case 2
                            s = "000" + CStr(NUMERO)
                       Case 3
                            s = "00" + CStr(NUMERO)
                       Case 4
                            s = "0" + CStr(NUMERO)
                       Case 5
                            s = CStr(NUMERO)
                End Select
                Set rsComprobante = New ADODB.Recordset
                rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
                                   "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.numero_cheque_trf='" & s & "' and pago_detalle.cheque_o_trf='C' and fc_cuenta_bancaria.Cta_codigo='" & DtCCuentaOrigen.Text & "'", db, adOpenKeyset, adLockOptimistic
                If rsComprobante.RecordCount > 0 Then
                If rsCheques.State = 1 Then rsCheques.Close
                rsCheques.Open "SELECT * FROM to_cheques_operaciones where numero_cheque='" & s & "' and Cta_codigo='" & DtCCuentaOrigen.Text & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
                If rsCheques.RecordCount > 0 Then
                        rsCheques("estado_entregado") = "S"
                Else
                        rsCheques.AddNew
                        rsCheques("numero_cheque") = s
                        rsCheques("estado_entregado") = "S"
                End If
                rsCheques("usr_usuario") = LblUsuario.Caption
                rsCheques("fecha_registro") = DTPFechaRegistro.Value 'Date
                rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
                
                rsCheques.Update
             End If
            NUMERO = ""
         End If
  Wend
End Sub

Public Sub Devuelto_Lista()
'========================================================================================
' Módulo:                   Devuelto_Lista
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmActivacionCheques.frm
' Descipción :              Se coloca el status de devuelto
'                           de acuerdo a una lista y en el caso de cheques
'                           de acuerdo a la cuenta bancaria
'                           si se trata de cheques
' Autor:                    Celia Elena Tarquino Peralta
' Versión:                  2.0
'========================================================================================

Dim AUX, NUMERO As String
Dim Car As String
Dim i As Integer
Dim LONGITUD As Integer

NUMERO = ""
AUX = TxtCheques.Text
LONGITUD = Len(AUX)
  While (LONGITUD + 1 > 0)
      i = i + 1
      Car = Mid(AUX, i, 1)
      LONGITUD = LONGITUD - 1
      If Car <> "," And Car <> "" Then
         NUMERO = NUMERO + Car
      Else
                MsgBox NUMERO
                T = CStr(NUMERO)
                Select Case Len(T)
                       Case 1
                            s = "0000" + CStr(NUMERO)
                       Case 2
                            s = "000" + CStr(NUMERO)
                       Case 3
                            s = "00" + CStr(NUMERO)
                       Case 4
                            s = "0" + CStr(NUMERO)
                       Case 5
                            s = CStr(NUMERO)
                End Select
                Set rsComprobante = New ADODB.Recordset
                rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
                                   "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.numero_cheque_trf='" & s & "' and pago_detalle.cheque_o_trf='C' and fc_cuenta_bancaria.cta_codigo='" & DtCCuentaOrigen.Text & "'", db, adOpenKeyset, adLockOptimistic
                If rsComprobante.RecordCount > 0 Then
                If rsCheques.State = 1 Then rsCheques.Close
                rsCheques.Open "SELECT * FROM to_cheques_operaciones where numero_cheque='" & s & "' and Cta_codigo='" & DtCCuentaOrigen.Text & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
                If rsCheques.RecordCount > 0 Then
                        rsCheques("estado_devuelto") = "S"
                Else
                        rsCheques.AddNew
                        rsCheques("numero_cheque") = s
                        rsCheques("estado_devuelto") = "S"
                End If
                rsCheques("usr_usuario") = LblUsuario.Caption
                rsCheques("fecha_registro") = DTPFechaRegistro.Value 'Date
                rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
                
                rsCheques.Update
             End If
            NUMERO = ""
         End If
  Wend

End Sub

Public Sub Anulado_Lista()
'========================================================================================
' Módulo:                   Anulado_Lista
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmActivacionCheques.frm
' Descipción :              Se Anulan de la lista dada de acuerdo a la cuenta bancaria
'                           si se trata de cheques
' Autor:                    Celia Elena Tarquino Peralta
' Versión:                  2.0
'========================================================================================
Dim AUX, NUMERO As String
Dim Car As String
Dim i As Integer
Dim LONGITUD As Integer

NUMERO = ""
AUX = TxtCheques.Text
LONGITUD = Len(AUX)
  While (LONGITUD + 1 > 0)
      i = i + 1
      Car = Mid(AUX, i, 1)
      LONGITUD = LONGITUD - 1
      If Car <> "," And Car <> "" Then
         NUMERO = NUMERO + Car
      Else
                MsgBox NUMERO
                T = CStr(NUMERO)
                Select Case Len(T)
                       Case 1
                            s = "0000" + CStr(NUMERO)
                       Case 2
                            s = "000" + CStr(NUMERO)
                       Case 3
                            s = "00" + CStr(NUMERO)
                       Case 4
                            s = "0" + CStr(NUMERO)
                       Case 5
                            s = CStr(NUMERO)
                End Select
                Set rsComprobante = New ADODB.Recordset
                rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
                                   "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.numero_cheque_trf='" & s & "' and pago_detalle.cheque_o_trf='C' and fc_cuenta_bancaria.cta_codigo='" & DtCCuentaOrigen.Text & "'", db, adOpenKeyset, adLockOptimistic
                If rsComprobante.RecordCount > 0 Then
                If rsCheques.State = 1 Then rsCheques.Close
                rsCheques.Open "SELECT * FROM to_cheques_operaciones where numero_cheque='" & s & "' and Cta_codigo='" & DtCCuentaOrigen.Text & "' order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
                If rsCheques.RecordCount > 0 Then
                        rsCheques("estado_anulado") = "S"
                Else
                        rsCheques.AddNew
                        rsCheques("numero_cheque") = s
                        rsCheques("estado_anulado") = "S"
                End If
                rsCheques("usr_usuario") = LblUsuario.Caption
                rsCheques("fecha_registro") = DTPFechaRegistro.Value 'Date
                rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
                rsCheques.Update
             End If
            NUMERO = ""
         End If
  Wend

End Sub

Public Sub Cobrado_Lista()
'========================================================================================
' Módulo:                   Cobrado_Lista
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmActivacionCheques.frm
' Descipción :              Se coloca el status de cobrado
'                           de acuerdo a una lista y en el caso de cheques
'                           de acuerdo a la cuenta bancaria
'                           si se trata de cheques
' Autor:                    Celia Elena Tarquino Peralta
' Versión:                  2.0
'========================================================================================

Dim AUX, NUMERO As String
Dim Car As String
Dim i As Integer
Dim LONGITUD As Integer

NUMERO = ""
AUX = TxtCheques.Text
LONGITUD = Len(AUX)
  While (LONGITUD + 1 > 0)
      i = i + 1
      Car = Mid(AUX, i, 1)
      LONGITUD = LONGITUD - 1
      If Car <> "," And Car <> "" Then
         NUMERO = NUMERO + Car
      Else
                MsgBox NUMERO
                T = CStr(NUMERO)
                Select Case Len(T)
                       Case 1
                            s = "0000" + CStr(NUMERO)
                       Case 2
                            s = "000" + CStr(NUMERO)
                       Case 3
                            s = "00" + CStr(NUMERO)
                       Case 4
                            s = "0" + CStr(NUMERO)
                       Case 5
                            s = CStr(NUMERO)
                End Select
                Set rsComprobante = New ADODB.Recordset
                rsComprobante.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo, pago_detalle.fecha_pago " & _
                                   "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where pago_detalle.numero_cheque_trf='" & s & "' and pago_detalle.cheque_o_trf='C' and fc_cuenta_bancaria.cta_codigo='" & DtCCuentaOrigen.Text & "'", db, adOpenKeyset, adLockOptimistic
                If rsComprobante.RecordCount > 0 Then
                If rsCheques.State = 1 Then rsCheques.Close
                rsCheques.Open "SELECT * FROM to_cheques_operaciones where numero_cheque='" & s & "' and Cta_codigo='" & DtCCuentaOrigen.Text & "'  order by  numero_cheque", db, adOpenKeyset, adLockOptimistic
                If rsCheques.RecordCount > 0 Then
                        rsCheques("estado_cobrado") = "S"
                Else
                        rsCheques.AddNew
                        rsCheques("numero_cheque") = s
                        rsCheques("estado_cobrado") = "S"
                End If
                rsCheques("usr_usuario") = LblUsuario.Caption
                rsCheques("fecha_registro") = DTPFechaRegistro.Value 'Date
                rsCheques("hora_registro") = Format(Time, "hh:mm:ss")
                
                rsCheques.Update
             End If
            NUMERO = ""
         End If
  Wend

End Sub

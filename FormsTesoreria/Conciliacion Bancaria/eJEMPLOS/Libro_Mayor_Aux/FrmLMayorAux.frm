VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmLMayorAux 
   Caption         =   "Reportes Contables - Libro Mayor Auxiliar"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra_Busqueda 
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1665
      Left            =   1560
      TabIndex        =   35
      Top             =   4080
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   5295
         Begin VB.TextBox TxtValor 
            Height          =   336
            Left            =   2520
            MultiLine       =   -1  'True
            TabIndex        =   43
            Text            =   "FrmLMayorAux.frx":0000
            Top             =   240
            Width           =   2640
         End
         Begin VB.ComboBox CboOperador 
            Height          =   315
            ItemData        =   "FrmLMayorAux.frx":0002
            Left            =   1560
            List            =   "FrmLMayorAux.frx":000C
            TabIndex        =   42
            Text            =   "="
            Top             =   240
            Width           =   915
         End
         Begin VB.ComboBox CboCampo 
            Height          =   315
            ItemData        =   "FrmLMayorAux.frx":0019
            Left            =   120
            List            =   "FrmLMayorAux.frx":0023
            TabIndex        =   41
            Top             =   240
            Width           =   1284
         End
      End
      Begin VB.Frame Frame5 
         Height          =   615
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   5295
         Begin VB.CommandButton Cmd_Normal 
            Caption         =   "Normal"
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
            Left            =   2040
            TabIndex        =   39
            Top             =   240
            Width           =   1125
         End
         Begin VB.CommandButton Cmd_BSalir 
            Caption         =   "Salir"
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
            Left            =   3840
            TabIndex        =   38
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmdEjecutar 
            Caption         =   "Ejecutar"
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
            Left            =   360
            TabIndex        =   37
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin Crystal.CrystalReport CryLMayorCtaBancaria 
      Left            =   720
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CryLMayorBenef 
      Left            =   120
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar PRB 
      Height          =   360
      Left            =   5400
      TabIndex        =   31
      Top             =   7560
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Height          =   4110
      Left            =   45
      TabIndex        =   30
      Top             =   780
      Width           =   1245
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   780
         Left            =   120
         Picture         =   "FrmLMayorAux.frx":0057
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Cmdsalir 
         Caption         =   "Salir"
         Height          =   780
         Left            =   135
         Picture         =   "FrmLMayorAux.frx":0499
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3180
         Width           =   855
      End
      Begin VB.CommandButton Cmdcancelar 
         Caption         =   "Cancelar"
         Height          =   780
         Left            =   135
         Picture         =   "FrmLMayorAux.frx":08DB
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2145
         Width           =   855
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   120
         Picture         =   "FrmLMayorAux.frx":0D1D
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1212
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   750
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   8535
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Reportes Contables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   2520
         TabIndex        =   29
         Top             =   255
         Width           =   3945
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4110
      Left            =   1440
      TabIndex        =   17
      Top             =   840
      Width           =   7245
      Begin VB.ComboBox cboCtaBancaria 
         Height          =   315
         Left            =   2160
         TabIndex        =   32
         Text            =   "Combo1"
         Top             =   2160
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.ComboBox cbosubcta1 
         Height          =   288
         Left            =   1335
         TabIndex        =   2
         Top             =   1005
         Width           =   1155
      End
      Begin VB.ComboBox cbosubcta2 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   1590
         Width           =   1140
      End
      Begin VB.ComboBox cbocta 
         Height          =   288
         Left            =   1350
         TabIndex        =   1
         Top             =   420
         Width           =   1170
      End
      Begin VB.CheckBox Chkaux1 
         Caption         =   "Auxiliar 1"
         Height          =   195
         Left            =   255
         TabIndex        =   4
         Top             =   2205
         Width           =   975
      End
      Begin VB.CheckBox Chkaux2 
         Caption         =   "Auxiliar 2"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   2685
         Width           =   1005
      End
      Begin VB.CheckBox Chkaux3 
         Caption         =   "Auxiliar 3"
         Height          =   270
         Left            =   240
         TabIndex        =   8
         Top             =   3135
         Width           =   1080
      End
      Begin VB.TextBox txtbusca1 
         Height          =   330
         Left            =   2160
         TabIndex        =   5
         Top             =   2115
         Width           =   4680
      End
      Begin VB.TextBox txtax1 
         Height          =   330
         Left            =   1344
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2100
         Width           =   585
      End
      Begin VB.TextBox Txtax2 
         Height          =   330
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2580
         Width           =   585
      End
      Begin VB.TextBox txtax3 
         Height          =   330
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3075
         Width           =   585
      End
      Begin VB.TextBox Txtbusca2 
         Height          =   330
         Left            =   2160
         TabIndex        =   7
         Top             =   2580
         Width           =   4665
      End
      Begin VB.TextBox Txtbusca3 
         Height          =   330
         Left            =   2175
         TabIndex        =   9
         Top             =   3075
         Width           =   4680
      End
      Begin MSComCtl2.DTPicker DTPfin 
         Height          =   360
         Left            =   3480
         TabIndex        =   11
         Top             =   3585
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   635
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   36614
      End
      Begin MSComCtl2.DTPicker DTPinicio 
         Height          =   345
         Left            =   1335
         TabIndex        =   10
         Top             =   3600
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   36526
         MaxDate         =   36526
         MinDate         =   36526
      End
      Begin VB.Label Lblsub1 
         Height          =   375
         Left            =   2760
         TabIndex        =   27
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label lblcuenta 
         Height          =   375
         Left            =   2880
         TabIndex        =   26
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Subcuenta 2"
         Height          =   195
         Left            =   210
         TabIndex        =   25
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Subcuenta 1"
         Height          =   195
         Left            =   210
         TabIndex        =   24
         Top             =   1050
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lbsub2 
         Height          =   495
         Left            =   2760
         TabIndex        =   22
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label4 
         Caption         =   "Desde:"
         Height          =   240
         Left            =   570
         TabIndex        =   21
         Top             =   3675
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Hasta:"
         Height          =   240
         Left            =   2880
         TabIndex        =   20
         Top             =   3690
         Width           =   645
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2760
      Left            =   105
      TabIndex        =   0
      Top             =   4935
      Width           =   8580
      Begin MSDataGridLib.DataGrid DTGBanco 
         Height          =   2295
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4048
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
         Caption         =   "CUENTAS BANCARIAS"
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
      Begin MSDataGridLib.DataGrid DtGbenef 
         Height          =   2370
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   4180
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         Caption         =   "BENEFICIARIOS"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "codigo_beneficiario"
            Caption         =   "Código Beneficiario"
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
            DataField       =   "denominacion_beneficiario"
            Caption         =   "Denominación"
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
   Begin Crystal.CrystalReport CryLMayor 
      Left            =   1320
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmLMayorAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/************  RECORDSETS

Dim sql1 As String
Dim sql2 As String
Dim nombenef As String
Dim combenef As ADODB.Command
Dim comctabancaria As ADODB.Command
Dim rsplanctas As ADODB.Recordset
Dim rscuentas As ADODB.Recordset
Dim rsnombresub1 As ADODB.Recordset
Dim rssubcuenta As ADODB.Recordset
Dim rscta_bancaria As ADODB.Recordset
Dim rsBeneficiario As ADODB.Recordset
Dim rssaldos As ADODB.Recordset
Dim rsctabancaria As ADODB.Recordset
Dim SaldoIBs As Double
Dim SaldoISus As Double
Dim benef As String
Dim ctabancaria As String
Dim nombanco As String
Dim nomctabancaria As String

'/**********


Dim existereporte As New ADODB.Recordset
Dim reporte As New ADODB.Recordset
Dim BUSCA As Integer
Dim parametro As String
Dim denominacion As String
Public aux1 As String
Public aux2 As String
Public aux3 As String
'Dim consul As New ADODB.Recordset
Dim saldobs As Double
Dim saldosus As Double
Dim saldobs1 As Double
Dim saldosus1 As Double
Dim auxsaldobs As Double
Dim auxsaldosus As Double
''Private Sub cboaux_LostFocus()
''If Me.cboaux = "01" Then
''Me.Frr01.Visible = True
''End If
''End Sub

Private Sub cbocta_Click()
  Me.cbosubcta1.Clear
  Me.cbosubcta2.Clear

  rsplanctas.MoveFirst
  rsplanctas.Find "cuenta=" & "'" & Trim(cbocta.Text) & "'"
  Me.Lblcuenta = rsplanctas!NombreCta
  If rscuentas.State = adStateOpen Then rscuentas.Close
  
  rscuentas.Open "SELECT Cuenta, SubCta1 FROM CC_Plan_Cuentas GROUP BY Cuenta, SubCta1 HAVING (SubCta1 <> '00') AND (Cuenta = '" & Trim(Me.cbocta.Text) & "')", db, adOpenKeyset, adLockReadOnly
  Do While Not rscuentas.EOF
    Me.cbosubcta1.AddItem rscuentas!Subcta1
    rscuentas.MoveNext
  Loop
  If rscuentas.RecordCount = 0 Then
  Me.cbosubcta1.AddItem "00"
  End If

End Sub
Private Sub cbosubcta1_Click()
On Error GoTo Laberror1
Me.cbosubcta2.Clear

  If rsnombresub1.State = adStateOpen Then rsnombresub1.Close
  rsnombresub1.Open "SELECT NombreCta FROM CC_Plan_Cuentas WHERE   (SubCta2 = '00') AND (Cuenta = '" & Trim(Me.cbocta.Text) & "') AND (SubCta1 ='" & (Me.cbosubcta1.Text) & "')", db, adOpenKeyset, adLockReadOnly
  Me.Lblsub1 = rsnombresub1!NombreCta
  If rssubcuenta.State = adStateOpen Then rssubcuenta.Close
  rssubcuenta.Open "SELECT Cuenta, SubCta1, SubCta2, NombreCta, Aux1, Aux2, Aux3 FROM CC_Plan_Cuentas WHERE (Cuenta ='" & Trim(Me.cbocta.Text) & "') AND (SubCta1 ='" & Trim(Me.cbosubcta1.Text) & "')", db, adOpenKeyset, adLockReadOnly
  If rssubcuenta.RecordCount = 0 Then
    Me.cbosubcta2 = "00"
    Else
      rssubcuenta.MoveFirst
      Do While Not rssubcuenta.EOF
        Me.cbosubcta2.AddItem rssubcuenta!Subcta2
        rssubcuenta.MoveNext
      Loop
    End If

Exit Sub
Laberror1:
If Err.Number = 3021 Then
 MsgBox "Elija una cuenta", vbCritical + vbDefaultButton1
 Me.cbocta.SetFocus
End If
End Sub
Private Sub cbosubcta2_Click()
'On Error GoTo labelerr2
    Me.CmdBuscar.Enabled = True
    Me.Chkaux1.Enabled = True
    Me.Chkaux2.Enabled = True
    Me.Chkaux3.Enabled = True
    Me.txtax1.Enabled = True
    Me.Txtax2.Enabled = True
    Me.txtax3.Enabled = True
    Me.txtbusca1.Enabled = True
    Me.Txtbusca2.Enabled = True
    Me.Txtbusca3.Enabled = True
    With rssubcuenta
      .MoveFirst
      .Find "subcta2=" & "'" & Trim(Me.cbosubcta2) & "'"
      Me.lbsub2 = !NombreCta
      Me.txtax1 = !aux1
      Me.Txtax2 = !aux2
      Me.txtax3 = !aux3
      If !aux1 = "00" Then
        Me.Chkaux1.Enabled = False
        Me.txtax1.Enabled = False
        Me.txtbusca1.Enabled = False
      End If
      If !aux2 = "00" Then
        Me.Chkaux2.Enabled = False
        Me.Txtax2.Enabled = False
        Me.Txtbusca2.Enabled = False
      End If
      If !aux3 = "00" Then
        Me.Chkaux3.Enabled = False
        Me.txtax3.Enabled = False
        Me.Txtbusca3.Enabled = False
      End If
      If Me.Chkaux1.Enabled = True And Me.Chkaux2.Enabled = False And Me.Chkaux3.Enabled = False Then
        Me.Chkaux1.Value = 1
      End If
      If Me.Chkaux1.Enabled = False And Me.Chkaux2.Enabled = True And Me.Chkaux3.Enabled = False Then
        Me.Chkaux2.Value = 1
      End If
      If Me.Chkaux1.Enabled = False And Me.Chkaux2.Enabled = False And Me.Chkaux3.Enabled = True Then
        Me.Chkaux3.Value = 1
      End If
    End With
    
    If (Me.txtax1 <> "00" And Me.txtax1 <> "01" And Me.txtax1 <> "02") Then
      f = 1
      Me.Chkaux1.Enabled = False
      Me.txtax1.Enabled = False
      Me.txtbusca1.Enabled = False
    End If
    If (Me.Txtax2 <> "00" And Me.Txtax2 <> "01" And Me.Txtax2 <> "02") Then
      f = 2
      Me.Chkaux1.Enabled = False
      Me.txtax1.Enabled = False
      Me.txtbusca1.Enabled = False
    End If
    If (Me.txtax3 <> "00" And Me.txtax3 <> "01" And Me.txtax3 <> "02") Then
      f = 3
      Me.Chkaux1.Enabled = False
      Me.txtax1.Enabled = False
      Me.txtbusca1.Enabled = False
    End If
    If f = 1 Or f = 2 Or f = 3 Then
        MsgBox "Por el momento solo se trabaja con Auxiliares de Beneficiarios y Ctas. Corrientes", vbInformation + vbDefaultButton1, "SAF/2000"
        Me.cbocta.SetFocus
    End If
    If Me.Chkaux1.Enabled = False And Me.Chkaux2.Enabled = False And Me.Chkaux3.Enabled = False Then
    Me.CmdBuscar.Enabled = False
    Else
'    Me.CmdAceptar.Enabled = False
    End If
If (Me.cbosubcta1.Text) = "00" And Me.cbosubcta2.Text = "00" Then
    'Me.CmdAceptar.Enabled = True
End If
'*******Se filtra si la cuenta es de bancos....
If Me.cbocta = "1111" And Me.cbosubcta1 = "02" Then
    Select Case Me.cbosubcta2
        Case "01"
            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_bancos.Bco_descripcion_larga FROM fc_cuenta_bancaria INNER JOIN " & _
                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '41' or fc_cuenta_bancaria.Fte_codigo = '10' order by fc_cuenta_bancaria.Cta_codigo"
        Case "02"
            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_bancos.Bco_descripcion_larga FROM fc_cuenta_bancaria INNER JOIN " & _
                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '43' order by fc_cuenta_bancaria.Cta_codigo"
        Case "03"
            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_bancos.Bco_descripcion_larga FROM fc_cuenta_bancaria INNER JOIN " & _
                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '80' order by fc_cuenta_bancaria.Cta_codigo"
     End Select
    Me.cboCtaBancaria.Clear
    If rscta_bancaria.State = 1 Then rscta_bancaria.Close
    rscta_bancaria.Open sql1, db, adOpenKeyset, adLockReadOnly
    If rscta_bancaria.RecordCount <> 0 Then
        rscta_bancaria.MoveFirst
    End If
        Do While Not rscta_bancaria.EOF
          cboCtaBancaria.AddItem rscta_bancaria!cta_codigo
          rscta_bancaria.MoveNext
        Loop
    Me.cboCtaBancaria.Visible = True
    Me.cboCtaBancaria.Text = Me.cboCtaBancaria.List(0)
    Me.txtbusca1.Visible = False
    Me.DTGBanco.Visible = True
    Me.DtGbenef.Visible = False
    Set Me.DTGBanco.DataSource = rscta_bancaria
End If

'************Se habilita la tabla de beneficiarios
    If Me.txtax1 = "01" Or Me.Txtax2 = "01" Or Me.txtax3 = "01" Then
        If rsBeneficiario.State = 1 Then rsBeneficiario.Close
        sql2 = "SELECT codigo_beneficiario, denominacion_beneficiario From fc_beneficiario order by denominacion_beneficiario"
        rsBeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
        Set Me.DtGbenef.DataSource = rsBeneficiario
        Me.DtGbenef.Visible = True
        Me.DTGBanco.Visible = False
        Me.txtbusca1.Visible = True
        Me.cboCtaBancaria.Visible = False
    End If
'****habilitamos boton de búsqueda
    If Me.txtax1 = "00" Or Me.txtax1 = "02" Then
        Me.CmdBuscar.Enabled = False
    Else
        Me.CmdBuscar.Enabled = True
    End If
    
    Exit Sub
labelerr2:
    If Err.Number = 3021 Then
      MsgBox "Elija una subcuenta", vbCritical + vbDefaultButton1
      Me.cbosubcta2.SetFocus
    End If
End Sub

Private Sub cbosubcta2_LostFocus()
If (Me.txtax1 = "01" And Me.Txtax2 = "00" And Me.txtax3 = "00") Then
  Me.DtGbenef.Visible = True
  Me.DTGBanco.Visible = False
End If
If (Me.txtax1 = "00" And Me.Txtax2 = "00" And Me.txtax3 = "00") Then
End If
End Sub

Private Sub Chkaux1_Click()
'habilita el grid de beneficiarios
If Me.Chkaux1.Value = 1 And (Me.txtax1 = "01" Or Me.Txtax2 = "01" Or Me.txtax3 = "01") Then
    Me.DtGbenef.Visible = True
    Me.DTGBanco.Visible = False
End If
'habilita el grid de cuentas corrientes
If Me.Chkaux1.Value = 1 And (Me.txtax1 = "02" Or Me.Txtax2 = "02" Or Me.txtax3 = "02") Then
    Me.DTGBanco.Visible = True
    Me.DtGbenef.Visible = False
End If
End Sub
Private Sub Chkaux2_Click()
'habilita el grid de beneficiarios
If Me.Chkaux2.Value = 1 And (Me.txtax1 = "01" Or Me.Txtax2 = "01" Or Me.txtax3 = "01") Then
    Me.DtGbenef.Visible = True
End If
'habilita el grid de cuentas corrientes
If Me.Chkaux2.Value = 1 And (Me.txtax1 = "02" Or Me.Txtax2 = "02" Or Me.txtax3 = "02") Then
    Me.DTGBanco.Visible = True
End If
End Sub
Private Sub Chkaux3_Click()
'habilita el grid de beneficiarios
If Me.Chkaux3.Value = 1 And (Me.txtax1 = "01" Or Me.Txtax2 = "01" Or Me.txtax3 = "01") Then
    Me.DtGbenef.Visible = True
End If
'habilita el grid de cuentas corrientes
If Me.Chkaux3.Value = 1 And (Me.txtax1 = "02" Or Me.Txtax2 = "02" Or Me.txtax3 = "02") Then
    Me.DTGBanco.Visible = True
End If
End Sub

Private Sub cmd_Ejecutar_Click()






End Sub

Private Sub Cmd_BSalir_Click()
Me.Fra_Busqueda.Visible = False
End Sub

Private Sub Cmd_Normal_Click()
  If rsBeneficiario.State = 1 Then rsBeneficiario.Close
    sql2 = "SELECT codigo_beneficiario, denominacion_beneficiario From fc_beneficiario order by denominacion_beneficiario"
    rsBeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
    Set Me.DtGbenef.DataSource = rsBeneficiario
End Sub

Private Sub CmdAceptar_Click()
If Me.cbocta.Text = "" Then
    MsgBox "Elija una cuenta", vbCritical + vbDefaultButton1
    Me.cbocta.SetFocus
    Exit Sub
End If
If Me.cbosubcta1.Text = "" Then
    MsgBox "Elija una subcuenta", vbCritical + vbDefaultButton1
    Me.cbosubcta1.SetFocus
    Exit Sub
End If
If Me.cbosubcta2.Text = "" Then
    MsgBox "Elija una subcuenta", vbCritical + vbDefaultButton1
    Me.cbosubcta2.SetFocus
    Exit Sub
End If
If Me.txtax1 = "02" Then
    If Me.cboCtaBancaria.Text = "" Then
        MsgBox "Elija una cuenta bancaria", vbCritical + vbDefaultButton1
        Me.cboCtaBancaria.SetFocus
        Exit Sub
    End If
End If
If Me.txtax1 = "01" Then
    If Me.txtbusca1.Text = "" Then
        MsgBox "Escriba un beneficiario", vbCritical + vbDefaultButton1
        Me.txtbusca1.SetFocus
        Exit Sub
    End If
End If
If (DTPinicio.Value > DTPfin.Value) Or (DTPfin.Value < DTPinicio.Value) Then
    MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
    Exit Sub
End If
If Me.txtax1 = "00" And Me.Txtax2 = "00" And Me.txtax3 = "00" Then
'****si la cuenta no tiene auxiliares
    Call Mayor000
Else
'****llamada al store procedure de Saldos para beneficiarios "SaldoBenef
'***** si el aux es 1
    If (Me.txtax1) = "01" Or (Me.Txtax2) = "01" Or (Me.txtax3) = "01" Then
        If rsBeneficiario.State = 1 Then rsBeneficiario.Close
        rsBeneficiario.Open "select * from fc_beneficiario where codigo_beneficiario = '" & Trim(Me.txtbusca1.Text) & "'", db, adOpenKeyset, adLockReadOnly
        If rsBeneficiario.RecordCount <> 0 Then
            nombenef = rsBeneficiario!denominacion_beneficiario
        Else
            nombenef = ""
        End If
        Dim IResult As Integer
        Set combenef = New ADODB.Command
        With combenef
            .CommandType = adCmdStoredProc
            .CommandText = "SaldoBenef"
            .Parameters.Append combenef.CreateParameter("FFInicio", adVarChar, adParamInput, 10)
            .Parameters.Append combenef.CreateParameter("FFFinal", adVarChar, adParamInput, 10)
            .Parameters.Append combenef.CreateParameter("cuenta", adVarChar, adParamInput, 5)
            .Parameters.Append combenef.CreateParameter("subcta1", adVarChar, adParamInput, 3)
            .Parameters.Append combenef.CreateParameter("subcta2", adVarChar, adParamInput, 3)
            .Parameters.Append combenef.CreateParameter("beneficiario", adVarChar, adParamInput, 15)
            .Parameters.Append combenef.CreateParameter("aux1", adVarChar, adParamInput, 3)
            .Parameters.Append combenef.CreateParameter("aux2", adVarChar, adParamInput, 3)
            .Parameters.Append combenef.CreateParameter("aux3", adVarChar, adParamInput, 3)
            .Parameters.Append combenef.CreateParameter("SIBs", adDouble, adParamOutput)
            .Parameters.Append combenef.CreateParameter("SISus", adDouble, adParamOutput)
            .Parameters("FFInicio") = Me.DTPinicio.Value
            .Parameters("FFFinal") = Me.DTPfin.Value
            .Parameters("cuenta") = Trim(Me.cbocta.Text)
            .Parameters("subcta1") = Trim(Me.cbosubcta1.Text)
            .Parameters("subcta2") = Trim(Me.cbosubcta2.Text)
            .Parameters("beneficiario") = Trim(Me.txtbusca1.Text)
            .Parameters("aux1") = Trim(Me.txtax1)
            .Parameters("aux2") = "00"
            .Parameters("aux3") = "00"
            .ActiveConnection = db
            .Execute
            SaldoIBs = .Parameters("SIBs")
            SaldoISus = .Parameters("SISus")
        End With
        
        'Me.ProgressBar1.Visible = True
        'Me.ProgressBar1.Value = 0
            CryLMayorBenef.Destination = crptToWindow
            CryLMayorBenef.ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMAuxBenef.rpt"
            CryLMayorBenef.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
            CryLMayorBenef.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
            CryLMayorBenef.StoredProcParam(2) = Trim(Me.cbocta.Text)
            CryLMayorBenef.StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
            CryLMayorBenef.StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
            CryLMayorBenef.StoredProcParam(5) = Trim(Me.txtbusca1)
            CryLMayorBenef.StoredProcParam(6) = Trim(Me.txtax1)
            CryLMayorBenef.StoredProcParam(7) = "00"
            CryLMayorBenef.StoredProcParam(8) = "00"
            
            CryLMayorBenef.Formulas(0) = "benef = '" & Trim(Me.txtbusca1) & "'"
            CryLMayorBenef.Formulas(1) = "cta = '" & Trim(Me.cbocta.Text) & "'"
            CryLMayorBenef.Formulas(2) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
            CryLMayorBenef.Formulas(3) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
            CryLMayorBenef.Formulas(4) = "nombenef = '" & nombenef & "'"
            CryLMayorBenef.Formulas(5) = "nomcta = '" & Trim(Me.Lblcuenta) & "'"
            CryLMayorBenef.Formulas(6) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
            CryLMayorBenef.Formulas(7) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
            CryLMayorBenef.Formulas(10) = "SIBs = " & SaldoIBs
            CryLMayorBenef.Formulas(11) = "SISus = " & SaldoISus
            CryLMayorBenef.Formulas(12) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
            CryLMayorBenef.Formulas(13) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
            IResult = CryLMayorBenef.PrintReport
     '*****fin aux1
    'Exit Sub
    ElseIf (Me.txtax1 = "02") Or (Me.Txtax2 = "02") Or (Me.txtax3 = "02") Then
        Set rsctabancaria = New ADODB.Recordset
        If rsctabancaria.State = 1 Then rsctabancaria.Close
        Dim SQLVar As String
        SQLVar = "SELECT fc_bancos.Bco_descripcion_larga,fc_cuenta_bancaria.Cta_codigo," & _
                 " fc_cuenta_bancaria.Cta_descripcion_larga FROM fc_bancos INNER JOIN " & _
                 " fc_cuenta_bancaria ON  fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo " & _
                 "WHERE fc_cuenta_bancaria.Cta_codigo='" & Trim(Me.cboCtaBancaria) & "'"
        rsctabancaria.Open SQLVar, db, adOpenKeyset, adLockReadOnly
        ctabancaria = Trim(rsctabancaria!cta_codigo)
        nombanco = Trim(rsctabancaria!Bco_descripcion_larga)
        nomctabancaria = Trim(rsctabancaria!cta_descripcion_larga)
        Set comctabancaria = New ADODB.Command
        With comctabancaria
            .CommandType = adCmdStoredProc
            .CommandText = "SaldoCtaBancaria"
            .Parameters.Append comctabancaria.CreateParameter("FFInicio", adVarChar, adParamInput, 10)
            .Parameters.Append comctabancaria.CreateParameter("FFFinal", adVarChar, adParamInput, 10)
            .Parameters.Append comctabancaria.CreateParameter("cuenta", adVarChar, adParamInput, 5)
            .Parameters.Append comctabancaria.CreateParameter("subcta1", adVarChar, adParamInput, 3)
            .Parameters.Append comctabancaria.CreateParameter("subcta2", adVarChar, adParamInput, 3)
            .Parameters.Append comctabancaria.CreateParameter("ctabancaria", adVarChar, adParamInput, 40)
            .Parameters.Append comctabancaria.CreateParameter("aux1", adVarChar, adParamInput, 3)
            .Parameters.Append comctabancaria.CreateParameter("aux2", adVarChar, adParamInput, 3)
            .Parameters.Append comctabancaria.CreateParameter("aux3", adVarChar, adParamInput, 3)
            .Parameters.Append comctabancaria.CreateParameter("SIBs", adDouble, adParamOutput)
            .Parameters.Append comctabancaria.CreateParameter("SISus", adDouble, adParamOutput)
            .Parameters("FFInicio") = Me.DTPinicio.Value
            .Parameters("FFFinal") = Me.DTPfin.Value
            .Parameters("cuenta") = Trim(Me.cbocta.Text)
            .Parameters("subcta1") = Trim(Me.cbosubcta1.Text)
            .Parameters("subcta2") = Trim(Me.cbosubcta2.Text)
            .Parameters("ctabancaria") = Trim(Me.cboCtaBancaria.Text)
            .Parameters("aux1") = Trim(Me.txtax1)
            .Parameters("aux2") = "00"
            .Parameters("aux3") = "00"
            .ActiveConnection = db
            .Execute
            SaldoIBs = .Parameters("SIBs")
            SaldoISus = .Parameters("SISus")
        End With
           CryLMayorCtaBancaria.Destination = crptToWindow
            CryLMayorCtaBancaria.ReportFileName = App.Path & "\REPORTES\Contabilidad\Libro_Mayor_Aux\CryLibroMAuxCta.rpt"
            CryLMayorCtaBancaria.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
            CryLMayorCtaBancaria.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
            CryLMayorCtaBancaria.StoredProcParam(2) = Trim(Me.cbocta.Text)
            CryLMayorCtaBancaria.StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
            CryLMayorCtaBancaria.StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
            CryLMayorCtaBancaria.StoredProcParam(5) = Trim(Me.cboCtaBancaria)
            CryLMayorCtaBancaria.StoredProcParam(6) = Trim(Me.txtax1)
            CryLMayorCtaBancaria.StoredProcParam(7) = "00"
            CryLMayorCtaBancaria.StoredProcParam(8) = "00"
            
            CryLMayorCtaBancaria.Formulas(0) = "cta = '" & Trim(Me.cbocta.Text) & "'"
            CryLMayorCtaBancaria.Formulas(1) = "ctabanco = '" & Trim(Me.cboCtaBancaria) & "'"
            CryLMayorCtaBancaria.Formulas(2) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
            CryLMayorCtaBancaria.Formulas(3) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
            CryLMayorCtaBancaria.Formulas(4) = "nombanco = '" & nombanco & "'"
            CryLMayorCtaBancaria.Formulas(5) = "nomcta = '" & Trim(Me.Lblcuenta) & "'"
            CryLMayorCtaBancaria.Formulas(6) = "nomctaBancaria = '" & nomctabancaria & "'"
            CryLMayorCtaBancaria.Formulas(7) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
            CryLMayorCtaBancaria.Formulas(8) = "nomsubcta2 = '" & Trim(Me.lbsub2) & "'"
            CryLMayorCtaBancaria.Formulas(11) = "SIBs = " & Val(SaldoIBs)
            CryLMayorCtaBancaria.Formulas(12) = "SISus= " & Val(SaldoISus)
            CryLMayorCtaBancaria.Formulas(13) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
            CryLMayorCtaBancaria.Formulas(14) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
            IResult = CryLMayorCtaBancaria.PrintReport
    End If
        If IResult <> 0 Then
            MsgBox CryLMayorBenef.LastErrorNumber & " : " & CryLMayorBenef.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
    '****fin aux 2
End If

End Sub

Private Sub CmdAceptar23_Click()

On Error GoTo gaby
If Me.txtax1 = "00" And Me.Txtax2 = "00" And Me.txtax3 = "00" Then
With DtEreportes
  
  If .rsauxiliar.State = adStateOpen Then
  .rsauxiliar.Close
  End If
   .Connection1.Execute "delete * from auxiliar where aux1='00' or aux1='01' or aux1='02'"
   .rsauxiliar.Open
   parametro = ""
   denominacion = ""
  
   If .rsauxiliar.State = adStateOpen Then
     .rsauxiliar.Close
   End If
    .rsauxiliar.Open
    .rsauxiliar.AddNew
    .rsauxiliar!cuenta = Trim(Me.cbocta.Text)
    .rsauxiliar!Subcta1 = Trim(Me.cbosubcta1.Text)
    .rsauxiliar!Subcta2 = Trim(Me.cbosubcta2.Text)
    .rsauxiliar!aux1 = Trim(Me.txtax1)
    .rsauxiliar!aux2 = Trim(Me.Txtax2)
    .rsauxiliar!aux3 = Trim(Me.txtax3)
    .rsauxiliar!nombre_cta = Me.Lblcuenta
    .rsauxiliar!Nombre_subcta1 = Me.Lblsub1
    .rsauxiliar!nombre_subcta2 = Me.lbsub2
    .rsauxiliar!Codigo_Beneficiario = parametro
    .rsauxiliar!denominacion_beneficiario = denominacion
    .rsauxiliar!fecha_Desde = Me.DTPinicio.Value
    .rsauxiliar!Fecha_A = Me.DTPfin.Value
    .rsauxiliar!saldo_inicialBs = 0
    .rsauxiliar!saldo_inicialSus = 0
    auxsaldobs = 0
    auxsaldosus = 0
    .rsauxiliar!cta_codigo = " "
    .rsauxiliar!cta_nombre = " "
    .rsauxiliar!Nombre_banco = ""
     saldobs1 = 0
     saldosus1 = 0
     saldobs = 0
     saldosus = 0
    .rsauxiliar!saldo_inicialBs = 0
    .rsauxiliar!saldo_inicialSus = 0
    ax_iniciobs = saldobs1
  .rsauxiliar.Update
     'Form3.Show
  If .rsdiario.State = adStateOpen Then
    .rsdiario.Close
  End If
  '.rsdiario.Open
  If .rsdiario_000.State = 1 Then
  .rsdiario_000.Close
  End If
  .rsdiario_000.Open
   'MsgBox .rsdiario_000.RecordCount
       'Form3.Show
  If .rsdiario_000.RecordCount = 0 Then
    MsgBox "La cuenta " & parametro & " no ha tenido ningun movimiento", vbInformation + vbDefaultButton1, "SAF2000"
    Me.cbocta.SetFocus
    .rsdiario_000.Close
    Exit Sub
  End If
   .rsdiario_000.MoveFirst
  '*******abrimo info del diario
  'If .rsaux_diario.State = adStateOpen Then
  ' .rsdiario_000.Close
  'End If
  
 '.rsaux_diario.Open
  '* borramos auxiliar del diario
    kj = .rsaux_diario.RecordCount
  Me.PRB.Visible = True
  If kj <> 0 Then
    Me.PRB.Max = kj
  Me.PRB.Value = 0
  End If
  If .rsaux_diario.RecordCount <> 0 Then
  .rsaux_diario.MoveFirst
  For i = 1 To .rsaux_diario.RecordCount
    .rsaux_diario.Delete
    .rsaux_diario.MoveNext
    Me.PRB.Value = Me.PRB.Value + 1
  Next i
  End If
  'MsgBox .rsaux_diario.RecordCount
  '.rsreporte_diario.Open
  .Connection1.Execute "delete * from reporte_diario  where  tipo_comprobante ='PAC' or tipo_comprobante='PCO' or tipo_comprobante='PCE' or tipo_comprobante='DAC'"
  
  If .rsreporte_diario.State = adStateOpen Then
   .rsreporte_diario.Close
  End If
  .rsreporte_diario.Open
  'MsgBox .rsreporte_diario.RecordCount
  
  ' query del diario
  u = 1
  '***************trabajamos con tabla auxiliar
  .rsdiario.Open
'  .rsdiario_000.Open
  aaa = .rsdiario_000.RecordCount
  Me.PRB.Max = aaa
  Me.PRB.Value = 0
  '.rsdiario.Close
  Do While Not .rsdiario_000.EOF
  .rsaux_diario.AddNew
  '***** si el movimiento de la cuenta es en el debe
  '**** si las fechas estan dentro del rango
    If (.rsdiario_000!D_Cuenta = .rsauxiliar!cuenta And .rsdiario_000!D_SubCta1 = .rsauxiliar!Subcta1 And .rsdiario_000!D_SubCta2 = .rsauxiliar!Subcta2 And .rsdiario_000!d_Aux1 = .rsauxiliar!aux1 And .rsdiario_000!d_Aux2 = .rsauxiliar!aux2 And .rsdiario_000!d_Aux3 = .rsauxiliar!aux3) Then
      .rsaux_diario!no_comprobante = .rsdiario_000!Cod_Comp
      .rsaux_diario!tipo_comprobante = .rsdiario_000!tipo_Comp
      .rsaux_diario!debe_bs = .rsdiario_000!D_MontoBs
      .rsaux_diario!debe_dl = .rsdiario_000!D_MontoDl
      .rsaux_diario!haber_bs = 0
      .rsaux_diario!haber_dl = 0
      .rsaux_diario!Movim_sus = .rsdiario_000!D_MontoDl
      .rsaux_diario!saldo_bs = 0
      .rsaux_diario!saldo_sus = 0
      .rsaux_diario!Glosa = .rsdiario_000!Glosa
      .rsaux_diario!Fecha_A = .rsdiario_000!Fecha_A
      .rsaux_diario!D_Cambio = .rsdiario_000!D_Cambio
    End If
  '*si el movimiento de la cuenta es en el haber
    If (.rsdiario_000!H_Cuenta = .rsauxiliar!cuenta And .rsdiario_000!H_SubCta1 = .rsauxiliar!Subcta1 And .rsdiario_000!H_SubCta2 = .rsauxiliar!Subcta2 And .rsdiario_000!h_Aux1 = .rsauxiliar!aux1 And .rsdiario_000!h_Aux2 = .rsauxiliar!aux2 And .rsdiario_000!h_Aux3 = .rsauxiliar!aux3) Then
      .rsaux_diario!no_comprobante = .rsdiario_000!Cod_Comp
      .rsaux_diario!tipo_comprobante = .rsdiario_000!tipo_Comp
      .rsaux_diario!debe_bs = 0
      .rsaux_diario!debe_dl = 0
      .rsaux_diario!haber_bs = .rsdiario_000!h_MontoBs
      .rsaux_diario!haber_dl = .rsdiario_000!h_MontoDl
      .rsaux_diario!Movim_sus = -.rsdiario_000!h_MontoDl
      .rsaux_diario!saldo_bs = 0
      .rsaux_diario!saldo_sus = 0
      .rsaux_diario!Glosa = .rsdiario_000!Glosa
      .rsaux_diario!Fecha_A = .rsdiario_000!Fecha_A
      .rsaux_diario!D_Cambio = .rsdiario_000!D_Cambio
    End If
    .rsaux_diario.Update
    .rsdiario_000.MoveNext
    u = u + 1
    Me.PRB.Value = Me.PRB.Value + 1
   
  Loop
  '******************
  hy = .rsaux_diario.RecordCount
  Me.PRB.Max = hy
  Me.PRB.Value = 0
  .rsaux_diario.MoveFirst
  Do While Not .rsaux_diario.EOF
    If .rsaux_diario!debe_bs <> 0 Then
      .rsaux_diario!saldo_bs = saldobs + .rsaux_diario!debe_bs
       saldobs = .rsaux_diario!saldo_bs
      .rsaux_diario!saldo_sus = saldosus + .rsaux_diario!debe_dl
       saldosus = .rsaux_diario!saldo_sus
    End If
    If .rsaux_diario!haber_bs <> 0 Then
      .rsaux_diario!saldo_bs = saldobs - .rsaux_diario!haber_bs
      saldobs = .rsaux_diario!saldo_bs
      .rsaux_diario!saldo_sus = saldosus - .rsaux_diario!haber_dl
       saldosus = .rsaux_diario!saldo_sus
    End If
    .rsaux_diario.Update
    .rsaux_diario.MoveNext
    Me.PRB.Value = Me.PRB.Value + 1
  Loop
  .rsaux_diario.MoveFirst
  tr = .rsaux_diario.RecordCount
  Me.PRB.Max = tr
  Me.PRB.Value = 0
  Do While Not .rsaux_diario.EOF
    If (.rsaux_diario!Fecha_A >= Me.DTPinicio.Value) And (.rsaux_diario!Fecha_A <= Me.DTPfin.Value) Then
    .rsreporte_diario.AddNew
    .rsreporte_diario!no_comprobante = .rsaux_diario!no_comprobante
    .rsreporte_diario!tipo_comprobante = .rsaux_diario!tipo_comprobante
    .rsreporte_diario!debe_bs = .rsaux_diario!debe_bs
    .rsreporte_diario!debe_dl = .rsaux_diario!debe_dl
    .rsreporte_diario!haber_bs = .rsaux_diario!haber_bs
    .rsreporte_diario!haber_dl = .rsaux_diario!haber_dl
    .rsreporte_diario!Movim_sus = .rsaux_diario!Movim_sus
    .rsreporte_diario!saldo_bs = .rsaux_diario!saldo_bs
    .rsreporte_diario!saldo_sus = .rsaux_diario!saldo_sus
    .rsreporte_diario!Glosa = .rsaux_diario!Glosa
    .rsreporte_diario!Fecha_A = .rsaux_diario!Fecha_A
    .rsreporte_diario!D_Cambio = .rsaux_diario!D_Cambio
    .rsreporte_diario.Update
  Else
    If (.rsaux_diario!Fecha_A < Me.DTPinicio.Value) Then
  '  saldobs1 = saldobs1 + .rsaux_diario!saldo_bs
  '  saldosus1 = saldosus1 + .rsaux_diario!saldo_sus
      saldobs1 = saldobs1 + .rsaux_diario!debe_bs - .rsaux_diario!haber_bs
      saldosus1 = saldosus1 + .rsaux_diario!debe_dl - .rsaux_diario!haber_dl
    End If
  End If
    .rsaux_diario.MoveNext
    Me.PRB.Value = Me.PRB.Value + 1
  Loop
  '*********
  auxsaldobs = saldobs1
 ' MsgBox auxsaldobs
 ' MsgBox auxsaldosus
  auxsaldosus = saldosus1
  'MsgBox .rsreporte_diario.RecordCount
  If .rsreporte_diario.RecordCount = 0 Then
     MsgBox "No existen registros en esas fechas", vbInformation + vbDefaultButton1
     Me.cbocta.SetFocus
     Exit Sub
  End If
  '***************
  .rsreporte_diario.MoveFirst
  ju = .rsreporte_diario.RecordCount
  Me.PRB.Max = ju
  Me.PRB.Value = 0
    Do While Not .rsreporte_diario.EOF
      If .rsreporte_diario!debe_bs <> 0 Then
        .rsreporte_diario!saldo_bs = saldobs1 + .rsreporte_diario!debe_bs
         saldobs1 = .rsreporte_diario!saldo_bs
        .rsreporte_diario!saldo_sus = saldosus1 + .rsreporte_diario!debe_dl
         saldosus1 = .rsreporte_diario!saldo_sus
      End If
      If .rsreporte_diario!haber_bs <> 0 Then
        .rsreporte_diario!saldo_bs = saldobs1 - .rsreporte_diario!haber_bs
        saldobs1 = .rsreporte_diario!saldo_bs
        .rsreporte_diario!saldo_sus = saldosus1 - .rsreporte_diario!haber_dl
         saldosus1 = .rsreporte_diario!saldo_sus
      End If
      .rsreporte_diario.Update
      .rsreporte_diario.MoveNext
      Me.PRB.Value = Me.PRB.Value + 1
    Loop
      .rsauxiliar!saldo_inicialBs = auxsaldobs
      .rsauxiliar!saldo_inicialSus = auxsaldosus
      .rsauxiliar.Update
  '.rsauxiliar.Open
  '********
    Me.PRB.Visible = False
    Frmrepmayor.Show
  End With
'***********caso con auxiliares
Else
  Set reporte = New ADODB.Recordset
  With DtEreportes
  fech = CDate(Date)
  If .rsaux_reportes.State = 1 Then .rsaux_reportes.Close
  '.Connection1.Execute "delete * from aux_reportes where fecha <>" & fech & ""
  .rsaux_reportes.Open
  If .rsaux_reportes.RecordCount <> 0 Then
    .rsaux_reportes.MoveFirst
    Do While Not .rsaux_reportes.EOF
      If .rsaux_reportes!FECHA <> Date Then
        .rsaux_reportes.Delete
      End If
    .rsaux_reportes.MoveNext
    Loop
   ' MsgBox .rsaux_reportes.RecordCount
  End If
  If .rsauxiliar.State = adStateOpen Then
  .rsauxiliar.Close
  End If
   .Connection1.Execute "delete * from auxiliar where aux1='00' or aux1='01' or aux1='02'"
   .rsauxiliar.Open
   'MsgBox .rsauxiliar.RecordCount
   If BUSCA = 1 Then
    parametro = .rsbenef!Codigo_Beneficiario
    denominacion = .rsbenef!denominacion_beneficiario
   End If
   If BUSCA = 2 Then
    parametro = .rscta_banco!no_cuenta
    denominacion = .rscta_banco!nombre_cta
   End If
   If .rsauxiliar.State = adStateOpen Then
     .rsauxiliar.Close
   End If
    .rsauxiliar.Open
    .rsauxiliar.AddNew
    .rsauxiliar!cuenta = Trim(Me.cbocta.Text)
    .rsauxiliar!Subcta1 = Trim(Me.cbosubcta1.Text)
    .rsauxiliar!Subcta2 = Trim(Me.cbosubcta2.Text)
    .rsauxiliar!aux1 = Trim(Me.txtax1)
    .rsauxiliar!aux2 = Trim(Me.Txtax2)
    .rsauxiliar!aux3 = Trim(Me.txtax3)
    .rsauxiliar!nombre_cta = Me.Lblcuenta
    .rsauxiliar!Nombre_subcta1 = Me.Lblsub1
    .rsauxiliar!nombre_subcta2 = Me.lbsub2
    .rsauxiliar!Codigo_Beneficiario = parametro
    .rsauxiliar!denominacion_beneficiario = denominacion
    .rsauxiliar!fecha_Desde = Me.DTPinicio.Value
    .rsauxiliar!Fecha_A = Me.DTPfin.Value
  'If Me.Chkaux1.Value = 1 Then
    If (Me.txtax1 = "01" And Me.Chkaux1.Value = 1) Or (Me.Txtax2 = "01" And Me.Chkaux2.Value = 1) Or (Me.txtax3 = "01" And Me.Chkaux3.Value = 1) Then
    '****saldos iniciales con beneficiarios=0
      .rsauxiliar!saldo_inicialBs = 0
      .rsauxiliar!saldo_inicialSus = 0
      auxsaldobs = 0
      auxsaldosus = 0
      .rsauxiliar!cta_codigo = " "
      .rsauxiliar!cta_nombre = " "
      '.rsauxiliar!Nombre_banco = ""
       'rsauxiliar!Nombre_banco = " "
       .rsauxiliar!Nombre_banco = ""
       saldobs1 = 0
       saldosus1 = 0
       saldobs = 0
       saldosus = 0
    End If
  '  If Me.Chkaux1.Value = 2 Then
    '****saldos iniciales con bancos
    If (Me.txtax1 = "02" And Me.Chkaux1.Value = 1) Or (Me.Txtax2 = "02" And Me.Chkaux2.Value = 1) Or (Me.txtax3 = "02" And Me.Chkaux3.Value = 1) Then
      .rsauxiliar!cta_codigo = .rscta_banco!no_cuenta
      .rsauxiliar!cta_nombre = .rscta_banco!nombre_cta
      .rsauxiliar!Nombre_banco = .rscta_banco!banco
       'rsauxiliar!Nombre_banco = .rscta_banco!banco
        .rsauxiliar!saldo_inicialBs = .rscta_banco!cta_saldo_inicial
        ' TIPO DE CAMBIO  DE VENTA AL 3 DE ENERO 6 , SE TRABAJA CON MENOS 2 PUNTOS
        .rsauxiliar!saldo_inicialSus = .rscta_banco!cta_saldo_inicial / 5.98
         saldobs1 = .rscta_banco!cta_saldo_inicial
         saldosus1 = .rscta_banco!cta_saldo_inicial / 5.98
         saldobs = .rscta_banco!cta_saldo_inicial
         saldosus = .rscta_banco!cta_saldo_inicial / 5.98
         auxsaldobs = saldobs1
         auxsaldosus = saldosus1
         ax_iniciobs = saldobs1
         
    End If
  .rsauxiliar.Update
     'Form3.Show
  If .rsdiario.State = adStateOpen Then
    .rsdiario.Close
  End If
  .rsdiario.Open
  'MsgBox .rsdiario.RecordCount
    'Form3.Show
  If .rsdiario.RecordCount = 0 Then
    MsgBox "La cuenta " & parametro & " no ha tenido ningun movimiento", vbInformation + vbDefaultButton1, "SAF2000"
    If Me.Chkaux1.Value = 1 Then
      Me.txtbusca1.SetFocus
      .rsdiario.Close
      Exit Sub
    End If
    If Me.Chkaux2.Value = 1 Then
      Me.Txtbusca2.SetFocus
      .rsdiario.Close
      Exit Sub
    End If
    If Me.Chkaux3.Value = 1 Then
      Me.Txtbusca3.SetFocus
      .rsdiario.Close
      Exit Sub
    End If
  End If
  .rsdiario.MoveFirst
  '*******abrimo info del diario
  If .rsaux_diario.State = adStateOpen Then
   .rsaux_diario.Close
  End If
  .rsaux_diario.Open
  '* borramos auxiliar del diario
  ccc = .rsaux_diario.RecordCount
  Me.PRB.Visible = True
  If ccc <> 0 Then
  Me.PRB.Max = ccc
  Me.PRB.Value = 0
  End If
  For i = 1 To .rsaux_diario.RecordCount
    .rsaux_diario.Delete
    .rsaux_diario.MoveNext
  Me.PRB.Value = Me.PRB.Value + 1
  Next i
  'MsgBox .rsaux_diario.RecordCount
  '.rsreporte_diario.Open
  .Connection1.Execute "delete * from reporte_diario  where  tipo_comprobante ='PAC' or tipo_comprobante='PCO' or tipo_comprobante='PCE' or tipo_comprobante='DAC'"
  'MsgBox .rsreporte_diario.RecordCount
  'If .rsreporte_diario.State = adStateOpen Then
   ' .rsreporte_diario.Close
  'End If
  '.rsreporte_diario.Open
  '.rsreporte_diario.MoveFirst
  'For j = 1 To .rsreporte_diario.RecordCount
   ' .rsreporte_diario.Delete
   ' .rsreporte_diario.MoveNext
  'Next j
  If .rsreporte_diario.State = adStateOpen Then
   .rsreporte_diario.Close
  End If
  .rsreporte_diario.Open
  'MsgBox .rsreporte_diario.RecordCount
  
  ' query del diario
  
  '***************trabajamos con tabla auxiliar
  tt = .rsdiario.RecordCount
  Me.PRB.Max = tt
  Me.PRB.Value = 0
  Do While Not .rsdiario.EOF
  .rsaux_diario.AddNew
  '***** si el movimiento de la cuenta es en el debe
  '**** si las fechas estan dentro del rango
    If (.rsdiario!D_Cuenta = .rsauxiliar!cuenta And .rsdiario!D_SubCta1 = .rsauxiliar!Subcta1 And .rsdiario!D_SubCta2 = .rsauxiliar!Subcta2 And .rsdiario!d_Aux1 = .rsauxiliar!aux1 And .rsdiario!d_Aux2 = .rsauxiliar!aux2 And .rsdiario!d_Aux3 = .rsauxiliar!aux3) Then
      .rsaux_diario!no_comprobante = .rsdiario!Cod_Comp
      .rsaux_diario!tipo_comprobante = .rsdiario!tipo_Comp
      .rsaux_diario!debe_bs = .rsdiario!D_MontoBs
      .rsaux_diario!debe_dl = .rsdiario!D_MontoDl
      .rsaux_diario!haber_bs = 0
      .rsaux_diario!haber_dl = 0
      .rsaux_diario!Movim_sus = .rsdiario!D_MontoDl
      .rsaux_diario!saldo_bs = 0
      .rsaux_diario!saldo_sus = 0
      .rsaux_diario!Glosa = .rsdiario!Glosa
      .rsaux_diario!Fecha_A = .rsdiario!Fecha_A
      .rsaux_diario!D_Cambio = .rsdiario!D_Cambio
    End If
  '*si el movimiento de la cuenta es en el haber
    If (.rsdiario!H_Cuenta = .rsauxiliar!cuenta And .rsdiario!H_SubCta1 = .rsauxiliar!Subcta1 And .rsdiario!H_SubCta2 = .rsauxiliar!Subcta2 And .rsdiario!h_Aux1 = .rsauxiliar!aux1 And .rsdiario!h_Aux2 = .rsauxiliar!aux2 And .rsdiario!h_Aux3 = .rsauxiliar!aux3) Then
      .rsaux_diario!no_comprobante = .rsdiario!Cod_Comp
      .rsaux_diario!tipo_comprobante = .rsdiario!tipo_Comp
      .rsaux_diario!debe_bs = 0
      .rsaux_diario!debe_dl = 0
      .rsaux_diario!haber_bs = .rsdiario!h_MontoBs
      .rsaux_diario!haber_dl = .rsdiario!h_MontoDl
      .rsaux_diario!Movim_sus = -.rsdiario!h_MontoDl
      .rsaux_diario!saldo_bs = 0
      .rsaux_diario!saldo_sus = 0
      .rsaux_diario!Glosa = .rsdiario!Glosa
      .rsaux_diario!Fecha_A = .rsdiario!Fecha_A
      .rsaux_diario!D_Cambio = .rsdiario!D_Cambio
    End If
    .rsaux_diario.Update
    .rsdiario.MoveNext
  Me.PRB.Value = Me.PRB.Value + 1
  Loop
  '******************
  .rsaux_diario.MoveFirst
  gg = .rsaux_diario.RecordCount
  Me.PRB.Max = gg
  Me.PRB.Value = 0
  Do While Not .rsaux_diario.EOF
    If .rsaux_diario!debe_bs <> 0 Then
      .rsaux_diario!saldo_bs = saldobs + .rsaux_diario!debe_bs
       saldobs = .rsaux_diario!saldo_bs
      .rsaux_diario!saldo_sus = saldosus + .rsaux_diario!debe_dl
       saldosus = .rsaux_diario!saldo_sus
    End If
    If .rsaux_diario!haber_bs <> 0 Then
      .rsaux_diario!saldo_bs = saldobs - .rsaux_diario!haber_bs
      saldobs = .rsaux_diario!saldo_bs
      .rsaux_diario!saldo_sus = saldosus - .rsaux_diario!haber_dl
       saldosus = .rsaux_diario!saldo_sus
    End If
    .rsaux_diario.Update
    .rsaux_diario.MoveNext
  Me.PRB.Value = Me.PRB.Value + 1
  Loop
  .rsaux_diario.MoveFirst
  tgg = .rsaux_diario.RecordCount
  Me.PRB.Max = tgg
  Me.PRB.Value = 0
  Do While Not .rsaux_diario.EOF
  If (.rsaux_diario!Fecha_A >= Me.DTPinicio.Value) And (.rsaux_diario!Fecha_A <= Me.DTPfin.Value) Then
    .rsreporte_diario.AddNew
    .rsreporte_diario!no_comprobante = .rsaux_diario!no_comprobante
    .rsreporte_diario!tipo_comprobante = .rsaux_diario!tipo_comprobante
    .rsreporte_diario!debe_bs = .rsaux_diario!debe_bs
    .rsreporte_diario!debe_dl = .rsaux_diario!debe_dl
    .rsreporte_diario!haber_bs = .rsaux_diario!haber_bs
    .rsreporte_diario!haber_dl = .rsaux_diario!haber_dl
    .rsreporte_diario!Movim_sus = .rsaux_diario!Movim_sus
    .rsreporte_diario!saldo_bs = .rsaux_diario!saldo_bs
    .rsreporte_diario!saldo_sus = .rsaux_diario!saldo_sus
    .rsreporte_diario!Glosa = .rsaux_diario!Glosa
    .rsreporte_diario!Fecha_A = .rsaux_diario!Fecha_A
    .rsreporte_diario!D_Cambio = .rsaux_diario!D_Cambio
    .rsreporte_diario.Update
  Else
    If (.rsaux_diario!Fecha_A < Me.DTPinicio.Value) Then
  '  saldobs1 = saldobs1 + .rsaux_diario!saldo_bs
  '  saldosus1 = saldosus1 + .rsaux_diario!saldo_sus
      saldobs1 = saldobs1 + .rsaux_diario!debe_bs - .rsaux_diario!haber_bs
      saldosus1 = saldosus1 + .rsaux_diario!debe_dl - .rsaux_diario!haber_dl
    End If
  End If
    .rsaux_diario.MoveNext
    Me.PRB.Value = Me.PRB.Value + 1
  Loop
  '*********
  auxsaldobs = saldobs1
  'MsgBox auxsaldobs
  'MsgBox auxsaldosus
  auxsaldosus = saldosus1
  'MsgBox .rsreporte_diario.RecordCount
  If .rsreporte_diario.RecordCount = 0 Then
     MsgBox "No existen registros en esas fechas", vbInformation + vbDefaultButton1
     If Me.Chkaux1.Value = 1 Then
      Me.txtbusca1.SetFocus
     End If
     If Me.Chkaux2.Value = 1 Then
      Me.Txtbusca2.SetFocus
     End If
     If Me.Chkaux3.Value = 1 Then
      Me.Txtbusca3.SetFocus
     End If
   Exit Sub
  End If
  '***************
  If .rsaux_reportes.State = 1 Then
    .rsaux_reportes.Close
  End If
  '.Connection1.Execute "select * from aux_reportes where cuenta=" & Me.cbocta.Text & "and  subcta1=" & Me.cbosubcta1.Text & "and subcta2=" & Me.cbosubcta2.Text & "and fecha=" & Date & "and asunto=" & Trim(Me.txtbusca1.Text) & ""
  If Me.txtax1 = "00" And Me.Txtax2 = "00" And Me.txtax3 = "00" Then
      .rsauxiliar.MoveFirst
      .rsauxiliar!saldo_inicialBs = 0
      .rsauxiliar!saldo_inicialSus = 0
      .rsauxiliar.Update
      saldobs1 = 0
      saldosus1 = 0
  Else
  'If .rsaux_reportes.State = 1 Then
'  .rsaux_reportes.Close
  'End If
   '.rsaux_reportes.Open
'  vw = 0
'  Do While Not .rsaux_reportes.EOF
'  If .rsaux_reportes!cuenta = Me.cbocta.Text And .rsaux_reportes!subcta1 = Me.cbosubcta1.Text And .rsaux_reportes!subcta2 = Me.cbosubcta2.Text And .rsaux_reportes!asunto = Me.txtbusca1.Text And .rsaux_reportes!fecha = Date Then
'
'  vw = 1
'  End If
'  .rsaux_reportes.MoveNext
'  Loop
  .rsaux_reportes.Open
  If .rsaux_reportes.State = 1 Then
  .rsaux_reportes.Close
  End If
  Set existereporte = New ADODB.Recordset
  
  If existereporte.State = 1 Then existereporte.Close
  existereporte.Open "select * from aux_reportes where cuenta = '" & Trim(Me.cbocta.Text) & "' and  subcta1 = '" & Trim(Me.cbosubcta1.Text) & "' and  asunto = '" & Trim(Me.txtbusca1.Text) & "' ", db, adOpenKeyset, adLockOptimistic 'fecha = #" & Date & "# and
  '.rsaux_reportes.Open " select * from aux_reportes where cuenta=" & Me.cbocta.Text & " and  subcta1=" & Me.cbosubcta1.Text & " and fecha=" & Date & " and asunto=" & (Me.txtbusca1.Text) & ""
  '.Connection1.Execute "select * from aux_reportes where cuenta=" & Me.cbocta.Text & "and  subcta1=" & Me.cbosubcta1.Text & "and fecha=" & Date & "and asunto=" & (Me.txtbusca1.Text) & ""
  'MsgBox .rsaux_reportes.RecordCount
  .rsaux_reportes.Open
  'MsgBox existereporte.RecordCount
  If existereporte.EOF And existereporte.BOF Then
  'MsgBox "vacio"
  End If
  'MsgBox .rsaux_reportes.RecordCount
  If existereporte.RecordCount = 0 Then
  'If .rsaux_reportes.RecordCount = 0 Then
    .rsaux_reportes.AddNew
      .rsaux_reportes!cuenta = .rsauxiliar!cuenta
      .rsaux_reportes!Subcta1 = .rsauxiliar!Subcta1
      .rsaux_reportes!Subcta2 = .rsauxiliar!Subcta2
      .rsaux_reportes!asunto = .rsauxiliar!cta_codigo
      .rsaux_reportes!FECHA = Date
    .rsaux_reportes.Update
  valor = 1
  Else
    existereporte.MoveFirst
    '.rsaux_reportes.MoveFirst
    'Do While Not .rsaux_reportes.EOF
    Do While Not existereporte.EOF
    'If .rsaux_reportes!subcta2 <> Trim(Me.cbosubcta2.Text) Then
     If existereporte!Subcta2 <> Trim(Me.cbosubcta2.Text) Then
      valor = 2
      .rsauxiliar!saldo_inicialBs = 0
      .rsauxiliar!saldo_inicialSus = 0
      .rsauxiliar.Update
      saldobs1 = 0
      saldosus1 = 0
    Else
    valor = 1
    End If
    '.rsaux_reportes.MoveNext
    existereporte.MoveNext
    Loop
  End If
  End If
    .rsreporte_diario.MoveFirst
    hh = .rsreporte_diario.RecordCount
    Me.PRB.Max = hh
    Me.PRB.Value = 0
    Do While Not .rsreporte_diario.EOF
      If .rsreporte_diario!debe_bs <> 0 Then
        .rsreporte_diario!saldo_bs = saldobs1 + .rsreporte_diario!debe_bs
         saldobs1 = .rsreporte_diario!saldo_bs
        .rsreporte_diario!saldo_sus = saldosus1 + .rsreporte_diario!debe_dl
         saldosus1 = .rsreporte_diario!saldo_sus
      End If
      If .rsreporte_diario!haber_bs <> 0 Then
        .rsreporte_diario!saldo_bs = saldobs1 - .rsreporte_diario!haber_bs
        saldobs1 = .rsreporte_diario!saldo_bs
        .rsreporte_diario!saldo_sus = saldosus1 - .rsreporte_diario!haber_dl
         saldosus1 = .rsreporte_diario!saldo_sus
      End If
      .rsreporte_diario.Update
      .rsreporte_diario.MoveNext
    Me.PRB.Value = Me.PRB.Value + 1
    Loop
    If valor = 1 Then
      .rsauxiliar!saldo_inicialBs = auxsaldobs
      .rsauxiliar!saldo_inicialSus = auxsaldosus
      .rsauxiliar.Update
    Else
      .rsauxiliar!saldo_inicialBs = 0
      .rsauxiliar!saldo_inicialSus = 0
      .rsauxiliar.Update
    End If
  
    If BUSCA = 2 Then
      .rsauxiliar!Codigo_Beneficiario = " "
      .rsauxiliar!denominacion_beneficiario = " "
      .rsauxiliar.Update
    End If
  
  '.rsauxiliar.Open
  '********
Me.PRB.Visible = False
  Frmrepmayor.Show
  End With
End If
gaby:
If Err.Number = 3021 Then
  MsgBox "No se encontró el auxiliar. Revise el Grid para encontrarlo", vbCritical + vbDefaultButton1
  Me.cbocta.SetFocus
Exit Sub
End If
End Sub

Private Sub CmdBusca_Click()
Me.Fra_Busqueda.Visible = True
End Sub



Private Sub CmdBuscar_Click()
Me.Fra_Busqueda.Visible = True
End Sub

Private Sub CmdCancelar_Click()
  Me.txtbusca1 = ""
  Me.Txtbusca2 = ""
  Me.Txtbusca3 = ""
  Me.cbocta.SetFocus
  'Me.txtaux = ""
  'Me.cboaux.Text = Me.cboaux.List(0)
End Sub

Private Sub CmdEjecutar_Click()
Select Case Me.CboCampo
    Case "codigo_beneficiario"
        Select Case Me.CboOperador
            Case "="
                 sql2 = "SELECT codigo_beneficiario, denominacion_beneficiario From fc_beneficiario  where  codigo_beneficiario =' " & Trim(Me.TxtValor) & "' order by codigo_beneficiario"
            Case "como"
                 sql2 = " select codigo_beneficiario, denominacion_beneficiario from  fc_beneficiario WHERE Codigo_beneficiario like '" & Trim(Me.TxtValor) & "'+'%' order by codigo_beneficiario"
        End Select
    Case "denominacion_beneficiario"
        Select Case Me.CboOperador
            Case "="
                sql2 = "SELECT codigo_beneficiario, denominacion_beneficiario From fc_beneficiario  where  denominacion_beneficiario =' " & Trim(Me.TxtValor) & "' order by denominacion_beneficiario"
        Case "como"
                sql2 = " select codigo_beneficiario, denominacion_beneficiario from  fc_beneficiario WHERE denominacion_beneficiario like '" & Trim(Me.TxtValor) & "'+'%'  order by denominacion_beneficiario"
    End Select
End Select
    If rsBeneficiario.State = 1 Then rsBeneficiario.Close
    rsBeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
    Set Me.DtGbenef.DataSource = rsBeneficiario
End Sub

Private Sub CmdSalir_Click()
'Dtereportes.Connection1.Close
Unload Me
End Sub

Private Sub DataGrid1_LostFocus()
parametro = DtEreportes.rsbenef!Codigo_Beneficiario
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DTGBanco_Click()
'Me.txtbusca1.Text = Me.DTGBanco.Columns(0).Value
   On Error GoTo error3
    Me.cboCtaBancaria.Text = Me.DTGBanco.Columns(0).Value
error3:
    If Err.Number = 7005 Then
        MsgBox "No existen datos", vbCritical + vbDefaultButton1
        Exit Sub
    End If
    
End Sub

Private Sub DtGbenef_Click()
On Error GoTo err1
Me.txtbusca1.Text = Me.DtGbenef.Columns(0)
err1:
If Err.Number = 7005 Then
DtGbenef.Refresh
End If

End Sub

Private Sub DTPfin_Validate(Cancel As Boolean)
If DTPfin.Value < DTPinicio.Value Then
    MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
    DTPfin.SetFocus
End If
End Sub

Private Sub DTPinicio_LostFocus()
Me.DTPfin.MinDate = Me.DTPinicio.Value
End Sub
Private Sub DTPinicio_Validate(Cancel As Boolean)
If DTPinicio.Value > DTPfin.Value Then
    MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
    DTPfin.SetFocus
End If
End Sub

Private Sub Form_Load()
Me.CmdAceptar.Enabled = True
On Error GoTo error_conec
    Set rsplanctas = New ADODB.Recordset
    Set rscuentas = New ADODB.Recordset
    Set rsnombresub1 = New ADODB.Recordset
    Set rssubcuenta = New ADODB.Recordset
    Set rscta_bancaria = New ADODB.Recordset
    Set rsBeneficiario = New ADODB.Recordset
    If rsplanctas.State = 1 Then rsplanctas.Close
    rsplanctas.Open "SELECT Cuenta, NombreCta FROM CC_Plan_Cuentas WHERE SubCta1 = '00' AND SubCta2 = '00' order by Cuenta", db, adOpenKeyset, adLockReadOnly
    rsplanctas.MoveFirst
    Do While Not rsplanctas.EOF
        Me.cbocta.AddItem rsplanctas!cuenta
        rsplanctas.MoveNext
    Loop
    If rsBeneficiario.State = 1 Then rsBeneficiario.Close
    sql2 = "SELECT codigo_beneficiario, denominacion_beneficiario From fc_beneficiario order by denominacion_beneficiario"
    rsBeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
    Set Me.DtGbenef.DataSource = rsBeneficiario
    
    Me.cbocta.Text = Me.cbocta.List(0)
    Me.DTPfin.MaxDate = CDate(Date)
    Me.DTPinicio.MaxDate = CDate(Date)
    Me.DTPfin.Value = Date
    Me.DTPinicio.Value = CDate("01/01/2000")
    Me.DTPinicio.MinDate = CDate("01/01/2000")
    Me.DTPfin.MinDate = CDate(Me.DTPinicio.Value)
    Me.PRB.Visible = False
    
    Exit Sub
error_conec:
    If Err.Number = -2147220992 Then
      MsgBox "ERROR EN LA CONECCION, Revise su conección a la red", vbCritical + vbDefaultButton1, "SAF/2000"
      End
    End If

	Call SeguridadSet(Me)
End Sub
Public Sub Mayor000()
  Dim IResult As Integer
    Set commayor = New ADODB.Command ' para obtener los saldos
    With commayor
        .CommandType = adCmdStoredProc
        .CommandText = "SaldoLMayor"
        .Parameters.Append commayor.CreateParameter("FFInicio", adVarChar, adParamInput, 10)
        .Parameters.Append commayor.CreateParameter("FFFinal", adVarChar, adParamInput, 10)
        .Parameters.Append commayor.CreateParameter("cuenta", adVarChar, adParamInput, 5)
        .Parameters.Append commayor.CreateParameter("subcta1", adVarChar, adParamInput, 3)
        .Parameters.Append commayor.CreateParameter("subcta2", adVarChar, adParamInput, 3)
        .Parameters.Append commayor.CreateParameter("SIBs", adDouble, adParamOutput)
        .Parameters.Append commayor.CreateParameter("SISus", adDouble, adParamOutput)
        .Parameters("FFInicio") = Me.DTPinicio.Value
        .Parameters("FFFinal") = Me.DTPfin.Value
        .Parameters("cuenta") = Trim(Me.cbocta.Text)
        .Parameters("subcta1") = Trim(Me.cbosubcta1.Text)
        .Parameters("subcta2") = Trim(Me.cbosubcta2.Text)
        .ActiveConnection = db
        .Execute
        SaldoIBs = .Parameters("SIBs")
        SaldoISus = .Parameters("SISus")
    End With
        CryLMayor.Destination = crptToWindow
        CryLMayor.ReportFileName = "C:\REPORTES_SQL\Libro_Mayor\CryLMayor.rpt"
        CryLMayor.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
        CryLMayor.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
        CryLMayor.StoredProcParam(2) = Trim(Me.cbocta.Text)
        CryLMayor.StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
        CryLMayor.StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
        
        CryLMayor.Formulas(0) = "cta = '" & Trim(Me.cbocta.Text) & "'"
        CryLMayor.Formulas(1) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
        CryLMayor.Formulas(2) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
        CryLMayor.Formulas(4) = "nomcta = '" & Trim(Me.Lblcuenta) & "'"
        CryLMayor.Formulas(5) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
        CryLMayor.Formulas(6) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
        CryLMayor.Formulas(9) = "SIBs = " & SaldoIBs
        CryLMayor.Formulas(10) = "SISus = " & SaldoISus
        CryLMayor.Formulas(11) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
        CryLMayor.Formulas(12) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
        IResult = CryLMayor.PrintReport
        If IResult <> 0 Then
            MsgBox CryLMayor.LastErrorNumber & " : " & CryLMayor.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
End Sub

Private Sub txtbusca1_LostFocus()
Me.CmdAceptar.Enabled = True
End Sub

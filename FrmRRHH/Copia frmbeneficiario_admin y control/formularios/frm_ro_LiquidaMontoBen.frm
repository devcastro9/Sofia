VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_ro_LiquidaMontoBen 
   BackColor       =   &H00000000&
   Caption         =   "Montos de Liquidación"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox fraToolBarGuarda 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   0
      Picture         =   "frm_ro_LiquidaMontoBen.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   1035
      TabIndex        =   36
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   120
         Picture         =   "frm_ro_LiquidaMontoBen.frx":6C032
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   720
         Width           =   765
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "frm_ro_LiquidaMontoBen.frx":6C23C
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Cancelar"
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   9900
         TabIndex        =   39
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.TextBox tdnTcUS 
      Height          =   285
      Left            =   7800
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox tdnOrgContraBS 
      Height          =   285
      Left            =   7560
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tdnOrgContraUS 
      Height          =   285
      Left            =   6360
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tdnOrgBaseBS 
      Height          =   285
      Left            =   7560
      TabIndex        =   32
      Text            =   "Text2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tdnOrgBaseUS 
      Height          =   285
      Left            =   6360
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tdnNroLiq 
      Height          =   285
      Left            =   4440
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   220
      Width           =   375
   End
   Begin VB.TextBox tdnMontoBS 
      Height          =   285
      Left            =   7560
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox tdnMontoUS 
      Height          =   285
      Left            =   6360
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame fraNveces 
      BackColor       =   &H00404040&
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   0
      Width           =   4095
      Begin MSComCtl2.UpDown udwNveces 
         Height          =   375
         Left            =   3735
         TabIndex        =   6
         Top             =   160
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1
         OrigLeft        =   4920
         OrigTop         =   840
         OrigRight       =   5160
         OrigBottom      =   1260
         Max             =   12
         Min             =   1
         Enabled         =   0   'False
      End
      Begin VB.Label lblNveces 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Liquidación:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.TextBox txtFteFinanHist 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3840
      Width           =   5175
   End
   Begin VB.TextBox txtNroConHist 
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3840
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo cboTipoMonedaBen 
      Height          =   315
      Left            =   6720
      TabIndex        =   10
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      Style           =   2
      BackColor       =   12648447
      Text            =   ""
   End
   Begin VB.Label lblBeneficiario 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   15
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label lblPorcTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16394
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5760
      TabIndex        =   4
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblPorcOrgContra 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16394
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   5760
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblPorcOrgBase 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16394
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   5760
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monto Bs."
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
      Height          =   285
      Left            =   7560
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monto $US."
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
      Height          =   285
      Left            =   6360
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblDesOrgBase 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monto a Ejecutar: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   1200
      TabIndex        =   26
      Top             =   1920
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label lblDesOrgContra 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Organismo contraparte: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   1200
      TabIndex        =   25
      Top             =   2280
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo moneda:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label50 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TDC:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   6840
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label46 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Límite de liquidación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3840
      TabIndex        =   24
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label lblMontoPendUS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16394
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   300
      Left            =   6360
      TabIndex        =   23
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblMontoPendBS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16394
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   300
      Left            =   7560
      TabIndex        =   22
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblMontoLimiteUS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16394
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6360
      TabIndex        =   21
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblMontoLimiteBS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16394
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7560
      TabIndex        =   20
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pendiente por Pagar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   300
      Left            =   3840
      TabIndex        =   19
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Financiador:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   7200
      TabIndex        =   13
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "FINANCIAMIENTO"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   27
      Top             =   1440
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5760
      TabIndex        =   0
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Monto Total Pagar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   300
      Left            =   3120
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Beneficiario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "frm_ro_LiquidaMontoBen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQLs As String ' usado para la elaboración de los querys
Dim CodBenef As String ' usado para guardar el codigo de beneficiario

Private Sub cboTipoMonedaBen_Change()
    cboTipoMonedaBen.ToolTipText = cboTipoMonedaBen.Text
    
    Select Case cboTipoMonedaBen.BoundText
      Case "$US"
'        tdnMontoUS.ReadOnly = False
        tdnMontoUS.Enabled = True
        
'        tdnMontoBS.ReadOnly = True
        tdnMontoBS.Enabled = False
      Case "Bs"
'        tdnMontoUS.ReadOnly = True
        tdnMontoUS.Enabled = False
        
'        tdnMontoBS.ReadOnly = False
        tdnMontoBS.Enabled = True
    End Select

End Sub

Private Sub cboTipoMonedaBen_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
      Case 13 ' si presiono enter
        SendKeys "{Tab}"
      Case 27 ' si presiono escape
        
        Call cmdSalir_Click
        
    End Select

End Sub

Private Sub cmdGuardar_Click()
    Dim swGuardar As Integer ' usado para saber si efectivamente se almaceno o elimino los datos en la base
                          ' swGuarda -> 0 si se realizo el proceso satisfactoriamente
                          ' swGuarda -> 1 si se produjo un evento de cancelar por parte del usuario en el proceso
                          ' swGuarda -> 2 si se produjo un error de integridad de la base de datos en el servidor por el proceso

    Screen.MousePointer = vbHourglass

    If fl_VerificaBeneficiario Then ' verificamos si la información está correcta antes de actualizar la BD
        
        Call pl_GuardarBeneficiario(swGuardar)

        If swGuardar = 0 Then   ' si el proceso se realizo satisfactoriamente
            frm_ro_LiquidaMain.lblEstadoBeneficiario.Caption = CodBenef ' guarada el codigo beneficiario
            Unload Me
        ElseIf swGuardar = 2 Then ' si se cancelo el proceso por un error controlado por el servidor
            If cboTipoMonedaBen.BoundText = "Bs" Then
                tdnMontoBS.SetFocus
              Else
                tdnMontoUS.SetFocus
            End If
        End If
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSalir_Click()
    If vbYes = MsgBox("Desea mostrar los valores originales, perdiendo cualquier modificación realizada?", vbDefaultButton2 + vbYesNo + vbQuestion, "Aviso") Then
        Unload Me
      Else
        Call cboTipoMonedaBen_Change
    End If

End Sub

Private Sub Form_Load()
    Dim rstTemp As ADODB.Recordset ' usado para la carga de los datos
    Dim fechax As String
    Dim horax  As String
    
    On Error GoTo EtiqError
    
    ' carga tipo de moneda base
    Set rstTemp = New ADODB.Recordset
    SQLs = "select tipo_moneda, tipo_moneda + ' - ' + denominacion_moneda as des_tipo_moneda from tipo_moneda where activo='S'"
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        Set cboTipoMonedaBen.RowSource = rstTemp
        cboTipoMonedaBen.BoundColumn = "tipo_moneda"
        cboTipoMonedaBen.ListField = "des_tipo_moneda"
      Else
        MsgBox "El catálogo de tipo de moneda no esta actualizado", vbInformation, "Aviso"
    End If
    
    
    ' obtiene datos de liquidacion, beneficiario y montos
    tdnNroLiq.Text = Val(frm_ro_LiquidaMain.grdBeneficiario.Tag) ' numero de liquidación
    CodBenef = frm_ro_LiquidaMain.lblEstadoBeneficiario.Tag ' codigo de beneficario
    
    Call pl_RefrescaDatosOrg
    lblMontoPendUS.Tag = lblMontoPendUS.Caption ' se guarda el valor para realizar un control al guardar y saldos
    lblMontoPendBS.Tag = lblMontoPendBS.Caption ' se guarda el valor para realizar un control al guardar y saldos
    
    ' obtiene datos de beneficiarios del pago
    ' dependiendo del tipo de proceso si es consultor por F05==>"producto- corto plazo"  o F10==>consultor por "tiempo - largo pazo"
    Select Case glProceso
      Case "F05"
        SQLs = "SELECT fc_beneficiario.paterno_beneficiario as paterno, fc_beneficiario.materno_beneficiario as materno, fc_beneficiario.nombres_beneficiario as nombre, ao_pagos_cronograma_detalle.codigo_beneficiario, ao_pagos_cronograma_detalle.monto_us, ao_pagos_cronograma_detalle.monto_bs, ao_pagos_cronograma_detalle.tc_us, ao_pagos_cronograma_detalle.tipo_moneda,"
        SQLs = SQLs & "ao_pagos_cronograma_detalle.emite_factura, ao_pagos_cronograma_detalle.estado_conformidad, ao_pagos_cronograma_detalle.estado_devengado, ao_pagos_cronograma_detalle.ncite_conformidad, ao_pagos_cronograma_detalle.fcite_conformidad, ao_pagos_cronograma_detalle.Numero_consultoriaHist, ao_pagos_cronograma_detalle.fte_financiamientoHist "
        SQLs = SQLs & "FROM ao_pagos_cronograma_detalle INNER JOIN fc_beneficiario ON ao_pagos_cronograma_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario "
        SQLs = SQLs & "WHERE ao_pagos_cronograma_detalle.ges_gestion = '" & frm_ro_LiquidaMain.lblGestion.Caption & "' AND ao_pagos_cronograma_detalle.codigo_grupo = " & Val(frm_ro_LiquidaMain.lblCodGrupo.Caption) & " AND ao_pagos_cronograma_detalle.codigo_unidad = '" & frm_ro_LiquidaMain.lblCodUniSol.Caption & "' AND ao_pagos_cronograma_detalle.numero_pago = " & Val(frm_ro_LiquidaMain.grdBeneficiario.Tag) & " AND ao_pagos_cronograma_detalle.correlativo_reg = " & Val(frm_ro_LiquidaMain.grdLiquida.Tag) & " AND ao_pagos_cronograma_detalle.codigo_beneficiario ='" & CodBenef & "'"
      
      Case "F10"
        SQLs = "SELECT RC_Personal.paterno as paterno, RC_Personal.materno as materno, RC_Personal.nombres as nombre, ao_pagos_cronograma_detalle.codigo_beneficiario, ao_pagos_cronograma_detalle.monto_us, ao_pagos_cronograma_detalle.monto_bs, ao_pagos_cronograma_detalle.tc_us, ao_pagos_cronograma_detalle.tipo_moneda,"
        SQLs = SQLs & "ao_pagos_cronograma_detalle.emite_factura, ao_pagos_cronograma_detalle.estado_conformidad, ao_pagos_cronograma_detalle.estado_devengado, ao_pagos_cronograma_detalle.ncite_conformidad, ao_pagos_cronograma_detalle.fcite_conformidad, ao_pagos_cronograma_detalle.Numero_consultoriaHist, ao_pagos_cronograma_detalle.fte_financiamientoHist "
        SQLs = SQLs & "FROM ao_pagos_cronograma_detalle INNER JOIN RC_Personal ON ao_pagos_cronograma_detalle.codigo_beneficiario = RC_Personal.ci "
        SQLs = SQLs & "WHERE ao_pagos_cronograma_detalle.ges_gestion = '" & frm_ro_LiquidaMain.lblGestion.Caption & "' AND ao_pagos_cronograma_detalle.codigo_grupo = " & Val(frm_ro_LiquidaMain.lblCodGrupo.Caption) & " AND ao_pagos_cronograma_detalle.codigo_unidad = '" & frm_ro_LiquidaMain.lblCodUniSol.Caption & "' AND ao_pagos_cronograma_detalle.numero_pago = " & Val(frm_ro_LiquidaMain.grdBeneficiario.Tag) & " AND ao_pagos_cronograma_detalle.correlativo_reg = " & Val(frm_ro_LiquidaMain.grdLiquida.Tag) & " AND ao_pagos_cronograma_detalle.codigo_beneficiario ='" & CodBenef & "'"

    End Select
    
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        lblBeneficiario.Caption = rstTemp!paterno + " " + rstTemp!materno + " " + rstTemp!Nombre
        cboTipoMonedaBen.BoundText = Trim(rstTemp!tipo_moneda)
        
        tdnMontoUS.Text = IIf(IsNull(rstTemp!monto_us), 0, rstTemp!monto_us)
        tdnMontoBS.Text = IIf(IsNull(rstTemp!monto_bs), 0, rstTemp!monto_bs)
'        tdnMontoUS.Tag = IIf(IsNull(rstTemp!monto_us), 0, rstTemp!monto_us) ' se guarda el valor para realizar un control al guardar y saldos
'        tdnMontoBS.Tag = IIf(IsNull(rstTemp!monto_bs), 0, rstTemp!monto_bs) ' se guarda el valor para realizar un control al guardar y saldos
        tdnOrgBaseUS.Text = tdnMontoUS.Text * Val(lblPorcOrgBase.Caption) / 100
        tdnOrgBaseBS.Text = tdnMontoBS.Text * Val(lblPorcOrgBase.Caption) / 100
        tdnOrgContraUS.Text = tdnMontoUS.Text * Val(lblPorcOrgContra.Caption) / 100
        tdnOrgContraBS.Text = tdnMontoBS.Text * Val(lblPorcOrgContra.Caption) / 100
        
        txtFteFinanHist.Text = rstTemp!numero_consultoriaHist & ""
        txtNroConHist.Text = rstTemp!fte_financiamientoHist & ""
      
        'JQ QR
        'DE.dbo_edGetProcessDateTime fechax, horax
        '' OBTIENE EL TIPO DE CAMBIO DEL DOLAR DEL DIA
        SQLs = "select cambio_oficial from ac_tipo_cambio where fecha_cambio = '" & fechax & "'"
        Set rstTemp = New ADODB.Recordset
        rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
        If rstTemp.RecordCount > 0 Then
            tdnTcUS.Text = rstTemp!cambio_oficial
        End If
      
      Else
        lblBeneficiario.Caption = ""
        cboTipoMonedaBen.BoundText = ""
        tdnTcUS.Text = 0
        tdnMontoUS.Text = 0
        tdnMontoBS.Text = 0
        tdnOrgBaseUS.Text = 0
        tdnOrgBaseBS.Text = 0
        tdnOrgContraUS.Text = 0
        tdnOrgContraBS.Text = 0
        
        txtFteFinanHist.Text = ""
        txtNroConHist.Text = ""
    
    End If
    
    Call cboTipoMonedaBen_Change
    Call pl_CalSubTotLim
    Screen.MousePointer = vbDefault
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub
    

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
    
	Call SeguridadSet(Me)
End Sub

Private Sub pl_RefrescaDatosOrg()
    'TITULO:                Procedimiento pl_RefrescaDatosOrg
    'PROPOSITO:             Actualiza los datos de organismo financiador
    'EJEMPLO DE LLAMADA:    call pl_RefrescaDatosOrg
    
    Dim MontoEXT_US As Currency
    Dim MontoNAL_US As Currency
    Dim MontoEXT_BS As Currency
    Dim MontoNAL_BS As Currency
    Dim MontoUSasignado As Currency
    Dim MontoBSasignado As Currency
    Dim DesOrgBase As String
    Dim DesOrgContra As String
    Dim PorcOrgBase As Double
    Dim PorcOrgContra As Double
    
    On Error GoTo EtiqError
    
    ' calcula montos limites y saldos
    'JQ QR
    'DE.dbo_ap_PagosSumaMontoLimBen frm_ro_LiquidaMain.lblGestion.Caption, frm_ro_LiquidaMain.lblCodUniSol.Caption, frm_ro_LiquidaMain.lblCodGrupo, CodBenef, MontoEXT_US, MontoNAL_US, MontoEXT_BS, MontoNAL_BS, MontoUSasignado, MontoBSasignado, DesOrgBase, DesOrgContra, PorcOrgBase, PorcOrgContra
    
    lblMontoLimiteUS.Caption = Format(MontoEXT_US + MontoNAL_US, "######0.00")
    lblMontoLimiteBS.Caption = Format(MontoEXT_BS + MontoNAL_BS, "######0.00")
    
    lblMontoPendUS.Caption = Format((MontoEXT_US + MontoNAL_US) - MontoUSasignado, "######0.00")
    lblMontoPendBS.Caption = Format((MontoEXT_BS + MontoNAL_BS) - MontoBSasignado, "######0.00")
    
    lblDesOrgBase.Caption = DesOrgBase & ": " ' organismo externo
    lblPorcOrgBase.Caption = Format(PorcOrgBase, "######0.00")
       
    If Len(Trim(DesOrgContra)) > 0 Then ' si tiene financiamiento de organismo contraparte (nacional)
        lblDesOrgContra.Caption = DesOrgContra & ": "
        lblPorcOrgContra.Caption = Format(PorcOrgContra, "######0.00")
      Else
        lblDesOrgContra.Caption = "Contraparte Nacional: "
        lblPorcOrgContra.Caption = "0.00"
    End If
    lblPorcTotal.Caption = PorcOrgBase + PorcOrgContra
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
    
End Sub

Private Sub lblDesOrgBase_Change()
    lblDesOrgBase.ToolTipText = lblDesOrgBase.Caption
End Sub

Private Sub lblDesOrgContra_Change()
    lblDesOrgContra.ToolTipText = lblDesOrgContra.Caption
End Sub

Private Sub tdnMontoBS_Change()
    If cboTipoMonedaBen.BoundText = "Bs" Then
        Call pl_CalSubTotLim
    End If
End Sub

Private Sub tdnMontoBS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub tdnMontoUS_Change()
    If cboTipoMonedaBen.BoundText = "$US" Then
        Call pl_CalSubTotLim
    End If
End Sub

Private Sub pl_CalSubTotLim()
    
    On Error GoTo EtiqError ' desactivamos el manejador de errores
    
    Select Case cboTipoMonedaBen.BoundText
      Case "$US"
        'If tdnMontoUS.Text * tdnTcUS.Text <= tdnMontoBS.MaxValue Then  'JQA JUL/2008
        If tdnMontoUS.Text * tdnTcUS.Text <= tdnMontoBS.Text Then
            tdnMontoBS.Text = tdnMontoUS.Text * tdnTcUS.Text
        End If
      Case "Bs"
        If tdnTcUS.Text > 0 Then
            tdnMontoUS.Text = tdnMontoBS.Text / tdnTcUS.Text
        End If
    End Select
    
    tdnOrgBaseUS.Text = tdnMontoUS.Text * Val(lblPorcOrgBase.Caption) / 100
    tdnOrgBaseBS.Text = tdnMontoBS.Text * Val(lblPorcOrgBase.Caption) / 100
    tdnOrgContraUS.Text = tdnMontoUS.Text * Val(lblPorcOrgContra.Caption) / 100
    tdnOrgContraBS.Text = tdnMontoBS.Text * Val(lblPorcOrgContra.Caption) / 100

    lblMontoPendUS.Caption = Format((Val(lblMontoPendUS.Tag) + Val(tdnMontoUS.Tag)) - tdnMontoUS.Text, "######0.00")
    lblMontoPendBS.Caption = Format((Val(lblMontoPendBS.Tag) + Val(tdnMontoBS.Tag)) - tdnMontoBS.Text, "######0.00")
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Select Case Err.Number
      Case 380
        Select Case cboTipoMonedaBen.BoundText
          Case "$US"
            tdnMontoUS.Text = IIf(Val(lblMontoPendUS.Caption) >= 0, lblMontoPendUS.Caption, IIf(Val(lblMontoPendUS.Tag) < 0, 0, Val(lblMontoPendUS.Tag)))
          Case "Bs"
            tdnMontoBS.Text = IIf(Val(lblMontoPendBS.Caption) >= 0, lblMontoPendBS.Caption, IIf(Val(lblMontoPendBS.Tag) < 0, 0, Val(lblMontoPendBS.Tag)))
        End Select
      Case Else ' si se produjo otro tipo de error
        MsgBox "Error: Valor indterminado." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
    End Select

End Sub

Private Function fl_VerificaBeneficiario() As Boolean
    'TITULO:                Función fl_VerificaBeneficiario
    'PROPOSITO:             Verifica los datos para el registro de una beneficiario
    'EJEMPLO DE LLAMADA:    fl_VerificaBeneficiario
    
    fl_VerificaBeneficiario = True ' asuminos que se cuenta con los datos mnimos para grabar

    ' verificamos tipo de cambio US
    If Len(Trim(cboTipoMonedaBen.BoundText)) = 0 Then
        MsgBox "El tipo de moneda no es válido." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        cboTipoMonedaBen.SetFocus
        fl_VerificaBeneficiario = False
        Exit Function
    End If

    If tdnTcUS.Text <= 0 Or IsNull(tdnTcUS.Text) Then
        MsgBox "El tipo de cambio $US no es válido." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        tdnTcUS.SetFocus
        fl_VerificaBeneficiario = False
        Exit Function
    End If
    
    ' verificamos montos  para tipo de moneda $US
    If Trim(cboTipoMonedaBen.BoundText) = "$US" And (tdnMontoUS.Text <= 0 Or tdnMontoUS.Text = Null) Then
        MsgBox "El monto en $US a liquidar correspondiente a [" & lblBeneficiario.Caption & "] no es válido." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        tdnMontoUS.SetFocus
        fl_VerificaBeneficiario = False
        Exit Function
    End If

    ' verificamos montos  para tipo de moneda BS
    If Trim(cboTipoMonedaBen.BoundText) = "Bs" = True And (tdnMontoBS.Text <= 0 Or tdnMontoBS.Text = Null) Then
        MsgBox "El monto a liquidar en BS no es válido." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        tdnMontoBS.SetFocus
        fl_VerificaBeneficiario = False
        Exit Function
    End If

    ' verificamos montos  no sobre pase el límite para tipo de moneda $US
    If Trim(cboTipoMonedaBen.BoundText) = "$US" And Val(lblMontoPendUS.Caption) < 0 Then
        MsgBox "El monto total de liquidación [" & tdnMontoUS.Text & " $US] excede en [" & Format((-1) * Val(lblMontoPendUS.Caption), "######0.00") & " $US] al monto límite a liquidar [" & lblMontoLimiteUS.Caption & " $US]." & Chr(13) & "Correspondiente a [" & lblBeneficiario.Caption & "]." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        tdnMontoUS.SetFocus
        fl_VerificaBeneficiario = False
        Exit Function
    End If

    ' verificamos montos no sobre pase el límite para tipo de moneda BS
    If Trim(cboTipoMonedaBen.BoundText) = "Bs" And Val(lblMontoPendBS.Caption) < 0 Then
        MsgBox "El monto total de liquidación [" & tdnMontoBS.Text & " BS] excede en [" & Format((-1) * Val(lblMontoPendBS.Caption), "######0.00") & " BS] al monto límite a liquidar [" & lblMontoLimiteBS.Caption & " $US]." & Chr(13) & "Correspondiente a [" & lblBeneficiario.Caption & "]." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        tdnMontoBS.SetFocus
        fl_VerificaBeneficiario = False
        Exit Function
    End If
    
End Function

Private Sub pl_GuardarBeneficiario(ByRef swGuarda As Integer)
    ' guarda la información en la base de datos
    ' swGuarda -> 0 si se realizo el proceso satisfactoriamente
    ' swGuarda -> 1 si se produjo un evento de cancelar por parte del usuario en el proceso
    ' swGuarda -> 2 si se produjo un error de integridad de la base de datos en el servidor por el proceso

    On Error GoTo EtiqError
    
    Select Case frm_ro_LiquidaMain.lblEstadoBeneficiario.Caption
      Case "E" ' se esta editando el registro actual
        'JQ QR
        'DE.dbo_ap_PagosGrabaPagoBenef frm_ro_LiquidaMain.lblGestion.Caption, frm_ro_LiquidaMain.lblCodUniSol.Caption, Val(frm_ro_LiquidaMain.lblCodGrupo.Caption), tdnNroLiq.Text, Val(frm_ro_LiquidaMain.grdLiquida.Tag), CodBenef, cboTipoMonedaBen.BoundText, tdnTcUS.Text, tdnMontoUS.Text, tdnMontoBS.Text, txtNroConHist.Text, txtFteFinanHist.Text, GlUsuario
        
      Case "N" ' se registra uno nuevo
'
        ' no se procesa en este modulo
        
    End Select
    
    swGuarda = 0 ' si llego a esta parte es porque los cambios se realizaron efectivamente en la base de datos
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub
    
EtiqError:
    Select Case Err.Number
      Case -2147217900
        MsgBox "Error: No puede existir dos códigos iguales, corrija el error y vuelva a intentarlo." & Chr(13) & "Los cambios no se llevaron a cabo." & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description, vbCritical, "Error"
        swGuarda = 2
      Case Else ' si se produjo otro tipo de error
        MsgBox "Error: Los cambios no se llevaron a cabo." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
        swGuarda = 2
    End Select
    
End Sub

Private Sub tdnMontoUS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub tdnNroLiq_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub tdnOrgBaseBS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub tdnOrgBaseUS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub tdnOrgContraBS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub tdnOrgContraUS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub tdnTcUS_Change()
    Call pl_CalSubTotLim
End Sub

Private Sub tdnTcUS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub txtFteFinanHist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

Private Sub txtNroConHist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Call cmdSalir_Click
    End If

End Sub

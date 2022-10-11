VERSION 5.00
Begin VB.Form FrmTipoCambio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Cambio"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "FrmTipoCambio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmTipoCambio.frx":0A02
   ScaleHeight     =   4650
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_rmb 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "###,##0.00000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   2595
      TabIndex        =   5
      Text            =   "0"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txt_brl 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "###,##0.00000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   2595
      TabIndex        =   4
      Text            =   "0"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txt_eur 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "###,##0.00000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   2595
      TabIndex        =   3
      Text            =   "0"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txt_ufv 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "###,##0.00000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   2595
      TabIndex        =   2
      Text            =   "0"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox tdbnMercado 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "###,##0.00000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   2595
      TabIndex        =   1
      Text            =   "0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox tdbnOficial 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "###,##0.00000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   2595
      TabIndex        =   0
      Text            =   "0"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton BtnCancelar 
      BackColor       =   &H00808000&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   705
      Left            =   2760
      Picture         =   "FrmTipoCambio.frx":B33B
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "Cancelar"
      Top             =   3840
      Width           =   900
   End
   Begin VB.CommandButton BtnGrabar 
      BackColor       =   &H00808000&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   705
      Left            =   960
      Picture         =   "FrmTipoCambio.frx":B545
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "Aceptar"
      Top             =   3840
      Width           =   900
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   4815
      Begin VB.Label tdbfTitulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cotización por Tipo de Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   480
         TabIndex        =   14
         Top             =   200
         Width           =   3765
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   885
      Left            =   -90
      TabIndex        =   12
      Top             =   3720
      Width           =   4905
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFC0&
      X1              =   0
      X2              =   4920
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bs por RMB."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3765
      TabIndex        =   22
      Top             =   3135
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bs por BRL."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3765
      TabIndex        =   21
      Top             =   2655
      Width           =   855
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bs por EUR"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3765
      TabIndex        =   20
      Top             =   2175
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bs por UFV"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3765
      TabIndex        =   19
      Top             =   1695
      Width           =   810
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cambio - China"
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
      Height          =   195
      Left            =   465
      TabIndex        =   18
      Top             =   3135
      Width           =   1995
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cambio - Brasil"
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
      Height          =   195
      Left            =   480
      TabIndex        =   17
      Top             =   2655
      Width           =   1980
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cambio - España"
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
      Height          =   195
      Left            =   330
      TabIndex        =   16
      Top             =   2175
      Width           =   2130
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cambio para UFV's"
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
      Height          =   195
      Left            =   135
      TabIndex        =   15
      Top             =   1695
      Width           =   2325
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cambio Oficial"
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
      Height          =   195
      Left            =   585
      TabIndex        =   11
      Top             =   900
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cambio para Compra"
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
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   1125
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bs por Dolar."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3765
      TabIndex        =   9
      Top             =   885
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bs por Dolar."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3765
      TabIndex        =   8
      Top             =   1140
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "FrmTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TipoCambioHoy As Boolean
Dim LcSQLAux As String
Dim Fecha As Date
Dim Usuario As String
Private Sub BtnGrabar_Click()
    '
    If CDbl(tdbnOficial) <= 0 Then
        MsgBox "El Tipo de Cambio de Venta debe ser mayor a CERO.", vbInformation + vbOKOnly, "Atención"
        tdbnOficial.SetFocus
        Exit Sub
    End If
    If CDbl(tdbnMercado) <= 0 Then
        MsgBox "El Tipo de Cambio de Compra debe ser mayor a CERO.", vbInformation + vbOKOnly, "Atención"
        tdbnMercado.SetFocus
        Exit Sub
    End If
    If CDbl(tdbnOficial) > 49 Then
        MsgBox "Error en el Tipo de Cambio de Venta, vuelva a intentar ...", vbInformation + vbOKOnly, "Atención"
        tdbnOficial.SetFocus
        Exit Sub
    End If
    If CDbl(tdbnMercado) > 49 Then
        MsgBox "Error en el Tipo de Cambio de Compra, vuelva a intentar ...", vbInformation + vbOKOnly, "Atención"
        tdbnMercado.SetFocus
        Exit Sub
    End If
    If CDbl(txt_ufv) <= 0 Then
        MsgBox "El Tipo de Cambio UFV debe ser mayor a CERO.", vbInformation + vbOKOnly, "Atención"
        txt_ufv.SetFocus
        Exit Sub
    End If
    If CDbl(txt_eur) <= 0 Then
        MsgBox "El Tipo de Cambio Euros debe ser mayor a CERO.", vbInformation + vbOKOnly, "Atención"
        txt_eur.SetFocus
        Exit Sub
    End If
    If CDbl(txt_brl) <= 0 Then
        MsgBox "El Tipo de Cambio Reales debe ser mayor a CERO.", vbInformation + vbOKOnly, "Atención"
        txt_brl.SetFocus
        Exit Sub
    End If
    If CDbl(txt_rmb) <= 0 Then
        MsgBox "El Tipo de Cambio Reminbis debe ser mayor a CERO.", vbInformation + vbOKOnly, "Atención"
        txt_rmb.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Se almacenará el siguiente Tipo de Cambio para la Fecha '" & Date & "'," & vbCrLf & vbCrLf & _
              vbTab & "- Tipo de Cambio USD para la Venta = " & Format(tdbnOficial, "###,###,##0.00") & vbCrLf & _
              vbTab & "- Tipo de Cambio USD para la Compra = " & Format(tdbnMercado, "###,###,##0.00") & vbCrLf & vbCrLf & _
              vbTab & "- Tipo de Cambio para UFV = " & Format(txt_ufv, "###,###,##0.00") & vbCrLf & _
              vbTab & "- Tipo de Cambio para Euros = " & Format(txt_eur, "###,###,##0.00") & vbCrLf & vbCrLf & _
              vbTab & "- Tipo de Cambio para Reales = " & Format(txt_brl, "###,###,##0.00") & vbCrLf & _
              vbTab & "- Tipo de Cambio para Reminbis = " & Format(txt_rmb, "###,###,##0.00") & vbCrLf & vbCrLf & _
              "Confirmar y almacenar esta información?", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
              
    '"Es esta información correcta? " & vbCrLf & _
    ' Almacenamos en la Base de datos
    LcSQLAux = "INSERT INTO gc_tipo_cambio (Fecha_Cambio, cambio_oficial_compra, cambio_mercado_venta, cambio_ufv, cambio_eur, cambio_rmb, cambio_brl, usr_codigo, Fecha_Registro, Hora_Registro) " & _
               "SELECT '" & CDate(Fecha) & "', " & CDbl(tdbnMercado) & ", " & CDbl(tdbnOficial) & ", " & CDbl(txt_ufv) & ", " & CDbl(txt_eur) & ", " & CDbl(txt_rmb) & ", " & CDbl(txt_brl) & ", '" & Usuario & "', '" & CDate(Date) & "', '" & Format(Time, "hh:mm:ss") & "'"
    db.Execute LcSQLAux
    GlTipoCambioOficial = tdbnOficial     ' Compra Dolar
    GlTipoCambioMercado = tdbnMercado     ' Venta Dolar
    GlTipoCambioEuro = txt_eur            'Euros España y otros
    GlTipoCambioGestion = tdbnOficial     'Para cierre Dolar
    GlTipoCambioUfv = txt_ufv             'UFV
    GlTipoCambioRmb = txt_rmb             'Reminbis China
    GlTipoCambioBrl = txt_brl             'Reales Brasil

    TipoCambioHoy = True
    Unload Me
End Sub

Private Sub BtnCancelar_Click()
    Unload Me
End Sub

Public Sub TcPrincipal(AQueFecha As Date, _
                       QueUsuario As String)
    Fecha = AQueFecha
    Usuario = QueUsuario
    'Verificamos si ya existe Tipo de cambio para AQueFecha
    If ExisteTCambio(AQueFecha, GlTipoCambioOficial, GlTipoCambioMercado, GlTipoCambioEuro, GlTipoCambioUfv, GlTipoCambioRmb, GlTipoCambioBrl) Then
        TipoCambioHoy = True
        Exit Sub
    Else
        If Not EsAdministrador(QueUsuario) Then Exit Sub
    End If
    '
    tdbfTitulo.Caption = "Fecha : " & Format(AQueFecha, "dd/mm/yyyy")
    '
    Me.Show vbModal
End Sub

Private Sub Form_Load()
    TipoCambioHoy = False
End Sub

Public Function ExisteTCambio(ByVal QueFecha As Date, _
                              ByRef TipoCambioOficial As Currency, _
                              ByRef TipoCambioMercado As Currency, _
                              ByRef TipoCambioEuro As Currency, _
                              ByRef TipoCambioUfv As Currency, _
                              ByRef TipoCambioRmb As Currency, _
                              ByRef TipoCambioBrl As Currency) As Boolean
Dim rsAux As ADODB.Recordset
    TipoCambioOficial = 0
    TipoCambioMercado = 0
    Set rsAux = New ADODB.Recordset
    LcSQLAux = "SELECT * FROM gc_Tipo_Cambio WHERE Fecha_Cambio = '" & QueFecha & "'"
    rsAux.Open LcSQLAux, db, adOpenStatic
    ExisteTCambio = rsAux.RecordCount > 0
    If ExisteTCambio Then TipoCambioOficial = rsAux!cambio_oficial_compra: TipoCambioMercado = rsAux!cambio_mercado_venta: TipoCambioEuro = rsAux!cambio_eur: TipoCambioUfv = rsAux!cambio_ufv: TipoCambioRmb = rsAux!cambio_rmb: TipoCambioBrl = rsAux!cambio_brl
    
'    Public GlTipoCambioOficial As Currency  'Compra Dolar
'Public GlTipoCambioMercado As Currency  'Venta Dolar
'Public GlTipoCambioGestion As Currency  'Para cierre Dolar
'Public GlTipoCambioEuro As Currency     'Euros España y otros
'Public GlTipoCambioUfv As Currency      'UFV
'Public GlTipoCambioRmb As Currency      'Reminbis China
'Public GlTipoCambioBrl As Currency      'Reales Brasil

End Function

Public Function EsAdministrador(QueUsuario As String) As Boolean
Dim rsAux As ADODB.Recordset
    Set rsAux = New ADODB.Recordset
    LcSQLAux = "SELECT * FROM gc_usuarios WHERE usr_codigo = '" & QueUsuario & "' AND (IDNivelAcceso = 1 OR IDNivelAcceso = 4 OR IDNivelAcceso = 20) "    ' Se modifico por que todos los usuarios tienen que tener un nivel de acceso
    rsAux.Open LcSQLAux, db, adOpenStatic
    EsAdministrador = rsAux.RecordCount > 0
End Function

Private Sub tdbnMercado_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
  '? . , 09
  ',.01234856789
End Sub

Private Sub tdbnOficial_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub tdbnOficial_LostFocus()
    If tdbnOficial.Text = "" Or tdbnOficial.Text = "0" Then
        tdbnMercado.Text = "6.96"
    Else
        tdbnMercado.Text = tdbnOficial.Text
    End If
End Sub

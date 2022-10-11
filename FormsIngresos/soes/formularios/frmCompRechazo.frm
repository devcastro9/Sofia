VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCompRechazo 
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   Icon            =   "frmCompRechazo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1035
      Left            =   135
      TabIndex        =   9
      Top             =   2415
      Width           =   5445
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   720
         Left            =   3360
         MousePointer    =   4  'Icon
         Picture         =   "frmCompRechazo.frx":324A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   210
         Width           =   1005
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   720
         Left            =   4350
         MousePointer    =   4  'Icon
         Picture         =   "frmCompRechazo.frx":3554
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   210
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2340
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   5430
      Begin VB.TextBox txtsor_razon_recha 
         DataField       =   "sor_razon_recha"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   945
         Left            =   1830
         TabIndex        =   3
         Top             =   570
         Width           =   3375
      End
      Begin VB.TextBox txtsor_ded_us 
         DataField       =   "sor_ded_us"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   1845
         TabIndex        =   2
         Top             =   1560
         Width           =   1320
      End
      Begin VB.TextBox txtsor_ded_bs 
         DataField       =   "sor_ded_bs"
         DataMember      =   "Command1"
         DataSource      =   "Datos"
         Height          =   285
         Left            =   1845
         TabIndex        =   1
         Top             =   1935
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtp_Fecha 
         Height          =   255
         Left            =   1860
         TabIndex        =   4
         Top             =   195
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   450
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   36656
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Rechazo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   8
         Top             =   225
         Width           =   1680
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Razon:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   1200
         TabIndex        =   7
         Top             =   615
         Width           =   615
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Deducción en Us."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   270
         TabIndex        =   6
         Top             =   1605
         Width           =   1545
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Deducción en Bs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   285
         TabIndex        =   5
         Top             =   1980
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmCompRechazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim accion As String, Cancelar As Boolean
Public frmCompRechazo_ret As String

Public Sub frmCompRechazo_procesar(proceso As String)
  Cancelar = True
  accion = proceso
  dtp_Fecha.Value = Date
  frmCompRechazo_ret = ""
  Caption = "Ingrese datos de la Planilla de deducciones"
  Show vbModal
End Sub

Private Sub CmdCancelar_Click()
  Cancelar = True
  Unload Me
End Sub

Private Sub CmdGrabar_Click()
  If MsgBox("Esta seguro de Grabar?", vbYesNo) = vbYes Then
    If validaRegistro Then
      Cancelar = False
      Datos.dbo_so_detalle_soes_rechazo "INSERT", frmDetalleSoes.adoDetalleSoes.Recordset!soc_nro_sol, frmDetalleSoes.adoDetalleSoes.Recordset!soe_cod_convenio, frmDetalleSoes.adoDetalleSoes.Recordset!soe_nro_sec, frmDetalleSoes.adoDetalleSoes.Recordset!ges_gestion, frmDetalleSoes.adoDetalleSoes.Recordset!org_codigo, frmDetalleSoes.adoDetalleSoes.Recordset!codigo_pago, frmDetalleSoes.adoDetalleSoes.Recordset!dso_nro_veces, Me.dtp_Fecha.Value, Me.txtsor_razon_recha, Val(Me.txtsor_ded_us), Val(Me.txtsor_ded_bs), Date, glusuario
      Unload Me
    End If
  End If
End Sub

Function validaRegistro() As Boolean
Dim ok As Boolean
ok = True
  If ok And Me.dtp_Fecha.Value > frmSoesMain.dtp_Fecha.Value Then
    ok = False
    MsgBox "Ingrese fecha de Rechazo mayor a la fecha de solicitud " & frmSoesMain.dtp_Fecha.Value
  End If
  If ok And (Me.dtp_Fecha.Value > Date Or Me.dtp_Fecha.Value < Date - 30) Then
    ok = False
    MsgBox "Ingrese fecha de Rechazo entre hoy y 30 dias anteriores a hoy"
  End If
  If ok And Me.txtsor_razon_recha = "" Then
    ok = False
    MsgBox "Ingrese razon del Rechazo"
  End If
  If ok And Val(Me.txtsor_ded_bs) <= 0 Then
    ok = False
    MsgBox "Ingrese Deducciones en Bs con valor mayor a cero"
  End If
  If ok And Val(Me.txtsor_ded_us) <= 0 Then
    ok = False
    MsgBox "Ingrese Deducciones en Us con valor mayor a cero"
  End If
'  If ok And  Then
'    ok = False
'    MsgBox ""
'  End If
  validaRegistro = ok
End Function

Private Sub Form_Unload(Cancel As Integer)
  If Cancelar Then
    If MsgBox("¿Desea Salir de esta ventana y cancelar los cambios ingresados?", vbQuestion + vbYesNo, "Diálogo Cerrar") = vbNo Then
      Cancel = -1
    End If
  End If
End Sub


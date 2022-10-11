VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ac_CapturaHojaVida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Hoja de Vida"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7395
   ControlBox      =   0   'False
   Icon            =   "ac_CapturaHojaVida.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Datos Hoja de Vida o Curriculum Vitae"
      ForeColor       =   &H00008000&
      Height          =   3495
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   6255
      Begin VB.TextBox TxtOcupacion 
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   44
         Top             =   3140
         Width           =   4695
      End
      Begin VB.TextBox TxtItem 
         Height          =   285
         Left            =   4320
         MaxLength       =   15
         TabIndex        =   36
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   3480
         MaxLength       =   80
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtBenef 
         Height          =   285
         Left            =   240
         MaxLength       =   80
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         MaxLength       =   80
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtTelefono 
         Height          =   285
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtRecomendaciones 
         Height          =   285
         Left            =   3840
         MaxLength       =   300
         TabIndex        =   12
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtPat 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1680
         Width           =   4695
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2040
         Width           =   4695
      End
      Begin VB.TextBox txtNom 
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   4
         Top             =   2400
         Width           =   4695
      End
      Begin VB.TextBox txtCI 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox cboTDoc 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComCtl2.DTPicker txtNac 
         DataField       =   "fecha_nacimiento"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   4320
         TabIndex        =   38
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   89653249
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker DTPFec_Seguro 
         DataField       =   "fecha_nacimiento"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   1440
         TabIndex        =   39
         Top             =   1200
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   89653249
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo Dtc_Par 
         Bindings        =   "ac_CapturaHojaVida.frx":0ECA
         DataField       =   "cod_pariente"
         DataSource      =   "frmBeneficiario.adoLista"
         Height          =   315
         Left            =   4560
         TabIndex        =   41
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "cod_pariente"
         BoundColumn     =   "cod_pariente"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Dtc_ParDes 
         Bindings        =   "ac_CapturaHojaVida.frx":0EE5
         DataField       =   "cod_pariente"
         DataSource      =   "frmBeneficiario.adoLista"
         Height          =   315
         Left            =   1440
         TabIndex        =   40
         Top             =   2760
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "nomb_pariente"
         BoundColumn     =   "cod_pariente"
         Text            =   ""
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ocupación"
         Height          =   195
         Index           =   3
         Left            =   550
         TabIndex        =   45
         Top             =   3120
         Width           =   780
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codigo Seguro Social"
         Height          =   195
         Index           =   11
         Left            =   4440
         TabIndex        =   37
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Parentesco"
         Height          =   195
         Index           =   8
         Left            =   525
         TabIndex        =   17
         Top             =   2760
         Width           =   810
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Nacimiento"
         Height          =   195
         Index           =   7
         Left            =   4560
         TabIndex        =   16
         Top             =   960
         Width           =   1290
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         Caption         =   "Aprobado"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Segundo apellido"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   11
         Top             =   2040
         Width           =   1230
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
         Height          =   195
         Index           =   1
         Left            =   705
         TabIndex        =   10
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Primer apellido"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   9
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Documento Identidad"
         Height          =   195
         Index           =   4
         Left            =   1500
         TabIndex        =   8
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Reg. Seguro"
         Height          =   195
         Index           =   5
         Left            =   1500
         TabIndex        =   7
         Top             =   960
         Width           =   1395
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Presupuesto requerido"
      ForeColor       =   &H00808000&
      Height          =   2895
      Left            =   1800
      TabIndex        =   20
      Top             =   240
      Width           =   5535
      Begin MSMask.MaskEdBox mskMonto_pendiente 
         Height          =   375
         Left            =   3720
         TabIndex        =   21
         Top             =   2400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   8421376
         ForeColor       =   16777215
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMonto_limite 
         Height          =   375
         Left            =   3720
         TabIndex        =   22
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   8421376
         ForeColor       =   16777215
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMonto_ext 
         Height          =   375
         Left            =   3720
         TabIndex        =   23
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   12648447
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMonto_nal 
         Height          =   375
         Left            =   3720
         TabIndex        =   24
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   12648447
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMonto 
         Height          =   375
         Left            =   3720
         TabIndex        =   25
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label labPorcExt 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0%"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   35
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pendiente de pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   34
         Top             =   2400
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Límite del pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   33
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contraparte nacional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fuente externa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label labTipoMoneda 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "labTipoMoneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label labPorcTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   27
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label labPorcNal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0%"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   26
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Height          =   3495
      Left            =   20
      TabIndex        =   0
      Top             =   0
      Width           =   990
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   720
         Left            =   120
         Picture         =   "ac_CapturaHojaVida.frx":0F00
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1440
         Width           =   765
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Grabar"
         Height          =   720
         Left            =   120
         Picture         =   "ac_CapturaHojaVida.frx":110A
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   600
         Width           =   780
      End
   End
   Begin MSAdodcLib.Adodc Ado_Pariente 
      Height          =   330
      Left            =   1080
      Top             =   3480
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_Pariente"
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
Attribute VB_Name = "ac_CapturaHojaVida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Para_Aceptado As String
Dim rs_pariente As New ADODB.Recordset

Dim nomb2 As String

Private Sub cmdCancel_Click()
'cancela la edicion de datos
Para_Aceptado = "N"
Me.Hide
End Sub

Private Sub cmdOk_Click()
'acepta las modificaciones realizadas
nomb2 = txtPat + " " + txtMat + " " + txtNom
If ValidaMontos Then
    Dim SQLS As String
    SQLS = ""
   If txtSW = "ADD" Then
      db.Execute "Insert INTO ro_Beneficiario_Dependiente (codigo_beneficiario, cod_dependiente, Cod_asegurado, Fecha_asegurado, fecha_nacimiento, primer_apellido, segundo_apellido, nombres, cod_pariente, nomb_pariente, estado_codigo, denominacion_beneficiario) Values ('" & txtBenef.Text & "', '" & txtCI.Text & "', '" & TxtItem.Text & "', '" & DTPFec_Seguro.Value & "', '" & txtNac.Value & "', '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', " & Dtc_Par.Text & ", '" & Dtc_ParDes.Text & "', '" & Txtestado.Text & "', '" & nomb2 & "')"
    '(txtBenef.Text , txtCI.Text , TxtItem.Text , DTPFec_Seguro.Value , txtNac.Value , txtPat.Text , txtMat.Text , txtNom.Text , Dtc_Par.Text , Dtc_ParDes.Text , txtEstado.Text , nomb2 )
    'DB.Execute "APEND ro_Beneficiario_Dependiente set codigo_beneficiario='" & txtBenef.Text & "', cod_dependiente='" & txtCI.Text & "', Cod_asegurado='" & TxtItem.Text & "', Fecha_asegurado=" & DTPFec_Seguro.Value & ", fecha_nacimiento=" & txtNac.Value & ", primer_apellido='" & txtPat.Text & "', segundo_apellido='" & txtMat.Text & "', nombres='" & txtNom.Text & "', cod_pariente=" & Dtc_Par.Text & ", nomb_pariente='" & Dtc_ParDes.Text & "', estado_registro='" & txtEstado.Text & "', denominacion_beneficiario='" & nomb2 & "'  "
   Else
      'DB.Execute "update ro_Beneficiario_Dependiente set codigo_beneficiario='" & txtBenef.Text & "', cod_dependiente='" & txtCI.Text & "', Cod_asegurado='" & TxtItem.Text & "', primer_apellido='" & txtPat.Text & "', segundo_apellido='" & txtMat.Text & "', nombres='" & txtNom.Text & "', cod_pariente=" & Dtc_Par.Text & ", nomb_pariente='" & Dtc_ParDes.Text & "', estado_registro='" & txtEstado.Text & "', denominacion_beneficiario='" & nomb2 & "'  "
      ' fecha_registro  hora_registro usr_usuario
      frmBeneficiario.AdoDependiente.Recordset("codigo_beneficiario").Value = txtBenef.Text
      frmBeneficiario.AdoDependiente.Recordset("cod_dependiente") = txtCI.Text
      frmBeneficiario.AdoDependiente.Recordset("Cod_asegurado").Value = TxtItem.Text
      frmBeneficiario.AdoDependiente.Recordset("primer_apellido").Value = txtPat.Text
      frmBeneficiario.AdoDependiente.Recordset("segundo_apellido").Value = txtMat.Text
      frmBeneficiario.AdoDependiente.Recordset("nombres").Value = txtNom.Text
      frmBeneficiario.AdoDependiente.Recordset("cod_pariente").Value = Dtc_Par.Text
      frmBeneficiario.AdoDependiente.Recordset("nomb_pariente").Value = Dtc_ParDes.Text
      frmBeneficiario.AdoDependiente.Recordset("estado_codigo").Value = Txtestado.Text
      frmBeneficiario.AdoDependiente.Recordset("denominacion_beneficiario").Value = nomb2
      
      frmBeneficiario.AdoDependiente.Recordset("Fecha_asegurado").Value = DTPFec_Seguro.Value
      frmBeneficiario.AdoDependiente.Recordset("fecha_nacimiento") = txtNac.Value
      frmBeneficiario.AdoDependiente.Recordset("usr_usuario").Value = glusuario 'frmLogin.txtUserName.Text
      frmBeneficiario.AdoDependiente.Recordset("fecha_registro").Value = Date
      frmBeneficiario.AdoDependiente.Recordset("hora_registro").Value = Format(Time, "hh:mm:ss")
      frmBeneficiario.AdoDependiente.Recordset.Update
   End If
   Para_Aceptado = "S"
   Me.Hide
End If
End Sub

Function ValidaMontos()
'valida que el monto asignado al beneficiario no sobrepase el monto pendiente de asignacion
ValidaMontos = True
'If Val(Me.mskMonto) > Val(Me.mskMonto_pendiente) Then
'    ValidaMontos = False
'    MsgBox "El monto indicado sobrepasa el monto pendiente de pago", vbInformation
'    Me.mskMonto.SelStart = 0
'    Me.mskMonto.SelLength = Len(Me.mskMonto)
'    Me.mskMonto.SetFocus
'End If
    If txtPat = "" Then
        ValidaMontos = False
    End If
End Function


Private Sub Dtc_Par_Click(Area As Integer)
    Dtc_ParDes.BoundText = Dtc_Par.BoundText
End Sub

Private Sub Dtc_ParDes_Click(Area As Integer)
    Dtc_Par.BoundText = Dtc_ParDes.BoundText
End Sub

Private Sub Form_Load()
If glProceso = "CONSULTORIA" Then
    Me.Caption = "Consultoría - Captura de datos personales"
Else
    Me.Caption = "Recursos Humanos - Captura de datos personales"
End If
Para_Aceptado = "N"
'LOS DATOS PERSONALES SE CARGAN EN EL FORMULARIO QUE LO LLAMA
'AQUI SE JALA LOS MONTOS REGISTRADOS EN AO_ADJUDICA_C
Dim Xmbe As Double, Xmde As Double, Xmbn As Double, Xmdn As Double
Dim XAbe As Double, XAde As Double, XAbn As Double, XAdn As Double
'With ac_Adjudicacion_c.adoSec.Recordset
'    Me.labTipoMoneda = !tipo_moneda
'    DE.dbo_edCmprSumaMontosLimiteBen1 !ges_gestion, !codigo_unidad, !codigo_solicitud, !numero_consultoria, Xmbe, Xmde, Xmbn, Xmdn, XAbe, XAde, XAbn, XAdn
'    If !tipo_moneda = "$US" Then
'        Me.mskMonto = Round(!monto_dolares_ext + !monto_dolares_nal, 2)
'        Me.mskMonto_ext = !monto_dolares_ext
'        Me.mskMonto_nal = !monto_dolares_nal
'        Me.mskMonto_limite = Xmde + Xmdn
'        Me.mskMonto_pendiente = Round(Xmde + Xmdn - XAde - XAdn + Val(Me.mskMonto), 2)
'        Me.labPorcExt = CStr(Format(Xmde / Val(Me.mskMonto_limite) * 100, "##0.00")) & "%"
'        Me.labPorcNal = CStr(Format(Xmdn / Val(Me.mskMonto_limite) * 100, "##0.00")) & "%"
'        Me.mskMonto = Round(!monto_dolares_ext + !monto_dolares_nal, 2)
'    Else
'        Me.mskMonto = Round(!monto_bolivianos_ext + !monto_bolivianos_nal)
'        Me.mskMonto_ext = !monto_bolivianos_ext
'        Me.mskMonto_nal = !monto_bolivianos_nal
'        Me.mskMonto_limite = Xmbe + Xmbn
'        Me.mskMonto_pendiente = Xmbe + Xmbn - XAbe - XAbn + Val(Me.mskMonto)
'        If Val(Me.mskMonto_limite) = 0 Then
'            Me.labPorcExt = "0 %"
'            Me.labPorcNal = "0 %"
'        Else
'            Me.labPorcExt = CStr(Format(Xmbe / Val(Me.mskMonto_limite) * 100, "##0.00")) & "%"
'            Me.labPorcNal = CStr(Format(Xmbn / Val(Me.mskMonto_limite) * 100, "##0.00")) & "%"
'        End If
'        Me.mskMonto = Round(!monto_bolivianos_ext + !monto_bolivianos_nal)
'    End If
'End With
    
    Set rs_pariente = New ADODB.Recordset
    rs_pariente.Open "SELECT * FROM rc_beneficiario_pariente ORDER BY nomb_pariente ", db, adOpenStatic
    Set Ado_Pariente.Recordset = rs_pariente
    
If Val(Me.mskMonto_limite) = 0 Then
    Me.labPorcExt = "0%"
    Me.labPorcNal = "0%"
End If
'mskMonto.SetFocus
End Sub

Private Sub mskMonto_Change()
    Call DivideXFte
End Sub

Sub DivideXFte()
'divide el monto total en montos correspondientes alos porcentajes
'externo y contraparte nacional
Me.mskMonto_ext = Round(Val(Me.mskMonto) * Val(Left(Me.labPorcExt, Len(Me.labPorcExt) - 1)) / 100, 2)
Me.mskMonto_nal = Round(Val(Me.mskMonto) - Val(Me.mskMonto_ext), 2)
End Sub

Private Sub mskMonto_ext_GotFocus()
Me.mskMonto.SetFocus
End Sub

Private Sub mskMonto_GotFocus()
mskMonto.SelStart = 0
mskMonto.SelLength = Len(mskMonto)
End Sub

Private Sub mskMonto_KeyPress(KeyAscii As Integer)
If Val(Chr(KeyAscii)) <> 0 Or Chr(KeyAscii) = "." Or Chr(KeyAscii) = "," Or Chr(KeyAscii) = "0" Or KeyAscii = 8 Then
    'asdfasdf
Else
    KeyAscii = 0
End If
End Sub

Private Sub mskMonto_limite_GotFocus()
Me.mskMonto.SetFocus
End Sub

Private Sub mskMonto_nal_GotFocus()
Me.mskMonto.SetFocus
End Sub

Private Sub mskMonto_pendiente_GotFocus()
Me.mskMonto.SetFocus
End Sub


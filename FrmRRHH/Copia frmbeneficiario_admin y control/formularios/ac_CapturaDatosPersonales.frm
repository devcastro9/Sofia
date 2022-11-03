VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ac_CapturaDatosPersonales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion de Personal - Ficha Personal - Dependientes del Funcionario"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7755
   ControlBox      =   0   'False
   Icon            =   "ac_CapturaDatosPersonales.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ac_CapturaDatosPersonales.frx":0ECA
   ScaleHeight     =   4785
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "ac_CapturaDatosPersonales.frx":6CEFC
      ScaleHeight     =   915
      ScaleWidth      =   7515
      TabIndex        =   43
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "ac_CapturaDatosPersonales.frx":D8F2E
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   240
         Picture         =   "ac_CapturaDatosPersonales.frx":D9138
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEPENDIENTES DEL FUNCIONARIO"
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
         Left            =   2100
         TabIndex        =   46
         Top             =   240
         Width           =   5355
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
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
      Height          =   3615
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   7575
      Begin VB.TextBox TxtOcupacion 
         DataField       =   "ocupacion_pariente"
         DataSource      =   "frmBeneficiario_admin.adoLista"
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   9
         Top             =   3140
         Width           =   4935
      End
      Begin VB.TextBox TxtItem 
         DataField       =   "Cod_asegurado"
         DataSource      =   "frmBeneficiario_admin.adoLista"
         Height          =   285
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtBenef 
         Height          =   285
         Left            =   600
         MaxLength       =   80
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         MaxLength       =   80
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtTelefono 
         Height          =   285
         Left            =   6600
         MaxLength       =   20
         TabIndex        =   19
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtRecomendaciones 
         Height          =   285
         Left            =   6600
         MaxLength       =   300
         TabIndex        =   17
         Top             =   2400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtPat 
         DataField       =   "primer_apellido"
         DataSource      =   "frmBeneficiario_admin.adoLista"
         Height          =   285
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1680
         Width           =   4935
      End
      Begin VB.TextBox txtMat 
         DataField       =   "segundo_apellido"
         DataSource      =   "frmBeneficiario_admin.adoLista"
         Height          =   285
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2040
         Width           =   4935
      End
      Begin VB.TextBox txtNom 
         DataField       =   "nombres"
         DataSource      =   "frmBeneficiario_admin.adoLista"
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2400
         Width           =   4935
      End
      Begin VB.TextBox txtCI 
         DataField       =   "cod_dependiente"
         DataSource      =   "frmBeneficiario_admin.adoLista"
         Height          =   285
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   0
         Top             =   400
         Width           =   1815
      End
      Begin VB.ComboBox cboTDoc 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComCtl2.DTPicker txtNac 
         DataField       =   "fecha_nacimiento"
         DataSource      =   "frmBeneficiario_admin.adoLista"
         Height          =   315
         Left            =   5520
         TabIndex        =   3
         Top             =   405
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   91095041
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker DTPFec_Seguro 
         DataField       =   "Fecha_asegurado"
         DataSource      =   "frmBeneficiario_admin.adoLista"
         Height          =   315
         Left            =   5520
         TabIndex        =   2
         Top             =   1080
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   91095041
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo Dtc_Par 
         Bindings        =   "ac_CapturaDatosPersonales.frx":D9342
         DataField       =   "pariente_codigo"
         DataSource      =   "frmBeneficiario_admin.adoLista"
         Height          =   315
         Left            =   5640
         TabIndex        =   8
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "pariente_codigo"
         BoundColumn     =   "pariente_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Dtc_ParDes 
         Bindings        =   "ac_CapturaDatosPersonales.frx":D935D
         DataField       =   "pariente_codigo"
         DataSource      =   "frmBeneficiario_admin.adoLista"
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   2760
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "pariente_descripcion"
         BoundColumn     =   "pariente_codigo"
         Text            =   ""
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Ocupación"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   675
         TabIndex        =   42
         Top             =   3120
         Width           =   780
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Codigo Seguro Social (Matrícula)"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   11
         Left            =   120
         TabIndex        =   41
         Top             =   1035
         Width           =   1410
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Parentesco"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   645
         TabIndex        =   22
         Top             =   2760
         Width           =   810
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Fecha Nacimiento "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   3960
         TabIndex        =   21
         Top             =   480
         Width           =   1530
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         Caption         =   "Aprobado"
         Height          =   195
         Index           =   6
         Left            =   2520
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Segundo apellido"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   16
         Top             =   2040
         Width           =   1230
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nombres"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   825
         TabIndex        =   15
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Primer apellido"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   435
         TabIndex        =   14
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Nro. Documento de Identidad "
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Fecha Vencimiento"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   3900
         TabIndex        =   12
         Top             =   1155
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Presupuesto requerido"
      ForeColor       =   &H00808000&
      Height          =   2895
      Left            =   600
      TabIndex        =   25
      Top             =   1320
      Width           =   5535
      Begin MSMask.MaskEdBox mskMonto_pendiente 
         Height          =   375
         Left            =   3720
         TabIndex        =   26
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
         TabIndex        =   27
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
         TabIndex        =   28
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
         TabIndex        =   29
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
         TabIndex        =   30
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   1440
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Ado_Pariente 
      Height          =   330
      Left            =   960
      Top             =   4440
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
Attribute VB_Name = "ac_CapturaDatosPersonales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Para_Aceptado As String
Dim rs_pariente As New ADODB.Recordset

Dim nomb2 As String

Private Sub BtnCancelar_Click()
'cancela la edicion de datos
Para_Aceptado = "N"
Me.Hide
End Sub

Private Sub BtnGrabar_Click()
'acepta las modificaciones realizadas
nomb2 = txtPat + " " + txtMat + " " + txtNom
If ValidaMontos Then
    Dim SQLS As String
    SQLS = ""
   If txtSW = "ADD" Then
      db.Execute "Insert INTO ro_Beneficiario_Dependiente (beneficiario_codigo, cod_dependiente, Cod_asegurado, Fecha_asegurado, fecha_nacimiento, primer_apellido, segundo_apellido, nombres, pariente_codigo, pariente_descripcion, estado_codigo, denominacion_beneficiario, ocupacion_pariente) Values ('" & txtBenef.Text & "', '" & txtCI.Text & "', '" & TxtItem.Text & "', '" & DTPFec_Seguro.Value & "', '" & txtNac.Value & "', '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', " & Dtc_Par.Text & ", '" & Dtc_ParDes.Text & "', '" & txtEstado.Text & "', '" & nomb2 & "', '" & TxtOcupacion & "')"
      frmBeneficiario_Admin.abrirtabla
   Else
      'DB.Execute "update ro_Beneficiario_Dependiente set beneficiario_codigo='" & txtBenef.Text & "', cod_dependiente='" & txtCI.Text & "', Cod_asegurado='" & TxtItem.Text & "', primer_apellido='" & txtPat.Text & "', segundo_apellido='" & txtMat.Text & "', nombres='" & txtNom.Text & "', pariente_codigo=" & Dtc_Par.Text & ", pariente_descripcion='" & Dtc_ParDes.Text & "', estado_codigo='" & txtEstado.Text & "', beneficiario_denominacion='" & nomb2 & "'  "
      ' fecha_registro  hora_registro usr_usuario
      
      frmBeneficiario_Admin.AdoDependiente.Recordset("beneficiario_codigo").Value = txtBenef.Text
      frmBeneficiario_Admin.AdoDependiente.Recordset("cod_dependiente") = txtCI.Text
      frmBeneficiario_Admin.AdoDependiente.Recordset("Cod_asegurado") = TxtItem.Text
      frmBeneficiario_Admin.AdoDependiente.Recordset("primer_apellido").Value = txtPat.Text
      frmBeneficiario_Admin.AdoDependiente.Recordset("segundo_apellido").Value = txtMat.Text
      frmBeneficiario_Admin.AdoDependiente.Recordset("nombres").Value = txtNom.Text
      frmBeneficiario_Admin.AdoDependiente.Recordset("pariente_codigo").Value = Dtc_Par.Text
      frmBeneficiario_Admin.AdoDependiente.Recordset("pariente_descripcion").Value = Dtc_ParDes.Text
      frmBeneficiario_Admin.AdoDependiente.Recordset("estado_codigo").Value = txtEstado.Text
      frmBeneficiario_Admin.AdoDependiente.Recordset("denominacion_beneficiario").Value = nomb2
      frmBeneficiario_Admin.AdoDependiente.Recordset("ocupacion_pariente").Value = TxtOcupacion.Text

      frmBeneficiario_Admin.AdoDependiente.Recordset("Fecha_asegurado").Value = DTPFec_Seguro.Value
      frmBeneficiario_Admin.AdoDependiente.Recordset("fecha_nacimiento") = txtNac.Value
      frmBeneficiario_Admin.AdoDependiente.Recordset("usr_codigo").Value = glusuario 'frmLogin.txtUserName.Text
      frmBeneficiario_Admin.AdoDependiente.Recordset("fecha_registro").Value = Date
      frmBeneficiario_Admin.AdoDependiente.Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
      frmBeneficiario_Admin.AdoDependiente.Recordset.Update
      frmBeneficiario_Admin.abrirtabla
   End If
   Para_Aceptado = "S"
   'frmBeneficiario.AdoDependiente.Refresh '.Recordset.Requery
   Unload Me
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
    If txtNom = "" Then
        ValidaMontos = False
    End If
    
    If TxtItem = "" Then
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
    rs_pariente.Open "SELECT * FROM rc_beneficiario_pariente ORDER BY pariente_descripcion ", db, adOpenStatic
    Set Ado_Pariente.Recordset = rs_pariente
    
If Val(Me.mskMonto_limite) = 0 Then
    Me.labPorcExt = "0%"
    Me.labPorcNal = "0%"
End If
'mskMonto.SetFocus
	Call SeguridadSet(Me)
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


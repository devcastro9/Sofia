VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ao_Asistencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Personal - File Funcionario - Control Asistencia"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7755
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_ao_Asistencia.frx":0000
   ScaleHeight     =   5160
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_ao_Asistencia.frx":6A41E
      ScaleHeight     =   915
      ScaleWidth      =   7515
      TabIndex        =   28
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   240
         Picture         =   "frm_ao_Asistencia.frx":D6450
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "frm_ao_Asistencia.frx":D665A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTROL DE ASISTENCIA"
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
         Left            =   2760
         TabIndex        =   31
         Top             =   240
         Width           =   4035
      End
   End
   Begin VB.Frame FraEmpresa 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00008000&
      Height          =   4020
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7560
      Begin VB.ComboBox TxtGestion 
         DataField       =   "ges_gestion"
         DataSource      =   "frmBeneficiario_control.AdoAsistencia"
         Enabled         =   0   'False
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
         ItemData        =   "frm_ao_Asistencia.frx":D6864
         Left            =   240
         List            =   "frm_ao_Asistencia.frx":D6874
         TabIndex        =   23
         Text            =   "2011"
         Top             =   3480
         Width           =   900
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         MaxLength       =   80
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtBenef 
         Height          =   285
         Left            =   720
         MaxLength       =   80
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox txt01 
         Enabled         =   0   'False
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
         ItemData        =   "frm_ao_Asistencia.frx":D6890
         Left            =   240
         List            =   "frm_ao_Asistencia.frx":D68B8
         TabIndex        =   6
         Text            =   "ENERO"
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox Txt02 
         Enabled         =   0   'False
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
         ItemData        =   "frm_ao_Asistencia.frx":D6921
         Left            =   5520
         List            =   "frm_ao_Asistencia.frx":D693A
         TabIndex        =   5
         Text            =   "LUNES"
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox txt05 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         ItemData        =   "frm_ao_Asistencia.frx":D697A
         Left            =   5160
         List            =   "frm_ao_Asistencia.frx":D6984
         TabIndex        =   4
         Text            =   "NO"
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox Txt06 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
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
         ItemData        =   "frm_ao_Asistencia.frx":D6990
         Left            =   6360
         List            =   "frm_ao_Asistencia.frx":D699D
         TabIndex        =   3
         Text            =   "AST"
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox Cmb01 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         ItemData        =   "frm_ao_Asistencia.frx":D69B0
         Left            =   5160
         List            =   "frm_ao_Asistencia.frx":D69BA
         TabIndex        =   2
         Text            =   "NO"
         Top             =   2640
         Width           =   735
      End
      Begin VB.ComboBox Cmb02 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
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
         ItemData        =   "frm_ao_Asistencia.frx":D69C6
         Left            =   6360
         List            =   "frm_ao_Asistencia.frx":D69D3
         TabIndex        =   1
         Text            =   "AST"
         Top             =   2640
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPFec_Inicio 
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   84475905
         CurrentDate     =   40909
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txt03 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "hh:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   4
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   84475906
         UpDown          =   -1  'True
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker txt04 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "hh:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   4
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   84475906
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker txt07 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "hh:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   4
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   84475906
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker txt08 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "hh:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   4
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   11
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   84475906
         CurrentDate     =   36494
      End
      Begin MSDataListLib.DataCombo Dtc_Par 
         Bindings        =   "frm_ao_Asistencia.frx":D69E6
         DataField       =   "turno"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "hh:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   4
         EndProperty
         DataSource      =   "frmBeneficiario_control.AdoAsistencia"
         Height          =   315
         Left            =   2040
         TabIndex        =   20
         Top             =   3480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "tope_hora_ingreso"
         BoundColumn     =   "turno"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo Dtc_ParDes 
         Bindings        =   "frm_ao_Asistencia.frx":D6A05
         DataField       =   "turno"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "frmBeneficiario_control.AdoAsistencia"
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "turno"
         BoundColumn     =   "turno"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DtcFec_Fin 
         Height          =   315
         Left            =   5760
         TabIndex        =   22
         Top             =   3480
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   84475905
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo Dtc_Par2 
         Bindings        =   "frm_ao_Asistencia.frx":D6A24
         DataField       =   "turno2"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   4
         EndProperty
         DataSource      =   "frmBeneficiario_control.AdoAsistencia"
         Height          =   315
         Left            =   4320
         TabIndex        =   24
         Top             =   3480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "tope_hora_ingreso"
         BoundColumn     =   "turno"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo Dtc_Par3 
         Bindings        =   "frm_ao_Asistencia.frx":D6A44
         DataField       =   "turno2"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "frmBeneficiario_control.AdoAsistencia"
         Height          =   315
         Left            =   240
         TabIndex        =   25
         Top             =   2640
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "turno"
         BoundColumn     =   "turno"
         Text            =   ""
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Gestión:                      Hora Limite Ctrl.1:                  Hora Limite Ctrl.2:"
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
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   3240
         Width           =   5625
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:              Hora Ingreso 2:          Hora de Salida 2:        Atraso 2:    Asist-Lic-Falta"
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
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   2400
         Width           =   7155
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         Caption         =   "Aprobado"
         Height          =   195
         Index           =   6
         Left            =   2520
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         Caption         =   "SW"
         Height          =   195
         Index           =   10
         Left            =   1440
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         Caption         =   "Benef"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:              Hora Ingreso 1:          Hora de Salida 1:        Atraso 1:     Asist-Lic-Falta"
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
         Height          =   240
         Index           =   40
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   7200
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes de Control:                                Fecha de Control:                    Día de Control:"
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
         Height          =   240
         Index           =   45
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   6630
      End
   End
   Begin MSAdodcLib.Adodc Ado_Clasificador 
      Height          =   330
      Left            =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Ado_Clasificador"
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
   Begin MSAdodcLib.Adodc Ado_Clasificador2 
      Height          =   330
      Left            =   2640
      Top             =   5040
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Ado_Clasificador"
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
Attribute VB_Name = "frm_ao_Asistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Para_Aceptado As String
Dim rs_Clasificador As New ADODB.Recordset
Dim rs_Clasificador2 As New ADODB.Recordset

Dim nomb2 As String
Dim hora01, hora02, hora03, hora04 As String
Dim fecha1 As String
'dim hora00 as Date = #HH:mm:ss#

Private Sub cmdCancel_Click()
    'cancela la edicion de datos
    Para_Aceptado = "N"
    Unload Me
    'Me.Hide
End Sub

Private Sub cmdOk_Click()
 'acepta las modificaciones realizadas
 Dim ctrl1, ctrl2  As Integer
 Dim Atr1, Atr2  As Integer
 Dim dia2, mes2 As String
 If ValidaMontos Then
   Dim SQLS As String
   SQLS = ""
   If txtSW = "ADD" Then
      hora01 = Format(txt03.Value, "HH:mm:ss")
      hora02 = Format(txt04.Value, "HH:mm:ss")
      hora03 = Format(txt07.Value, "HH:mm:ss")
      hora04 = Format(txt08.Value, "HH:mm:ss")
      db.Execute "Insert INTO ro_ControlAsistencia (beneficiario_codigo, Fecha_control, mes_control, dia_control, HoraUno, HoraDos, Atraso, Falta, HoraTres, HoraCuatro, AtrasoI, Falta2, estado_codigo, fecha_registro, usr_usuario) Values ('" & txtBenef.Text & "', '" & DTPFec_Inicio.Value & "', '" & txt01 & "', '" & Txt02 & "', '" & hora01 & "', '" & hora02 & "', '" & txt05 & "', '" & Txt06 & "', '" & hora03 & "', '" & hora04 & "', '" & Cmb01 & "', '" & Cmb02 & "', 'NO', '" & Date & "', '" & GlUsuario & "') "
      frmBeneficiario_Control.AdoAsistencia.Recordset("turno").Value = "AM"
      frmBeneficiario_Control.AdoAsistencia.Recordset("turno2").Value = "PM"
      frmBeneficiario_Control.AdoAsistencia.Recordset("Fecha_control").Value = DTPFec_Inicio.Value
      frmBeneficiario_Control.AdoAsistencia.Recordset("ges_gestion").Value = Year(DTPFec_Inicio.Value)
      dia2 = WeekdayName(Weekday(DTPFec_Inicio.Value))
      mes2 = MonthName(Month(DTPFec_Inicio.Value))
      frmBeneficiario_Control.AdoAsistencia.Recordset("mes_control") = mes2 'Txt01.Text
      frmBeneficiario_Control.AdoAsistencia.Recordset("dia_control").Value = dia2   'txt02.Text
   Else
      frmBeneficiario_Control.AdoAsistencia.Recordset("HoraUno").Value = Format(txt03.Value, "HH:mm:ss")
      frmBeneficiario_Control.AdoAsistencia.Recordset("HoraDos").Value = Format(txt04.Value, "HH:mm:ss")
      ctrl1 = DateDiff("n", frmBeneficiario_Control.AdoAsistencia.Recordset("HoraUno").Value, frmBeneficiario_Control.AdoAsistencia.Recordset("HoraDos").Value)
      Atr1 = DateDiff("n", CDate(GlHora1), frmBeneficiario_Control.AdoAsistencia.Recordset("HoraUno").Value)
      If Atr1 > 0 Then
        frmBeneficiario_Control.AdoAsistencia.Recordset("Atraso").Value = "SI"
        frmBeneficiario_Control.AdoAsistencia.Recordset("Falta").Value = "AST"
        frmBeneficiario_Control.AdoAsistencia.Recordset("AtrasoMin1").Value = Atr1
        'frmBeneficiario_Control.AdoAsistencia.Recordset("AtrasoMin1").Value = DateDiff("n", frmBeneficiario_Control.AdoAsistencia.Recordset("HoraDos").Value, frmBeneficiario_Control.AdoAsistencia.Recordset("HoraUno").Value)
      Else
        frmBeneficiario_Control.AdoAsistencia.Recordset("Atraso").Value = "NO"
        frmBeneficiario_Control.AdoAsistencia.Recordset("Falta").Value = "AST"
        frmBeneficiario_Control.AdoAsistencia.Recordset("AtrasoMin1").Value = 0
      End If
      frmBeneficiario_Control.AdoAsistencia.Recordset("TotalMin1").Value = ctrl1
      'frmBeneficiario_Control.AdoAsistencia.Recordset("TotalMin1").Value = DateDiff("n", frmBeneficiario_Control.AdoAsistencia.Recordset("HoraDos").Value, frmBeneficiario_Control.AdoAsistencia.Recordset("HoraUno").Value)
      'ctrl1 = DateDiff("mi", Format(txt03.Value, "HH:mm:ss"), Format(txt04.Value, "HH:mm:ss"))
      '480 min = 4 hrs
      'ctrl1 = DateDiff("n", frmBeneficiario_Control.AdoAsistencia.Recordset("HoraDos").Value, frmBeneficiario_Control.AdoAsistencia.Recordset("HoraUno").Value)
      frmBeneficiario_Control.AdoAsistencia.Recordset("HoraTres").Value = Format(txt07.Value, "HH:mm:ss")
      frmBeneficiario_Control.AdoAsistencia.Recordset("HoraCuatro").Value = Format(txt08.Value, "HH:mm:ss")
      ctrl2 = DateDiff("n", frmBeneficiario_Control.AdoAsistencia.Recordset("HoraTres").Value, frmBeneficiario_Control.AdoAsistencia.Recordset("HoraCuatro").Value)
      Atr2 = DateDiff("n", CDate(GlHora2), frmBeneficiario_Control.AdoAsistencia.Recordset("HoraTres").Value)
      If Atr2 > 0 Then
        frmBeneficiario_Control.AdoAsistencia.Recordset("AtrasoI").Value = "SI"
        frmBeneficiario_Control.AdoAsistencia.Recordset("Falta2").Value = "AST"
        frmBeneficiario_Control.AdoAsistencia.Recordset("AtrasoMin2").Value = Atr2
        'frmBeneficiario_Control.AdoAsistencia.Recordset("AtrasoMin1").Value = DateDiff("n", frmBeneficiario_Control.AdoAsistencia.Recordset("HoraDos").Value, frmBeneficiario_Control.AdoAsistencia.Recordset("HoraUno").Value)
      Else
        frmBeneficiario_Control.AdoAsistencia.Recordset("AtrasoI").Value = "NO"
        frmBeneficiario_Control.AdoAsistencia.Recordset("Falta2").Value = "AST"
        frmBeneficiario_Control.AdoAsistencia.Recordset("AtrasoMin2").Value = 0
      End If
      frmBeneficiario_Control.AdoAsistencia.Recordset("TotalMin2").Value = ctrl2
      
      frmBeneficiario_Control.AdoAsistencia.Recordset("Hora1").Value = Hour(frmBeneficiario_Control.AdoAsistencia.Recordset("HoraUno").Value)
      frmBeneficiario_Control.AdoAsistencia.Recordset("Hora2").Value = Hour(frmBeneficiario_Control.AdoAsistencia.Recordset("HoraDos").Value)
      frmBeneficiario_Control.AdoAsistencia.Recordset("Hora3").Value = Hour(frmBeneficiario_Control.AdoAsistencia.Recordset("HoraTres").Value)
      frmBeneficiario_Control.AdoAsistencia.Recordset("Hora4").Value = Hour(frmBeneficiario_Control.AdoAsistencia.Recordset("HoraCuatro").Value)
      'ctrl1 = frmBeneficiario_Control.AdoAsistencia.Recordset("Hora2").Value - frmBeneficiario_Control.AdoAsistencia.Recordset("Hora1").Value
      'ctrl1 = DateDiff(mi, frmBeneficiario_Control.AdoAsistencia.Recordset("HoraDos").Value, frmBeneficiario_Control.AdoAsistencia.Recordset("HoraUno").Value)
'      frmBeneficiario_Control.AdoAsistencia.Recordset("Min1").Value = Hour(frmBeneficiario_Control.AdoAsistencia.Recordset("HoraUno").Value)
'      frmBeneficiario_Control.AdoAsistencia.Recordset("Min2").Value = Hour(frmBeneficiario_Control.AdoAsistencia.Recordset("HoraDos").Value)
'      frmBeneficiario_Control.AdoAsistencia.Recordset("Min3").Value = Hour(frmBeneficiario_Control.AdoAsistencia.Recordset("HoraTres").Value)
'      frmBeneficiario_Control.AdoAsistencia.Recordset("Min4").Value = Hour(frmBeneficiario_Control.AdoAsistencia.Recordset("HoraCuatro").Value)
'      ctrl2 = frmBeneficiario_Control.AdoAsistencia.Recordset("Min2").Value - frmBeneficiario_Control.AdoAsistencia.Recordset("Min1").Value
      frmBeneficiario_Control.AdoAsistencia.Recordset("turno").Value = Dtc_ParDes.Text
      frmBeneficiario_Control.AdoAsistencia.Recordset("turno2").Value = Dtc_Par3.Text
      frmBeneficiario_Control.AdoAsistencia.Recordset("estado_codigo").Value = IIf(Txtestado.Text = "", "NO", Txtestado.Text)
      frmBeneficiario_Control.AdoAsistencia.Recordset("fecha_registro") = Date
      frmBeneficiario_Control.AdoAsistencia.Recordset("usr_usuario").Value = GlUsuario
      frmBeneficiario_Control.AdoAsistencia.Recordset("hora_registro").Value = Format(Time, "HH:mm:ss")
      frmBeneficiario_Control.AdoAsistencia.Recordset.Update
   End If
   Para_Aceptado = "S"
   Unload Me
   'Me.Hide
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
    If txt01 = "" Then
        ValidaMontos = False
    End If
    If Txt02 = "" Then
        ValidaMontos = False
    End If
    If txt03 = "" Then
        ValidaMontos = False
    End If
    If txt04 = "" Then
        ValidaMontos = False
    End If
End Function

Private Sub Dtc_Par_Click(Area As Integer)
    Dtc_ParDes.BoundText = Dtc_Par.BoundText
End Sub

Private Sub Dtc_Par2_Click(Area As Integer)
    Dtc_Par2.BoundText = Dtc_Par3.BoundText
End Sub

Private Sub Dtc_Par3_Click(Area As Integer)
    Dtc_Par2.BoundText = Dtc_Par3.BoundText
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
   Set rs_Clasificador = New ADODB.Recordset
   Set rs_Clasificador2 = New ADODB.Recordset
   'If txtSW = "ADD" Then
      rs_Clasificador.Open "SELECT * FROM rc_horarios ", db, adOpenStatic
      rs_Clasificador2.Open "SELECT * FROM rc_horarios ", db, adOpenStatic
   'Else
   '   'rs_Clasificador.Open "SELECT * FROM rc_horarios WHERE Dia_control = '" & Trim(Txt02) & "' ", DB, adOpenStatic
   '   rs_Clasificador.Open "SELECT * FROM rc_horarios WHERE TURNO = 'AM' ", DB, adOpenStatic
   '   rs_Clasificador2.Open "SELECT * FROM rc_horarios WHERE TURNO = 'PM' ", DB, adOpenStatic
   'End If
   Set Ado_Clasificador.Recordset = rs_Clasificador
   Set Ado_Clasificador2.Recordset = rs_Clasificador2
   Dtc_Par.BoundText = Dtc_ParDes.BoundText
   Dtc_Par2.BoundText = Dtc_Par3.BoundText
'mskMonto.SetFocus
	Call SeguridadSet(Me)
End Sub

Private Sub Txt02_Change()
    Dtc_ParDes.Text = Txt02.Text
End Sub

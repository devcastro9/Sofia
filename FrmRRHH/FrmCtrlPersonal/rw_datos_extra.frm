VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rw_datos_extra 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7200
   ClientLeft      =   1065
   ClientTop       =   -30
   ClientWidth     =   7305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      FillStyle       =   2  'Horizontal Line
      ForeColor       =   &H80000008&
      Height          =   676
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   10920
      TabIndex        =   21
      Top             =   0
      Width           =   10920
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2400
         Picture         =   "rw_datos_extra.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   23
         Top             =   0
         Width           =   1335
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3675
         Picture         =   "rw_datos_extra.frx":07D6
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   22
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENTAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   13215
         TabIndex        =   24
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame SWS 
      BackColor       =   &H00C0C0C0&
      Height          =   6375
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   7095
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   6885
         Begin MSDataListLib.DataCombo dtc_desc 
            Bindings        =   "rw_datos_extra.frx":10C2
            DataField       =   "codigo_empresa"
            DataSource      =   "rw_ficha_rrhh.Ado_datos"
            Height          =   315
            Left            =   240
            TabIndex        =   38
            Top             =   600
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "denominacion_empresa"
            BoundColumn     =   "codigo_empresa"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_cod 
            Bindings        =   "rw_datos_extra.frx":10DC
            DataField       =   "codigo_empresa"
            DataSource      =   "rw_ficha_rrhh.Ado_datos"
            Height          =   315
            Left            =   5160
            TabIndex        =   39
            Top             =   600
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "codigo_empresa"
            BoundColumn     =   "codigo_empresa"
            Text            =   ""
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Empresa a la que pertenece"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   2955
         End
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H000000FF&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         TabIndex        =   35
         Top             =   4920
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000FF00&
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   34
         Top             =   4920
         Width           =   855
      End
      Begin VB.Frame fra_doc 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   120
         TabIndex        =   19
         Top             =   5160
         Width           =   6885
         Begin VB.TextBox txt_nuevo_num 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   240
            MaxLength       =   15
            TabIndex        =   30
            Top             =   600
            Width           =   2805
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nuevo Numero Documento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   2790
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos Complementarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   1935
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   6855
         Begin VB.ComboBox cmb_tutor 
            Height          =   315
            ItemData        =   "rw_datos_extra.frx":10F5
            Left            =   240
            List            =   "rw_datos_extra.frx":10FF
            TabIndex        =   32
            Top             =   1440
            Width           =   2535
         End
         Begin VB.ComboBox cmb_discapacidad 
            Height          =   315
            ItemData        =   "rw_datos_extra.frx":110B
            Left            =   240
            List            =   "rw_datos_extra.frx":1115
            TabIndex        =   31
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   7440
            TabIndex        =   20
            Top             =   540
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   9600
            TabIndex        =   18
            Top             =   540
            Width           =   360
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2865
            TabIndex        =   17
            Top             =   2115
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   6345
            TabIndex        =   15
            Top             =   2115
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   9945
            TabIndex        =   14
            Top             =   2115
            Width           =   255
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   9840
            TabIndex        =   12
            Top             =   1320
            Visible         =   0   'False
            Width           =   360
         End
         Begin MSDataListLib.DataCombo Txt_campo2 
            DataField       =   "bien_codigo"
            DataSource      =   "fw_compras_gral.ado_detalle1"
            Height          =   315
            Left            =   10320
            TabIndex        =   16
            Top             =   2640
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "marca_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "¿Tutor de Persona Con Discapacidad?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   13
            Top             =   1140
            Width           =   4020
         End
         Begin VB.Label lbl_descripcion 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "¿Persona con Discapacidad?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   3060
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cambiar Numero documento ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   33
         Top             =   4920
         Width           =   3240
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Documeto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   1710
      End
      Begin VB.Label txt_tipo_doc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataSource      =   "Ado_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1965
         TabIndex        =   27
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label txt_ext 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataSource      =   "Ado_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4680
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Txt_descripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataSource      =   "Ado_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   8
         Left            =   285
         TabIndex        =   25
         Top             =   240
         Width           =   960
      End
      Begin VB.Label txt_codigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataSource      =   "Ado_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1920
         TabIndex        =   8
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lbl_codigo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nro. Documeto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2430
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   7305
      TabIndex        =   0
      Top             =   7200
      Width           =   7305
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_empresa 
      Height          =   330
      Left            =   120
      Top             =   7200
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
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
      Caption         =   "Ado_datos_busq"
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
Attribute VB_Name = "rw_datos_extra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos1 As New ADODB.Recordset
Attribute rs_datos1.VB_VarHelpID = -1
Dim rs_Depto As New ADODB.Recordset
Dim rs_TipoDocId As New ADODB.Recordset
Dim rs_empresa As New ADODB.Recordset

'BUSCADOR
Dim cambiar, tutor, discapacidad As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnCancelar_Click()
   On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        aw_p_ao_solicitud.Ado_detalle1.Recordset.CancelUpdate
        Unload Me
    End If
End Sub

Private Sub BtnGrabar_Click()
 On Error GoTo UpdateErr
 
     If cmb_discapacidad.Text = "SI" Then
        discapacidad = "1"
     Else
        discapacidad = "0"
     End If
     
     If cmb_tutor.Text = "SI" Then
        tutor = "1"
     Else
        tutor = "0"
     End If
    'rw_ficha_rrhh.Ado_datos.Recordset!tutor = tutor
    'rw_ficha_rrhh.Ado_datos.Recordset!discapacidad = discapacidad
     db.Execute "update ro_personal_contratado set tutor = '" & tutor & "' where beneficiario_codigo = '" & txt_codigo & "'"
     db.Execute "update ro_personal_contratado set discapacidad = '" & discapacidad & "' where beneficiario_codigo = '" & txt_codigo & "'"
     db.Execute "update ro_personal_contratado set codigo_empresa = '" & dtc_cod.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
     If cambiar = "SI" Then
     
        db.Execute "update gc_beneficiario set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ro_ControlAsistencia set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ro_liquidaciones set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ro_memorandas set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ro_movilidad_personal set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ro_pagos_cronograma_Detalle set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ro_permisos set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ro_permisos_detalle set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ro_personal_contratado set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ro_prestamo_prog set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ro_prestamos set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ro_retroactivo_aux set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ro_rrhh_adjudica_personas set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ro_vacaciones_programadas set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update ao_solicitud set beneficiario_codigo_resp = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_resp = '" & txt_codigo & "'"
        db.Execute "update ao_ventas_cabecera set beneficiario_codigo_resp = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_resp = '" & txt_codigo & "'"
        db.Execute "update ao_ventas_cabecera set beneficiario_codigo_cobr = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_cobr = '" & txt_codigo & "'"
        db.Execute "update ao_ventas_cabecera set beneficiario_codigo_alm = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_alm = '" & txt_codigo & "'"
        db.Execute "update ao_ventas_cabecera set beneficiario_codigo_almR = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_almR = '" & txt_codigo & "'"
        db.Execute "update ao_ventas_cabecera set beneficiario_codigo_almH = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_almH = '" & txt_codigo & "'"
        db.Execute "update ao_ventas_cabecera set beneficiario_codigo_tec = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_tec = '" & txt_codigo & "'"
        db.Execute "update ao_ventas_cabecera set beneficiario_codigo_tecR = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_tecR = '" & txt_codigo & "'"
        db.Execute "update ao_ventas_cabecera set beneficiario_codigo_tecH = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_tecH = '" & txt_codigo & "'"
        db.Execute "update ao_ventas_cobranza set beneficiario_codigo_resp = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_resp = '" & txt_codigo & "'"
        db.Execute "update ao_ventas_cobranza_det set beneficiario_codigo_resp = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_resp = '" & txt_codigo & "'"
        db.Execute "update ao_ventas_cobranza_prog set beneficiario_codigo_resp = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_resp = '" & txt_codigo & "'"
        db.Execute "update tc_zona_piloto_edif set beneficiario_codigo_rep = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_rep = '" & txt_codigo & "'"
        db.Execute "update tc_zona_piloto_edif set beneficiario_codigo_cobr = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_cobr = '" & txt_codigo & "'"
        db.Execute "update to_cronograma set beneficiario_codigo_resp = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_resp = '" & txt_codigo & "'"
        db.Execute "update to_cronograma_mensual set beneficiario_codigo_resp = '" & txt_nuevo_num.Text & "' where beneficiario_codigo_resp = '" & txt_codigo & "'"
        db.Execute "update gc_usuarios set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
        db.Execute "update rc_unidad_vs_responsable set beneficiario_codigo = '" & txt_nuevo_num.Text & "' where beneficiario_codigo = '" & txt_codigo & "'"
     Else
        txt_nuevo_num.Text = txt_codigo
     End If
     'rw_ficha_rrhh.Ado_datos.Recordset.Update
    Call rw_ficha_rrhh.encontrar
    Unload Me
Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub dtc_cod_Click(Area As Integer)
    dtc_desc.BoundText = dtc_cod.BoundText
End Sub

Private Sub dtc_desc_Change()
    dtc_cod.BoundText = dtc_desc.BoundText
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLA
    cambiar = "NO"
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLA()
    Set rs_empresa = New ADODB.Recordset
    If rs_empresa.State = 1 Then rs_empresa.Close
    rs_empresa.Open "select * from gc_empresas where estado_codigo ='APR' ", db, adOpenKeyset, adLockOptimistic
    Set Ado_empresa.Recordset = rs_empresa
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Option1_Click()
    cambiar = "SI"
    fra_doc.Enabled = True
    txt_nuevo_num.SetFocus
End Sub

Private Sub Option2_Click()
    cambiar = "NO"
    fra_doc.Enabled = False
End Sub

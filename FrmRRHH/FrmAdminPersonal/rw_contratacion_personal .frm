VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rw_contratacion_personal 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RRHH - Procesos de RRHH - Contratación de Personal"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   435
   ClientWidth     =   18300
   Icon            =   "rw_contratacion_personal .frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   18300
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame FraNavega 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LISTADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   7815
      Left            =   120
      TabIndex        =   95
      Top             =   720
      Width           =   6975
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   3720
         TabIndex        =   98
         Top             =   7530
         Width           =   975
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pendientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   1440
         TabIndex        =   97
         Top             =   7530
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   7095
         Left            =   120
         TabIndex        =   96
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   12515
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "beneficiario_codigo"
            Caption         =   "C.I."
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
            DataField       =   "beneficiario_denominacion"
            Caption         =   "Nombre Completo"
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
         BeginProperty Column02 
            DataField       =   "beneficiario_fecha_inicio"
            Caption         =   "Fecha.Inicio"
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
         BeginProperty Column03 
            DataField       =   "beneficiario_haber_mensual_bs"
            Caption         =   "Haber.Mensual.Bs."
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
         BeginProperty Column04 
            DataField       =   "beneficiario_fecha_fin"
            Caption         =   "Fecha.Fin"
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
         BeginProperty Column05 
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3075.024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   659.906
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_contratacion 
         Height          =   330
         Left            =   120
         Top             =   7440
         Width           =   6735
         _ExtentX        =   11880
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
         BackColor       =   16777215
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
         Caption         =   ""
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
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20400
      TabIndex        =   81
      Top             =   0
      Width           =   20400
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   16800
         Picture         =   "rw_contratacion_personal .frx":0A02
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   89
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5760
         Picture         =   "rw_contratacion_personal .frx":11C4
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   88
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4320
         Picture         =   "rw_contratacion_personal .frx":1979
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   87
         ToolTipText     =   "Aprueba Cronograma"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3000
         Picture         =   "rw_contratacion_personal .frx":21AC
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   86
         ToolTipText     =   "Anular Cronograma"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1440
         Picture         =   "rw_contratacion_personal .frx":28F8
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   85
         ToolTipText     =   "Modifica Datos Cabecera"
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "rw_contratacion_personal .frx":320D
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   84
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   9840
         Picture         =   "rw_contratacion_personal .frx":39CC
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7080
         Picture         =   "rw_contratacion_personal .frx":3BD6
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   82
         ToolTipText     =   "Imprimir el Listado de los Registros"
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO"
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
         Left            =   12360
         TabIndex        =   90
         Top             =   195
         Width           =   885
      End
   End
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
      ScaleWidth      =   20520
      TabIndex        =   91
      Top             =   0
      Visible         =   0   'False
      Width           =   20520
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2880
         Picture         =   "rw_contratacion_personal .frx":44A3
         ScaleHeight     =   615
         ScaleWidth      =   1305
         TabIndex        =   93
         Top             =   0
         Width           =   1300
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4275
         Picture         =   "rw_contratacion_personal .frx":4C79
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   92
         Top             =   0
         Width           =   1400
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
         Left            =   14175
         TabIndex        =   94
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00C00000&
      Height          =   7815
      Left            =   7080
      TabIndex        =   25
      Top             =   720
      Width           =   11055
      Begin VB.OptionButton Option2 
         BackColor       =   &H000040C0&
         Caption         =   "Postulante Existente en la Base de Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   5280
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000040C0&
         Caption         =   "Postulante Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   960
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   2775
         Left            =   240
         TabIndex        =   43
         Top             =   1200
         Visible         =   0   'False
         Width           =   10605
         Begin VB.TextBox TxtEmail 
            DataField       =   "beneficiario_email"
            DataSource      =   "Ado_contratacion"
            Height          =   285
            Left            =   6360
            MaxLength       =   50
            TabIndex        =   10
            Top             =   2280
            Width           =   3855
         End
         Begin VB.TextBox txtDireccion 
            DataField       =   "beneficiario_domicilio_legal"
            DataSource      =   "Ado_contratacion"
            Height          =   285
            Left            =   360
            MaxLength       =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   1680
            Width           =   9855
         End
         Begin VB.TextBox txtMat 
            DataField       =   "beneficiario_segundo_apellido"
            DataSource      =   "Ado_contratacion"
            Height          =   285
            Left            =   360
            MaxLength       =   30
            TabIndex        =   5
            Top             =   1100
            Width           =   4575
         End
         Begin VB.TextBox txtTelefono 
            DataField       =   "beneficiario_telefono_Cel"
            DataSource      =   "Ado_contratacion"
            Height          =   285
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   9
            Top             =   2280
            Width           =   3615
         End
         Begin VB.TextBox txtNom 
            DataField       =   "beneficiario_nombres"
            DataSource      =   "Ado_contratacion"
            Height          =   285
            Left            =   5280
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1100
            Width           =   4935
         End
         Begin VB.TextBox txtCI 
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   285
            Left            =   360
            MaxLength       =   20
            TabIndex        =   2
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox txtPat 
            DataField       =   "beneficiario_primer_apellido"
            DataSource      =   "Ado_contratacion"
            Height          =   285
            Left            =   5280
            MaxLength       =   35
            TabIndex        =   4
            Top             =   480
            Width           =   4935
         End
         Begin MSDataListLib.DataCombo dtc_depto_codigo 
            Bindings        =   "rw_contratacion_personal .frx":5565
            DataField       =   "depto_sigla"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   3720
            TabIndex        =   3
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ForeColor       =   0
            ListField       =   "depto_sigla"
            BoundColumn     =   "depto_sigla"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_genero 
            Bindings        =   "rw_contratacion_personal .frx":557D
            DataField       =   "genero_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   360
            TabIndex        =   8
            Top             =   2280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ForeColor       =   0
            ListField       =   "genero_codigo"
            BoundColumn     =   "genero_codigo"
            Text            =   ""
         End
         Begin VB.Label lblbien 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Correo Electrónico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   6360
            TabIndex        =   112
            Top             =   2040
            Width           =   2370
         End
         Begin VB.Label lblbien 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Genero"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   106
            Top             =   2040
            Width           =   1050
         End
         Begin VB.Label lblbien 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Documento de Identidad "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   4
            Left            =   360
            TabIndex        =   104
            Top             =   240
            Width           =   2715
         End
         Begin VB.Label lbl_campo6 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Expedido en"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3720
            TabIndex        =   103
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label lblbien 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección Postulante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   48
            Top             =   1410
            Width           =   2115
         End
         Begin VB.Label lblbien 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfonos del Postulante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   2400
            TabIndex        =   47
            Top             =   2040
            Width           =   2370
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombres"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   5280
            TabIndex        =   46
            Top             =   855
            Width           =   1440
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Segundo Apellido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   2
            Left            =   360
            TabIndex        =   45
            Top             =   855
            Width           =   1620
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Primer Apellido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   5280
            TabIndex        =   44
            Top             =   240
            Width           =   1980
         End
      End
      Begin VB.Frame Frame5 
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
         ForeColor       =   &H00FFFF80&
         Height          =   3855
         Left            =   240
         TabIndex        =   63
         Top             =   3840
         Visible         =   0   'False
         Width           =   10605
         Begin VB.ComboBox cmb_meses 
            Height          =   315
            ItemData        =   "rw_contratacion_personal .frx":5595
            Left            =   4440
            List            =   "rw_contratacion_personal .frx":55E1
            TabIndex        =   20
            Text            =   "01"
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox txt_monto1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "beneficiario_monto_adjudica_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_contratacion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   9000
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   80
            Top             =   3360
            Width           =   1350
         End
         Begin VB.TextBox txt_tiempo 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "beneficiario_tiempo_meses"
            DataSource      =   "Ado_contratacion"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9000
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   24
            Top             =   2880
            Width           =   1350
         End
         Begin VB.TextBox txt_monto2 
            DataField       =   "beneficiario_haber_mensual_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_contratacion"
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
            Height          =   285
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   21
            Top             =   3360
            Width           =   1575
         End
         Begin VB.TextBox txt_monto3 
            DataField       =   "beneficiario_otro_mensual_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_contratacion"
            Height          =   285
            Left            =   5280
            MaxLength       =   20
            TabIndex        =   22
            Top             =   3360
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "rw_contratacion_personal .frx":5645
            DataField       =   "ocup_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   4200
            TabIndex        =   67
            Top             =   240
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   16777215
            ListField       =   "ocup_codigo"
            BoundColumn     =   "ocup_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "rw_contratacion_personal .frx":565F
            DataField       =   "ocup_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   360
            TabIndex        =   11
            Top             =   480
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "ocup_descripcion"
            BoundColumn     =   "ocup_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "rw_contratacion_personal .frx":5679
            DataField       =   "nivel_educ_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   5520
            TabIndex        =   12
            Top             =   480
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "nivel_educ_descripcion"
            BoundColumn     =   "nivel_educ_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "rw_contratacion_personal .frx":5693
            DataField       =   "nivel_educ_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   9600
            TabIndex        =   68
            Top             =   240
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   16777215
            ListField       =   "nivel_educ_codigo"
            BoundColumn     =   "nivel_educ_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "rw_contratacion_personal .frx":56AD
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   4080
            TabIndex        =   69
            Top             =   840
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "munic_codigo"
            BoundColumn     =   "munic_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "rw_contratacion_personal .frx":56C7
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   360
            TabIndex        =   13
            Top             =   1170
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "munic_descripcion"
            BoundColumn     =   "munic_codigo"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker txtFecha 
            DataField       =   "cotiza_fecha"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   7080
            TabIndex        =   70
            Top             =   3120
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   4210752
            CheckBox        =   -1  'True
            Format          =   111476737
            CurrentDate     =   44562
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker txtFecha2 
            DataField       =   "cotiza_fecha_limite_postulacion"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   3720
            TabIndex        =   71
            Top             =   3120
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   111476737
            CurrentDate     =   44562
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker txtFecha3 
            DataField       =   "cotiza_fecha_programada_contrato"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   360
            TabIndex        =   72
            Top             =   3120
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   111476737
            CurrentDate     =   44562
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   7080
            TabIndex        =   73
            Top             =   2640
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   111476737
            CurrentDate     =   44562
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1800
            TabIndex        =   19
            Top             =   2880
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   111476737
            CurrentDate     =   44562
            MaxDate         =   55153
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo dtc_desc_3 
            Bindings        =   "rw_contratacion_personal .frx":56E1
            DataField       =   "modalidad_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   360
            TabIndex        =   15
            Top             =   1800
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "modalidad_descripcion"
            BoundColumn     =   "modalidad_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo detc_cod_3 
            Bindings        =   "rw_contratacion_personal .frx":56FA
            DataField       =   "modalidad_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   3120
            TabIndex        =   76
            Top             =   1560
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "modalidad_codigo"
            BoundColumn     =   "modalidad_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc_4 
            Bindings        =   "rw_contratacion_personal .frx":5712
            DataField       =   "solicitud_tipo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   4200
            TabIndex        =   16
            Top             =   1800
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "solicitud_tipo_descripcion"
            BoundColumn     =   "solicitud_tipo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_cod_4 
            Bindings        =   "rw_contratacion_personal .frx":572A
            DataField       =   "solicitud_tipo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   9480
            TabIndex        =   78
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "solicitud_tipo"
            BoundColumn     =   "solicitud_tipo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc_5 
            Bindings        =   "rw_contratacion_personal .frx":5742
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   360
            TabIndex        =   17
            Top             =   2400
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "unidad_descripcion"
            BoundColumn     =   "unidad_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_cod_5 
            Bindings        =   "rw_contratacion_personal .frx":575A
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   3960
            TabIndex        =   100
            Top             =   2160
            Visible         =   0   'False
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "unidad_codigo"
            BoundColumn     =   "unidad_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_des_6 
            Bindings        =   "rw_contratacion_personal .frx":5772
            DataField       =   "puesto_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   5520
            TabIndex        =   18
            Top             =   2400
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "puesto_descripcion"
            BoundColumn     =   "puesto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_cod_6 
            Bindings        =   "rw_contratacion_personal .frx":578A
            DataField       =   "puesto_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   9480
            TabIndex        =   102
            Top             =   2160
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "puesto_codigo"
            BoundColumn     =   "puesto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_dpto_cod 
            Bindings        =   "rw_contratacion_personal .frx":57A2
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   3360
            TabIndex        =   105
            Top             =   840
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "depto_codigo"
            BoundColumn     =   "munic_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_cargo 
            Bindings        =   "rw_contratacion_personal .frx":57BC
            DataField       =   "puesto_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   8280
            TabIndex        =   108
            Top             =   2160
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "cargo_codigo"
            BoundColumn     =   "puesto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_empresa_den 
            Bindings        =   "rw_contratacion_personal .frx":57D4
            DataField       =   "codigo_empresa"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   5520
            TabIndex        =   14
            Top             =   1170
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "denominacion_empresa"
            BoundColumn     =   "codigo_empresa"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_empresa_cod 
            Bindings        =   "rw_contratacion_personal .frx":57EC
            DataField       =   "codigo_empresa"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   9600
            TabIndex        =   109
            Top             =   840
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   16777215
            ListField       =   "codigo_empresa"
            BoundColumn     =   "codigo_empresa"
            Text            =   ""
         End
         Begin VB.Label LblTiempo 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tiempo (Meses)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7080
            TabIndex        =   111
            Top             =   2880
            Width           =   1560
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5520
            TabIndex        =   110
            Top             =   930
            Width           =   1230
         End
         Begin VB.Label LblMeses 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Meses                    (0 es indefinido)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3720
            TabIndex        =   107
            Top             =   2880
            Width           =   3000
         End
         Begin VB.Label LblPuesto 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Puesto a Ocupar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5520
            TabIndex        =   101
            Top             =   2160
            Width           =   1500
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento (Oficina)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   360
            TabIndex        =   99
            Top             =   2160
            Width           =   2070
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Contrato"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   4200
            TabIndex        =   79
            Top             =   1560
            Width           =   1500
         End
         Begin VB.Label lbl_mod 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Modalidad de Contratación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   360
            TabIndex        =   77
            Top             =   1560
            Width           =   2430
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Inicio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   7
            Left            =   360
            TabIndex        =   75
            Top             =   2880
            Width           =   1365
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   $"rw_contratacion_personal .frx":5804
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   2
            Left            =   360
            TabIndex        =   74
            Top             =   3360
            Width           =   8265
         End
         Begin VB.Label lbl_campo2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Perfil Profesional"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   360
            TabIndex        =   66
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label lbl_campo3 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel Educacional (Mayor Importancia)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5520
            TabIndex        =   65
            Top             =   240
            Width           =   3825
         End
         Begin VB.Label lbl_campo4 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar de Postulación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   64
            Top             =   930
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Elije la Persona"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2415
         Left            =   480
         TabIndex        =   34
         Top             =   1200
         Visible         =   0   'False
         Width           =   10005
         Begin VB.CommandButton BtnNo 
            Height          =   555
            Left            =   5160
            MaskColor       =   &H00000000&
            Picture         =   "rw_contratacion_personal .frx":5891
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Cancelar"
            Top             =   1320
            Width           =   1365
         End
         Begin VB.CommandButton BtnOk 
            Height          =   555
            Left            =   3480
            Picture         =   "rw_contratacion_personal .frx":626B
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   1320
            Width           =   1365
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "rw_contratacion_personal .frx":6B6F
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   720
            TabIndex        =   35
            Top             =   840
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "rw_contratacion_personal .frx":6B89
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   720
            TabIndex        =   36
            Top             =   1200
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux1 
            Bindings        =   "rw_contratacion_personal .frx":6BA3
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   4200
            TabIndex        =   37
            Top             =   1200
            Visible         =   0   'False
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "beneficiario_primer_apellido"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux3 
            Bindings        =   "rw_contratacion_personal .frx":6BBD
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   720
            TabIndex        =   38
            Top             =   1560
            Visible         =   0   'False
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "beneficiario_nombres"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux4 
            Bindings        =   "rw_contratacion_personal .frx":6BD7
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   3720
            TabIndex        =   39
            Top             =   1560
            Visible         =   0   'False
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "beneficiario_telefono_Cel"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux2 
            Bindings        =   "rw_contratacion_personal .frx":6BF1
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   6960
            TabIndex        =   40
            Top             =   360
            Visible         =   0   'False
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "beneficiario_segundo_apellido"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux5 
            Bindings        =   "rw_contratacion_personal .frx":6C0B
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_contratacion"
            Height          =   315
            Left            =   6480
            TabIndex        =   41
            Top             =   1560
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "beneficiario_domicilio_legal"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Apellidos y Nombres"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   6
            Left            =   720
            TabIndex        =   42
            Top             =   480
            Width           =   1890
         End
      End
      Begin VB.TextBox txtEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5235
         MaxLength       =   80
         TabIndex        =   58
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00404040&
         Caption         =   "Elije el Medio de Comunicación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   2415
         Left            =   480
         TabIndex        =   52
         Top             =   1200
         Visible         =   0   'False
         Width           =   8685
         Begin VB.CommandButton BtnOk2 
            BackColor       =   &H00C0C000&
            Caption         =   "Aceptar"
            Height          =   675
            Left            =   3000
            Picture         =   "rw_contratacion_personal .frx":6C25
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   1200
            Width           =   765
         End
         Begin VB.CommandButton BtnNo2 
            BackColor       =   &H00C0C000&
            Caption         =   "Cancelar"
            Height          =   675
            Left            =   4320
            MaskColor       =   &H00000000&
            Picture         =   "rw_contratacion_personal .frx":7627
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Cancelar"
            Top             =   1200
            Width           =   765
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "rw_contratacion_personal .frx":7BB1
            DataField       =   "nivel_educ_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   2520
            TabIndex        =   55
            Top             =   600
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "nivel_educ_descripcion"
            BoundColumn     =   "nivel_educ_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "rw_contratacion_personal .frx":7BCB
            DataField       =   "nivel_educ_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   2880
            TabIndex        =   57
            Top             =   360
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "nivel_educ_codigo"
            BoundColumn     =   "nivel_educ_codigo"
            Text            =   ""
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Nombre del Medio"
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
            Left            =   720
            TabIndex        =   56
            Top             =   600
            Width           =   1680
         End
      End
      Begin VB.TextBox txt_campo1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         DataField       =   "unidad_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle2"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         MaxLength       =   80
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtSW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5880
         MaxLength       =   80
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "rw_contratacion_personal .frx":7BE5
         DataField       =   "puesto_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle2"
         Height          =   315
         Left            =   7680
         TabIndex        =   61
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "puesto_codigo"
         BoundColumn     =   "puesto_codigo"
         Text            =   "-"
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H000040C0&
         Caption         =   "Elije puesto convocado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   480
         TabIndex        =   62
         Top             =   1320
         Width           =   4215
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "rw_contratacion_personal .frx":7BFF
         DataField       =   "puesto_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle2"
         Height          =   315
         Left            =   2280
         TabIndex        =   23
         Top             =   1320
         Visible         =   0   'False
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "puesto_descripcion"
         BoundColumn     =   "puesto_codigo"
         Text            =   "-"
      End
      Begin VB.Label lbl_convoca 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "cotiza_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle2"
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
         Left            =   9600
         TabIndex        =   60
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. RRHH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8040
         TabIndex        =   59
         Top             =   225
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Puesto al que postula"
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
         Left            =   480
         TabIndex        =   51
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Convocatoria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   9555
         TabIndex        =   32
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label txtBenef 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "rrhh_codigo"
         DataSource      =   "Ado_contratacion"
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
         Left            =   8040
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   11160
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Trámite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   220
         Width           =   1065
      End
      Begin VB.Label lbl_campo_des 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Ejecutora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   2025
         TabIndex        =   29
         Top             =   225
         Width           =   1560
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "Ado_contratacion"
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
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label Txt_descripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DEPARTAMENTO DE RECURSOS HUMANOS"
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
         TabIndex        =   27
         Top             =   480
         Width           =   5535
      End
   End
   Begin MSAdodcLib.Adodc Ado_clasif1 
      Height          =   330
      Left            =   360
      Top             =   7680
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_clasif1"
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
   Begin MSAdodcLib.Adodc Ado_clasif2 
      Height          =   330
      Left            =   2520
      Top             =   7680
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_clasif2"
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
   Begin MSAdodcLib.Adodc Ado_clasif3 
      Height          =   330
      Left            =   4680
      Top             =   7680
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_clasif3"
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
   Begin MSAdodcLib.Adodc Ado_clasif4 
      Height          =   330
      Left            =   360
      Top             =   8040
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_clasif4"
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
   Begin MSAdodcLib.Adodc Ado_clasif5 
      Height          =   330
      Left            =   2520
      Top             =   8040
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_clasif5"
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
   Begin MSAdodcLib.Adodc Ado_clasif6 
      Height          =   330
      Left            =   4680
      Top             =   8040
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_clasif6"
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
   Begin MSAdodcLib.Adodc Ado_datos 
      Height          =   330
      Left            =   6840
      Top             =   7680
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_clasif3"
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
   Begin MSAdodcLib.Adodc Ado_aux_1 
      Height          =   330
      Left            =   360
      Top             =   8400
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_aux_1"
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
   Begin MSAdodcLib.Adodc Ado_aux_2 
      Height          =   330
      Left            =   2520
      Top             =   8400
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_aux_2"
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
   Begin MSAdodcLib.Adodc Ado_aux_3 
      Height          =   330
      Left            =   4680
      Top             =   8400
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_aux_3"
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
   Begin MSAdodcLib.Adodc Ado_aux_4 
      Height          =   330
      Left            =   6840
      Top             =   8400
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_aux_4"
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
   Begin MSAdodcLib.Adodc Ado_aux_5 
      Height          =   330
      Left            =   9000
      Top             =   8400
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_aux_5"
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
   Begin MSAdodcLib.Adodc Ado_aux_6 
      Height          =   330
      Left            =   11160
      Top             =   8400
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_aux_6"
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
   Begin MSAdodcLib.Adodc Ado_aux_7 
      Height          =   330
      Left            =   13320
      Top             =   8400
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_aux_7"
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
   Begin MSAdodcLib.Adodc Ado_aux_8 
      Height          =   330
      Left            =   360
      Top             =   8760
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Ado_aux_8"
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
Attribute VB_Name = "rw_contratacion_personal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Para_Aceptado As String
Dim rs_datos As New ADODB.Recordset
Dim rs_contratacion As New ADODB.Recordset
Dim RS_BENEF As New ADODB.Recordset
Dim rs_clasif1 As New ADODB.Recordset
Dim rs_clasif2 As New ADODB.Recordset
Dim rs_clasif3 As New ADODB.Recordset
Dim rs_clasif4 As New ADODB.Recordset
Dim rs_clasif5 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux_1 As New ADODB.Recordset
Dim rs_aux_2 As New ADODB.Recordset
Dim rs_aux_3 As New ADODB.Recordset
Dim rs_aux_4 As New ADODB.Recordset
Dim rs_aux_5 As New ADODB.Recordset
Dim rs_aux_6 As New ADODB.Recordset
Dim rs_aux_7 As New ADODB.Recordset
Dim rs_aux_8 As New ADODB.Recordset
Dim rs_puestos As New ADODB.Recordset
Dim rs_puesto_nuevo As New ADODB.Recordset
Dim rs_UNIDAD As New ADODB.Recordset

Dim rs_pla As New ADODB.Recordset
Dim rs_sub_pla As New ADODB.Recordset
Dim rs_guardar As New ADODB.Recordset

Dim gestion, mes, dia As Integer
Dim VAR_TIME As Integer
Dim VAR_ADJUDICA As Integer

Dim nomb2, VAR_NOMBRE As String
Dim puesto2, iniciales As String
Dim modif, Base, Nuevo, subplanilla As String
Dim VAR_CITE As String

Dim FECHA_FN As Date

Private Sub Ado_contratacion_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If Nuevo = "NO" Then
If Ado_contratacion.Recordset.RecordCount > 0 Then


DTPicker1.Value = Ado_contratacion.Recordset!beneficiario_fecha_inicio
DTPicker2.Value = Ado_contratacion.Recordset!beneficiario_fecha_fin
If Ado_contratacion.Recordset!beneficiario_tiempo_meses < 10 Then
cmb_meses.Text = "0" & Ado_contratacion.Recordset!beneficiario_tiempo_meses
Else
'cmb_meses.Text = Ado_contratacion.Recordset!beneficiario_tiempo_meses
End If
End If
End If
End Sub

Private Sub nuevo_2()
  txtCI2.Text = ""
    txtPat2.Text = ""
    txtMat2.Text = ""
    txtNom2.Text = ""
    TxtTelefono2.Text = ""
    txtDireccion2.Text = ""
    
    txtCI2.Visible = True
    txtPat2.Visible = True
    txtMat2.Visible = True
    txtNom2.Visible = True
    TxtTelefono2.Visible = True
    txtDireccion2.Visible = True
    
    txtCI.Visible = False
    txtPat.Visible = False
    txtMat.Visible = False
    txtNom.Visible = False
    txtTelefono.Visible = False
    txtDireccion.Visible = False
End Sub

Private Sub BtnAñadir_Click()
modif = "NO"
Nuevo = "SI"
Ado_contratacion.Recordset.AddNew
DTPicker1.Value = Date
DTPicker2.Value = Date
Option1.Value = True
 Fra_datos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    FraNavega.Enabled = False
    dg_datos.Enabled = False
     Frame4.Enabled = True
    txt_tiempo.Text = "0"
    cmb_meses.Text = "0"
    dtc_genero.Text = ""
    txtCI.Text = ""
    txtPat.Text = ""
    txtMat.Text = ""
    txtNom.Text = ""
    txtTelefono.Text = ""
    txtDireccion.Text = ""
    dtc_depto_codigo.Text = ""
    dtc_desc2.Text = ""
    dtc_desc3.Text = ""
    dtc_desc4.Text = ""
    dtc_desc_3.Text = ""
    dtc_desc_4.Text = ""
    dtc_desc_5.Text = ""
    dtc_des_6.Text = ""
    txt_monto2.Text = "0"
    txt_monto3.Text = "0"
    txt_monto1.Text = "0"
    
    dtc_codigo2.Text = ""
    dtc_codigo3.Text = ""
    dtc_codigo4.Text = ""

    detc_cod_3.Text = ""
    dtc_cod_4.Text = ""
     dtc_cod_5.Text = ""
        dtc_cod_6.Text = ""
    
   
    'DTPicker1.Value = Date
    'DTPicker2.Value = Date
    Call ABRIR_TABLA_AUX
End Sub

Private Sub BtnAprobar_Click()
On Error GoTo UpdateErr
   If rs_contratacion!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         Set rs_pla = New ADODB.Recordset
         If rs_pla.State = 1 Then rs_pla.Close
         rs_pla.Open "Select * from ro_personal_contratado WHERE beneficiario_codigo = '" & Ado_contratacion.Recordset!beneficiario_codigo & "' AND codigo_empresa = " & Ado_contratacion.Recordset!codigo_empresa & "", db, adOpenKeyset, adLockOptimistic, adCmdText
         If rs_pla.RecordCount > 0 Then
            sino = MsgBox("Esta persona ya fue contratada desea ACTUALIZAR el contrato con los nuevos DATOS?", vbYesNo + vbQuestion, "Atención")
            If sino = vbYes Then
                db.Execute "delete ro_personal_contratado where beneficiario_codigo ='" & Ado_contratacion.Recordset!beneficiario_codigo & "'"
            Else
                  MsgBox "No se aprobó el registro!"
                  Exit Sub
            End If
         End If
         If dtc_cod_4.Text = "24" Then
            Select Case dtc_dpto_cod.Text
               Case 1, 6, 5
                  subplanilla = "P040"
               Case 2, 4
                  subplanilla = "P010"
               Case 3
                  subplanilla = "P030"
               Case 7, 8, 9
                  subplanilla = "P020"
            End Select
'             Set rs_pla = New ADODB.Recordset
'             If rs_pla.State = 1 Then rs_pla.Close
'             rs_pla.Open "Select * from rc_planilla_grupo WHERE depto_codigo = '" & dtc_dpto_cod.Text & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
'            If rs_pla.RecordCount > 0 Then
'             subplanilla = rs_pla!planilla_codigo & "0"
'             Set rs_sub_pla = New ADODB.Recordset
'             If rs_sub_pla.State = 1 Then rs_sub_pla.Close
'             rs_sub_pla.Open "Select * from rc_planilla_sub_grupo WHERE unidad_codigo_pla = '" & subplanilla & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
'               If rs_sub_pla.RecordCount = 0 Then
'              'sino = MsgBox("No existe una sub plainilla para personal a prueba en " & dtc_desc4.Text, vbInformation, "Aviso")
'
'               End If
'            End If
         Else
            subplanilla = ""
         End If
         db.Execute "Insert INTO ro_personal_contratado (rrhh_codigo,adjudica_codigo,beneficiario_codigo,puesto_codigo,unidad_codigo,cargo_codigo,Fecha_ingreso,fecha_expiracion,ocup_codigo,beneficiario_haber_mensual,beneficiario_otro_mensual,estado_codigo,usr_codigo,fecha_registro,solicitud_tipo,unidad_codigo_pla, bono_antiguedad, genero_codigo, codigo_empresa)" & _
         " Values (" & Ado_contratacion.Recordset!adjudica_codigo & "," & Ado_contratacion.Recordset!adjudica_codigo & ",'" & Ado_contratacion.Recordset!beneficiario_codigo & "','" & Ado_contratacion.Recordset!puesto_codigo & "','" & Ado_contratacion.Recordset!unidad_codigo & "','" & dtc_cargo.Text & "','" & Ado_contratacion.Recordset!beneficiario_fecha_inicio & "','" & Ado_contratacion.Recordset!beneficiario_fecha_fin & "'," & Ado_contratacion.Recordset!ocup_codigo & ",'" & Ado_contratacion.Recordset!beneficiario_haber_mensual_bs & "'," & Ado_contratacion.Recordset!beneficiario_otro_mensual_bs & ",'REG', '" & glusuario & "',  '" & Date & "'," & Ado_contratacion.Recordset!solicitud_tipo & ",'" & subplanilla & "', '0', '" & Ado_contratacion.Recordset!genero_codigo & "'," & Ado_contratacion.Recordset!codigo_empresa & ")"
                'db.Execute "Insert INTO ro_personal_contratado (rrhh_codigo, beneficiario_codigo, puesto_codigo, unidad_codigo, cargo_codigo, fecha_ingreso, fecha_expiracion, ocup_codigo, beneficiario_haber_mensual, estado_codigo, usr_codigo, fecha_registro, beneficiario_otro_mensual) Values (" & frm_ao_contratacion.Ado_contratacion.Recordset!rrhh_codigo & ", '" & dtc_codigo5.Text & "', '" & GlPuesto & "', '" & Ado_contratacion.Recordset!unidad_codigo & "',  '" & Ado_contratacion.Recordset!cargo_codigo & "',  '" & Ado_contratacion.Recordset!beneficiario_fecha_inicio & "', '" & Ado_contratacion.Recordset!beneficiario_fecha_fin & "', '" & VAR_OCUP & "', " & frm_ao_contratacion.Ado_detalle3.Recordset!beneficiario_haber_mensual_bs & ", 'REG', '" & glusuario & "',  '" & Date & "',  " & txt_monto3.Text & ")"
         rs_contratacion!estado_codigo = "APR"
         VAR_AUX = Left(Ado_contratacion.Recordset("beneficiario_nombres"), 1) + Ado_contratacion.Recordset("beneficiario_primer_apellido")
         VAR_PWD = Encriptar(Trim(Ado_contratacion.Recordset("beneficiario_codigo")))
'            db.Execute "insert into gc_usuarios(usr_codigo, beneficiario_codigo, usr_nombres, usr_primer_apellido, usr_segundo_apellido, usr_clave, IdNivelAcceso, estado_codigo, fecha_registro, dgral_codigo, da_codigo, unidad_codigo, ocup_codigo, usr_observaciones)" & _
'            "values ('" & Left(Ado_datos.Recordset("beneficiario_nombres"), 1) & "' + '" & Ado_datos.Recordset("beneficiario_primer_apellido") & "', '" & Ado_datos.Recordset("beneficiario_codigo") & "','" & Trim(Ado_datos.Recordset("beneficiario_nombres")) & "', '" & Ado_datos.Recordset("beneficiario_primer_apellido") & "','" & Trim(Ado_datos.Recordset("beneficiario_segundo_apellido")) & "','" & Ado_datos.Recordset("beneficiario_codigo") & "', '1', 'REG', '" & Date & "', '0', '0', '0', '0', '0') "
            
         Set rs_aux_8 = New ADODB.Recordset
         If rs_aux_8.State = 1 Then rs_aux_8.Close
         rs_aux_8.Open "SELECT * FROM gc_usuarios where beneficiario_codigo = '" & Ado_contratacion.Recordset("beneficiario_codigo") & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
         If rs_aux_8.RecordCount = 0 Then
            db.Execute "insert into gc_usuarios(usr_codigo, beneficiario_codigo, usr_nombres, usr_primer_apellido, usr_segundo_apellido, usr_clave, dgral_codigo, da_codigo, unidad_codigo, ocup_codigo, usr_observaciones, idnivelacceso, estado_codigo, fecha_registro)" & _
            "values ('" & VAR_AUX & "', '" & Ado_contratacion.Recordset("beneficiario_codigo") & "','" & Trim(Ado_contratacion.Recordset("beneficiario_nombres")) & "', '" & Ado_contratacion.Recordset("beneficiario_primer_apellido") & "','" & Trim(Ado_contratacion.Recordset("beneficiario_segundo_apellido")) & "','" & VAR_PWD & "', '1', '0', '0', '0', '-', '1', 'REG', '" & Date & "') "
            iniciales = Left(Trim(Ado_contratacion.Recordset("beneficiario_primer_apellido")), 1) & Left(Trim(Ado_contratacion.Recordset("beneficiario_segundo_apellido")), 1) & Left(Trim(Ado_contratacion.Recordset("beneficiario_nombres")), 1)
            RUTA1 = "PERSONAL" + "\" + iniciales + "-" + Trim(Ado_contratacion.Recordset("beneficiario_codigo"))
           'MsgBox RUTA1
            MkDir RUTA1
            MkDir RUTA1 + "\CONTRATOS"
            MkDir RUTA1 + "\FINIQUITO"
            MkDir RUTA1 + "\MEMOS"
            MkDir RUTA1 + "\RESPALDOS"
            MkDir RUTA1 + "\HOJA_VIDA"
            MkDir RUTA1 + "\OTROS"
            MkDir RUTA1 + "\EVALUACIONES"
            MkDir RUTA1 + "\LICENCIAS"
            MkDir RUTA1 + "\VACACIONES"
         End If
          
         db.Execute "UPDATE gc_beneficiario set depto_codigo = " & dtc_dpto_cod.Text & ", beneficiario_iniciales = '" & iniciales & "', estado_codigo = 'APR', tipoben_codigo = 1 where beneficiario_codigo = '" & Ado_contratacion.Recordset!beneficiario_codigo & "'"
    
         rs_contratacion!fecha_registro = Date
         rs_contratacion!usr_codigo = glusuario
         rs_contratacion.UpdateBatch adAffectAll
         
         db.Execute "UPDATE ro_rrhh_adjudica_personas set rrhh_codigo = " & Ado_contratacion.Recordset!adjudica_codigo & ", solicitud_codigo = " & Ado_contratacion.Recordset!adjudica_codigo & " WHERE adjudica_codigo = " & Ado_contratacion.Recordset!adjudica_codigo
        

      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado (ERR) o Aprobado (APR) anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnBuscar_Click()
   Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexión = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = dg_datos
    ClBuscaGrid.QueryUtilizado = modif
    Set ClBuscaGrid.RecordsetTrabajo = rs_contratacion
    'ClBuscaGrid.CamposVisibles = "11010011"
    ClBuscaGrid.Ejecutar
End Sub

Private Sub BtnCancelar_Click()
'cancela la edicion de datos
 rs_contratacion.Cancel
 Ado_contratacion.Recordset.Cancel
  Call OptFilGral1_Click
    Para_Aceptado = "N"
    Nuevo = "NO"
'    txtSW = "0"
  Fra_datos.Enabled = False
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FraNavega.Enabled = True
    dg_datos.Enabled = True
    Option1.Value = True
    Option2.Value = False

   OptFilGral2.Value = True
End Sub

Private Sub BtnEliminar_Click()
If rs_contratacion!estado_codigo = "APR" Then
 sino = MsgBox("Esta Seguro de anular este registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_contratacion!estado_codigo = "ANL"
         rs_contratacion!fecha_registro = Date
         rs_contratacion!usr_codigo = glusuario
         rs_contratacion.UpdateBatch adAffectAll
      End If
Else
 sino = MsgBox("No sepuede ANULAR si esta solo REGISTRADO (REG) o ya ANULADO (ANL)", vbExclamation, "Atención")
End If
End Sub

Private Sub BtnGrabar_Click()
On Error GoTo UpdateErr

If Valida Then
    If modif = "NO" Then
        Set rs_guardar = New ADODB.Recordset
        If rs_guardar.State = 1 Then rs_guardar.Close
        rs_guardar.Open "Select * from ro_rrhh_adjudica_personas", db, adOpenKeyset, adLockOptimistic, adCmdText
        'rs_benef
        Set RS_BENEF = New ADODB.Recordset
        If RS_BENEF.State = 1 Then RS_BENEF.Close
        RS_BENEF.Open "Select * from gc_beneficiario", db, adOpenKeyset, adLockOptimistic, adCmdText
        If Base = "NO" Then
            If txtCI.Text = "" Then
                MsgBox "Debe registrar ... " + lblbien(4).Caption, vbCritical + vbExclamation, "Validación de datos"
                Exit Sub
            End If
            If (dtc_depto_codigo.Text = "") Then
               MsgBox "Debe registrar ... " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
               Exit Sub
            End If
            Set RS_BENEF = New ADODB.Recordset
            If RS_BENEF.State = 1 Then RS_BENEF.Close
            RS_BENEF.Open "Select * from gc_beneficiario WHERE beneficiario_codigo = '" & txtCI.Text & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
            If RS_BENEF.RecordCount > 0 Then
                sino = MsgBox("Esta persona ya EXISTE en la BASE de DATOS", vbCritical, "Aviso")
                Frame2.Visible = True
                Frame4.Visible = False
                Frame3.Visible = False
                Base = "SI"
                dtc_codigo5.Text = ""
                dtc_codigo5.BoundText = Trim(txtCI.Text)
                dtc_desc5.BoundText = dtc_codigo5.BoundText
                dtc_aux1.BoundText = dtc_codigo5.BoundText
                dtc_aux2.BoundText = dtc_codigo5.BoundText
                dtc_aux3.BoundText = dtc_codigo5.BoundText
                dtc_aux4.BoundText = dtc_codigo5.BoundText
                dtc_aux5.BoundText = dtc_codigo5.BoundText
                Option2.Value = True
                '    txtSW = "2"
                Exit Sub
            End If
            RS_BENEF.AddNew
            RS_BENEF!beneficiario_codigo = Trim(txtCI.Text)
        Else
            Set RS_BENEF = New ADODB.Recordset
            If RS_BENEF.State = 1 Then RS_BENEF.Close
            RS_BENEF.Open "Select * from gc_beneficiario WHERE beneficiario_codigo = '" & txtCI.Text & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        End If
        RS_BENEF!depto_sigla = dtc_depto_codigo.Text
        RS_BENEF!beneficiario_primer_apellido = UCase(txtPat.Text)
        RS_BENEF!beneficiario_segundo_apellido = UCase(txtMat.Text)
        RS_BENEF!beneficiario_nombres = UCase(txtNom.Text)
        RS_BENEF!beneficiario_telefono_Cel = IIf(txtTelefono.Text = "", "0", txtTelefono.Text)
        RS_BENEF!beneficiario_domicilio_legal = txtDireccion.Text
        VAR_NOMBRE = IIf(txtPat.Text = "", "", UCase(txtPat.Text) + " ") & IIf(txtMat.Text = "", "", UCase(txtMat.Text) + " ") & IIf(txtNom.Text = "", "", UCase(txtNom.Text))
        RS_BENEF!beneficiario_denominacion = VAR_NOMBRE
        RS_BENEF!beneficiario_email = IIf(TxtEmail.Text = "", "0", TxtEmail.Text)
        RS_BENEF!estado_codigo = "REG"
        RS_BENEF!tipoben_codigo = "1"
        RS_BENEF!tipodoc_codigo = "C.I"
        RS_BENEF!usr_codigo = glusuario
        RS_BENEF!fecha_registro = Date
        RS_BENEF!depto_codigo = dtc_dpto_cod.Text
        RS_BENEF.Update
        rs_guardar.AddNew
        rs_guardar!estado_codigo = "REG"
        'db.Execute "Insert INTO ro_rrhh_adjudica_personas" & _
        "(ges_gestion, rrhh_codigo, adjudica_codigo, beneficiario_codigo, unidad_codigo, solicitud_codigo, solicitud_tipo, nivel_educ_codigo, modalidad_codigo, ocup_codigo, puesto_codigo, munic_codigo, beneficiario_monto_adjudica_bs, beneficiario_haber_mensual_bs, beneficiario_otro_mensual_bs,beneficiario_tiempo_meses, tipo_moneda, beneficiario_fecha_inicio,beneficiario_fecha_fin, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, cite_tramite, observaciones, estado_codigo, usr_codigo, fecha_registro) values ('" & Year(Date) & "', " & rs_aux2.RecordCount + 1 & ",'" & txtCI.Text & "','" & dtc_cod_5.Text & "'," & rs_aux2.RecordCount + 1 & ",'" & dtc_desc_4.Text & "'," & dtc_codigo3.Text & "'," & dtc_desc_3.Text & "'," & dtc_codigo2.Text & "'," & dtc_cod_6.Text & "'," & dtc_codigo4.Text & "'," & Val(txt_monto2.Text) & "," & Val(txt_monto3.Text) & ",'" & "BOB" & "'," & DTPicker1.Value & "," & DTPicker2.Value & " , '-', '-', '-', '-', '-', '-', '-','-','REG'" & _
        ", '" & glusuario & "','" & Date & "')"
        
        'db.Execute "Insert INTO  ro_rrhh_adjudica_personas (ges_gestion,rrhh_codigo, beneficiario_codigo, unidad_codigo, solicitud_codigo, puesto_codigo, beneficiario_haber_mensual_bs, beneficiario_otro_mensual_bs, tipo_moneda, beneficiario_fecha_inicio, beneficiario_fecha_fin,proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, cite_tramite, observaciones, estado_codigo, usr_codigo, fecha_registro) values ('" & Year(Date) & "', " & rs_aux2.RecordCount + 1 & ",'" & txtCI.Text & "','" & dtc_cod_5.Text & "'," & rs_aux2.RecordCount + 1 & ",'" & dtc_cod_6.Text & "'," & Val(Txt_monto2.Text) & "," & Val(Txt_monto3.Text) & ",'" & "BOB" & "'," & DTPicker1.Value & "," & DTPicker2.Value & " , '-', '-', '-', '-', '-', '-', '-', '-', 'REG', '" & glusuario & "','" & Date & "')"
    Else
        Set rs_guardar = New ADODB.Recordset
        If rs_guardar.State = 1 Then rs_guardar.Close
        rs_guardar.Open "Select * from ro_rrhh_adjudica_personas WHERE beneficiario_codigo = '" & Ado_contratacion.Recordset!beneficiario_codigo & "' AND ges_gestion = '" & Ado_contratacion.Recordset!ges_gestion & "' AND adjudica_codigo = " & Ado_contratacion.Recordset!adjudica_codigo & "", db, adOpenKeyset, adLockOptimistic, adCmdText
    End If
    rs_guardar!ges_gestion = Year(DTPicker1.Value)
    'rs_guardar!adjudica_codigo = 0
    rs_guardar!genero_codigo = dtc_genero.Text
    rs_guardar!beneficiario_codigo = txtCI.Text
    rs_guardar!unidad_codigo = dtc_cod_5.Text
    rs_guardar!rrhh_codigo = rs_guardar!adjudica_codigo
    rs_guardar!solicitud_codigo = rs_guardar!adjudica_codigo
    rs_guardar!solicitud_tipo = dtc_cod_4.Text
    rs_guardar!nivel_educ_codigo = dtc_codigo3.Text
    rs_guardar!ocup_codigo = dtc_codigo2.Text
    rs_guardar!codigo_empresa = dtc_empresa_cod.Text
    
    Set rs_UNIDAD = New ADODB.Recordset
    If rs_UNIDAD.State = 1 Then rs_UNIDAD.Close
    rs_UNIDAD.Open "Select * from gc_unidad_ejecutora WHERE unidad_codigo = '" & dtc_cod_5.Text & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
    
    'db.Execute "UPDATE rc_puestos set puesto_vacante = 'NO', beneficiario_codigo = '" & Trim(txtCI.Text) & "', usr_codigo ='" & glusuario & "' WHERE cargo_codigo = " & dtc_cargo.Text & " AND puesto_codigo = '" & dtc_cod_6.Text & "' AND unidad_codigo = '" & dtc_cod_5.Text & "' "
    
    rs_guardar!puesto_codigo = dtc_cod_6.Text
    rs_guardar!cargo_codigo = dtc_cargo.Text
    
    rs_guardar!modalidad_codigo = detc_cod_3.Text
    rs_guardar!munic_codigo = dtc_codigo4
    rs_guardar!beneficiario_monto_adjudica_bs = txt_monto1.Text
    rs_guardar!beneficiario_monto_adjudica_dol = Val(txt_monto1.Text) / GlTipoCambioOficial
    
    rs_guardar!beneficiario_haber_mensual_bs = txt_monto2.Text
    rs_guardar!beneficiario_haber_mensual_dol = Val(txt_monto2.Text) / GlTipoCambioOficial
    rs_guardar!beneficiario_otro_mensual_bs = txt_monto3.Text
    rs_guardar!beneficiario_otro_mensual_dol = Val(txt_monto3.Text) / GlTipoCambioOficial
    'rs_guardar!beneficiario_tiempo_meses = Year(DTPicker1.Value)
    rs_guardar!tipo_moneda = "BOB"
    rs_guardar!beneficiario_fecha_inicio = DTPicker1.Value
    If dtc_cod_4.Text = "11" Then
        rs_guardar!beneficiario_fecha_fin = "31/12/2095"
        rs_guardar!beneficiario_tiempo_meses = "0"      'txt_tiempo.Text
        rs_guardar!doc_codigo = "R-170"
    Else
        'DateAdd(intervalo, número, fecha)
        'Txt_tiempo.Text = DateDiff("m", CDate(DTPicker1.Value), CDate(DTPicker2.Value))
        txt_tiempo.Text = IIf(txt_tiempo.Text = "0", Val(cmb_meses), txt_tiempo.Text)
        rs_guardar!beneficiario_fecha_fin = DateAdd("m", Val(txt_tiempo.Text), CDate(DTPicker1.Value))
        rs_guardar!beneficiario_tiempo_meses = Val(txt_tiempo.Text)
        rs_guardar!doc_codigo = "R-179"
    End If
    rs_guardar!beneficiario_fecha_adjudica = Date
    'rs_guardar!beneficiario_fecha_contrato = Year(DTPicker1.Value)
    rs_guardar!etapa_codigo = "ADM-01-01"
    rs_guardar!proceso_codigo = "ADM"
    rs_guardar!subproceso_codigo = "ADM-01"
    rs_guardar!clasif_codigo = "ADM"
    'rs_guardar!doc_codigo = "R-170"
    rs_guardar!doc_numero = rs_guardar!adjudica_codigo
    rs_guardar!cite_tramite = "DRRHH-" + Str(rs_guardar!adjudica_codigo)
    nomb2 = "CONTRATADO: " & VAR_NOMBRE
    rs_guardar!observaciones = nomb2
    rs_guardar!usr_codigo = glusuario
    rs_guardar!fecha_registro = Date
    rs_guardar.Update
    
    VAR_CITE = Str(rs_guardar!adjudica_codigo)
    VAR_ADJUDICA = rs_guardar!adjudica_codigo
    db.Execute "UPDATE rc_puestos set puesto_vacante = 'NO', beneficiario_codigo = '" & Trim(txtCI.Text) & "', usr_codigo ='" & glusuario & "' WHERE cargo_codigo = " & dtc_cargo.Text & " AND puesto_codigo = '" & dtc_cod_6.Text & "' AND unidad_codigo = '" & dtc_cod_5.Text & "' "
    db.Execute "UPDATE ro_rrhh_adjudica_personas set rrhh_codigo = adjudica_codigo, doc_numero =adjudica_codigo, cite_tramite= 'DRRHH-'+'" & RTrim(VAR_CITE) & "' WHERE rrhh_codigo = '0' "
    
 'db.Execute "Insert INTO  ro_rrhh_adjudica_personas (ges_gestion,rrhh_codigo, beneficiario_codigo, unidad_codigo, solicitud_codigo, puesto_codigo, beneficiario_haber_mensual_bs, beneficiario_otro_mensual_bs, tipo_moneda, beneficiario_fecha_inicio, beneficiario_fecha_fin,proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, cite_tramite, observaciones, estado_codigo, usr_codigo, fecha_registro) values ('" & Year(Date) & "', " & rs_aux2.RecordCount + 1 & ",'" & txtCI.Text & "','" & dtc_cod_5.Text & "'," & rs_aux2.RecordCount + 1 & ",'" & dtc_cod_6.Text & "'," & Val(txt_monto2.Text) & "," & Val(txt_monto3.Text) & ",'" & "BOB" & "'," & DTPicker1.Value & "," & DTPicker2.Value & " , '-', '-', '-', '-', '-', '-', '-', '-', 'REG', '" & glusuario & "','" & Date & "')"
    Para_Aceptado = "N"
    Nuevo = "NO"
    ' txtSW = "0"
    Fra_datos.Enabled = False
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FraNavega.Enabled = True
    dg_datos.Enabled = True
    Option1.Value = False
    Option2.Value = False
    OptFilGral1.Value = True
 
    Set rs_aux_5 = New ADODB.Recordset
    If rs_aux_5.State = 1 Then rs_aux_5.Close
    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
    rs_aux_5.Open "SELECT * FROM rc_puestos WHERE estado_codigo <> 'ANL' AND puesto_vacante = 'SI' ", db, adOpenStatic
    Set Ado_aux_5.Recordset = rs_aux_5
 
    Call OptFilGral1_Click
 End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub GRABA_CALIFICA()
    db.Execute "Insert INTO ro_rrhh_apertura_sobres (ges_gestion, rrhh_codigo, beneficiario_codigo, unidad_codigo, solicitud_codigo, observaciones, puesto_codigo, ocup_codigo, nivel_educ_codigo, munic_codigo, estado_codigo, usr_codigo, fecha_registro, modalidad_codigo, cotiza_codigo) Values ('" & glGestion & "', '" & frm_ao_contratacion.Ado_datos.Recordset!rrhh_codigo & "',  '" & txtCI.Text & "', '" & txt_campo1.Text & "', '" & txt_codigo.Caption & "', '" & nomb2 & "', '" & dtc_codigo1.Text & "', " & dtc_codigo2.Text & ", " & dtc_codigo3.Text & ", '" & dtc_codigo4.Text & "', 'REG', '" & glusuario & "',  '" & Date & "', '" & frm_ao_contratacion.Ado_datos.Recordset!modalidad_codigo & "', " & Val(lbl_convoca.Caption) & ")"
End Sub

Function Valida()
'valida que el monto asignado al beneficiario no sobrepase el monto pendiente de asignacion
      Valida = True

'    If (dtc_depto_codigo.Text = "") Then
'      MsgBox "Debe registrar ... " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
'      Valida = False
'    End If
    'PUESTO SOLICITADO
'    If (dtc_codigo1.Text = "") Then
'      MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
'      Valida = False
'    End If
    
    'Perfil Profesional
    If (dtc_codigo2.Text = "") Then
      MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
      Valida = False
    End If
    'Nivel Educacional (Mayor Importancia)
    If (dtc_codigo3.Text = "") Then
      MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
      Valida = False
    End If
    'Lugar de Postulación
    If (dtc_codigo4.Text = "") Then
      MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
      Valida = False
    End If
    'PRIMER APELLIDO
    If txtPat = "" Then
        MsgBox "Debe registrar el PRIMER APELLIDO ... ", vbCritical + vbExclamation, "Validación de datos"
        Valida = False
    End If
    'NOMBRE(S)
    If txtNom = "" Then
        MsgBox "Debe registrar el(los) NOMBRE(S) ... ", vbCritical + vbExclamation, "Validación de datos"
        Valida = False
    End If
    'Modalidad de Contratación
    If (detc_cod_3.Text = "") Then
      MsgBox "Debe registrar ... " + lbl_mod.Caption, vbCritical + vbExclamation, "Validación de datos"
      Valida = False
    End If
    'Tipo de Contrato
    If (dtc_cod_4.Text = "") Then
      MsgBox "Debe registrar ... " + Label3.Caption, vbCritical + vbExclamation, "Validación de datos"
      Valida = False
    End If
    'Departamento (Oficina)
    If (dtc_cod_5.Text = "") Then
      MsgBox "Debe registrar ... " + Label5.Caption, vbCritical + vbExclamation, "Validación de datos"
      Valida = False
    End If
    'PUESTO A OCUPAR (REGISTRADO)
    If (dtc_cod_6.Text = "") Then
      MsgBox "Debe registrar ... " + LblPuesto.Caption, vbCritical + vbExclamation, "Validación de datos"
      Valida = False
    End If
     
    If (txt_monto2.Text = "") Then
      MsgBox "Debe registrar ... " + "Sueldo Mensual", vbCritical + vbExclamation, "Validación de datos"
      Valida = False
    End If
    
    If (dtc_depto_codigo.Text = "") Then
      MsgBox "Debe registrar ... " + "La extencion de CI", vbCritical + vbExclamation, "Validación de datos"
      Valida = False
    End If
    
    If (dtc_genero.Text = "") Then
      MsgBox "Debe registrar ... " + "El genero de la persona", vbCritical + vbExclamation, "Validación de datos"
      Valida = False
    End If
    
    'Empresa
    If (dtc_empresa_den.Text = "") Then
     MsgBox "Debe registrar ... " + Label7.Caption, vbCritical + vbExclamation, "Validación de datos"
      Valida = False
    End If
End Function

Private Sub BtnModificar_Click()
    If Ado_contratacion.Recordset!estado_codigo = "REG" Then
        modif = "SI"
        Fra_datos.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        FraNavega.Enabled = False
        dg_datos.Enabled = False
        Frame4.Enabled = False
        
    '     Set rs_aux_5 = New ADODB.Recordset
    '    If rs_aux_5.State = 1 Then rs_aux_5.Close
    '    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
    '    rs_aux_5.Open "SELECT * FROM rc_puestos WHERE unidad_codigo = '" & dtc_cod_5.Text & "'", db, adOpenStatic
    '    Set Ado_aux_5.Recordset = rs_aux_5
        
    Else
        sino = MsgBox("No se puede modifocar un registro aprobado (APR) o anulado (ANL)", vbCritical, "Aviso")
    End If
End Sub

Private Sub BtnNo_Click()
'    Frame2.Visible = False
'    Frame3.Visible = False
    
    Frame4.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
    Option1.Value = True
End Sub

Private Sub BtnOk_Click()
    txtCI.Text = dtc_codigo5.Text
    txtPat.Text = Trim(dtc_aux1.Text)
    txtMat.Text = Trim(dtc_aux2.Text)
    txtNom.Text = Trim(dtc_aux3.Text)
    txtTelefono.Text = Trim(dtc_aux4.Text)
    txtDireccion.Text = Trim(dtc_aux5.Text)
    Frame2.Visible = False
    Frame4.Visible = True
    Frame3.Visible = False
End Sub

Private Sub BtnSalir_Click()
  Unload Me
End Sub

Private Sub cmb_meses_Click()
    If cmb_meses.Text = "0" Or cmb_meses.Text = "00" Then
        'txt_tiempo.Text = "Indefinido"
        txt_tiempo.Text = "0"
        LblTiempo.Caption = "INDEFINIDO"
    Else
        txt_tiempo.Text = cmb_meses.Text
        LblTiempo.Caption = "Tiempo (Meses)"
    End If
End Sub

Private Sub detc_cod_3_Click(Area As Integer)
    dtc_desc_3 = detc_cod_3.BoundText
End Sub

Private Sub dtc_aux1_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux1.BoundText
    dtc_desc5.BoundText = dtc_aux1.BoundText
    dtc_aux2.BoundText = dtc_aux1.BoundText
    dtc_aux3.BoundText = dtc_aux1.BoundText
    dtc_aux4.BoundText = dtc_aux1.BoundText
    dtc_aux5.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux2.BoundText
    dtc_desc5.BoundText = dtc_aux2.BoundText
    dtc_aux1.BoundText = dtc_aux2.BoundText
    dtc_aux3.BoundText = dtc_aux2.BoundText
    dtc_aux4.BoundText = dtc_aux2.BoundText
    dtc_aux5.BoundText = dtc_aux2.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux3.BoundText
    dtc_desc5.BoundText = dtc_aux3.BoundText
    dtc_aux1.BoundText = dtc_aux3.BoundText
    dtc_aux2.BoundText = dtc_aux3.BoundText
    dtc_aux4.BoundText = dtc_aux3.BoundText
    dtc_aux5.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux4.BoundText
    dtc_desc5.BoundText = dtc_aux4.BoundText
    dtc_aux1.BoundText = dtc_aux4.BoundText
    dtc_aux2.BoundText = dtc_aux4.BoundText
    dtc_aux3.BoundText = dtc_aux4.BoundText
    dtc_aux5.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_aux5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux5.BoundText
    dtc_desc5.BoundText = dtc_aux5.BoundText
    dtc_aux1.BoundText = dtc_aux5.BoundText
    dtc_aux2.BoundText = dtc_aux5.BoundText
    dtc_aux3.BoundText = dtc_aux5.BoundText
    dtc_aux4.BoundText = dtc_aux5.BoundText
End Sub

Private Sub dtc_cod_4_Click(Area As Integer)
dtc_desc_4.BoundText = dtc_cod_4.BoundText
End Sub

Private Sub dtc_cod_5_Change()
If Nuevo = "SI" Then
Set rs_aux_5 = New ADODB.Recordset
    If rs_aux_5.State = 1 Then rs_aux_5.Close
    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
    rs_aux_5.Open "SELECT * FROM rc_puestos where estado_codigo = 'APR' and puesto_vacante = 'SI' and estado_codigo = 'APR' AND unidad_codigo = '" & dtc_cod_5.Text & "' ORDER BY puesto_descripcion", db, adOpenStatic
    Set Ado_aux_5.Recordset = rs_aux_5
    dtc_des_6.Text = ""
Else
Set rs_aux_5 = New ADODB.Recordset
    If rs_aux_5.State = 1 Then rs_aux_5.Close
    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
    rs_aux_5.Open "SELECT * FROM rc_puestos where estado_codigo = 'APR'  and unidad_codigo = '" & dtc_cod_5.Text & "' ORDER BY puesto_descripcion", db, adOpenStatic
    Set Ado_aux_5.Recordset = rs_aux_5
    dtc_des_6.Text = ""
End If
End Sub

Private Sub dtc_cod_5_Click(Area As Integer)
   dtc_codigo5.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_cod_6_Click(Area As Integer)
  dtc_desc1.BoundText = dtc_codigo5.BoundText
  dtc_cargo.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    dtc_dpto_cod.BoundText = dtc_codigo4.BoundText
End Sub



Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux1.BoundText = dtc_codigo5.BoundText
    dtc_aux2.BoundText = dtc_codigo5.BoundText
    dtc_aux3.BoundText = dtc_codigo5.BoundText
    dtc_aux4.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_des_6_Click(Area As Integer)
    dtc_cod_6.BoundText = dtc_des_6.BoundText
     dtc_cargo.BoundText = dtc_des_6.BoundText
End Sub

Private Sub dtc_desc_3_Click(Area As Integer)
    detc_cod_3.BoundText = dtc_desc_3.BoundText
End Sub

Private Sub dtc_desc_4_Click(Area As Integer)
    dtc_cod_4.BoundText = dtc_desc_4.BoundText
End Sub

Private Sub dtc_desc_4_LostFocus()
    If dtc_cod_4.Text = "11" Then
        LblMeses.Visible = False
        cmb_meses.Visible = False
        cmb_meses.Text = "00"
        txt_tiempo.Text = "0"
    Else
        LblMeses.Visible = True
        cmb_meses.Visible = True

    End If
End Sub

Private Sub dtc_desc_5_Click(Area As Integer)
        dtc_cod_5.BoundText = dtc_desc_5.BoundText
'        If dtc_desc_5.Text <> "" Then
'
'
'     Set rs_aux_5 = New ADODB.Recordset
'    If rs_aux_5.State = 1 Then rs_aux_5.Close
'    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
'    rs_aux_5.Open "SELECT * FROM rc_cargos WHERE estado_codigo = 'APR'", db, adOpenStatic
'    Set Ado_aux_5.Recordset = rs_aux_5
'
'    End If
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc1_LostFocus()
    If txtSW = "IDIR" Then
        Option1.Visible = False
        Option2.Visible = False
        Frame3.Visible = True
    Else
        Option1.Visible = True
        Option2.Visible = True
        Frame3.Visible = False
    End If
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
      dtc_dpto_cod.BoundText = dtc_desc4.BoundText
End Sub


Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    dtc_aux1.BoundText = dtc_desc5.BoundText
    dtc_aux2.BoundText = dtc_desc5.BoundText
    dtc_aux3.BoundText = dtc_desc5.BoundText
    dtc_aux4.BoundText = dtc_desc5.BoundText
    dtc_aux5.BoundText = dtc_desc5.BoundText
End Sub




Private Sub dtc_dpto_cod_Click(Area As Integer)
dtc_desc4.BoundText = dtc_dpto_cod.BoundText
dtc_codigo4.BoundText = dtc_dpto_cod.BoundText
End Sub

Private Sub dtc_empresa_cod_Click(Area As Integer)
 dtc_empresa_den.BoundText = dtc_empresa_cod.BoundText
End Sub

Private Sub dtc_empresa_den_Click(Area As Integer)
 dtc_empresa_cod.BoundText = dtc_empresa_den.BoundText
End Sub

Private Sub DTPicker1_Change()
'txt_tiempo.Text = DateDiff("m", DTPicker1.Value, DTPicker2.Value)
'If txt_monto2.Text <> "" Then
'txt_monto1.Text = Val(txt_monto2.Text) * Val(txt_tiempo.Text)
'End If
End Sub

Private Sub DTPicker2_Change()
'txt_tiempo.Text = DateDiff("m", DTPicker1.Value, DTPicker2.Value)
'If txt_monto2.Text <> "" Then
'txt_monto1.Text = Val(txt_monto2.Text) * Val(txt_tiempo.Text)
'End If
End Sub

Private Sub Form_Load()
'rs_contratacion
 Call ABRIR_TABLA_AUX
 Call OptFilGral1_Click
 OptFilGral1.Value = True
  OptFilGral2.Value = False
 
  Base = "NO"
  Nuevo = "NO"

'If glProceso = "CONSULTORIA" Then
'    Me.Caption = "Consultoría - Captura de datos personales"
'Else
'    Me.Caption = "Recursos Humanos - Captura de datos personales"
'End If


'Para_Aceptado = "N"
'LOS DATOS PERSONALES SE CARGAN EN EL FORMULARIO QUE LO LLAMA
    'txtSW = "0"
    parametro = Aux
    dtc_desc1.Visible = False
    Option1.Visible = True
    Option2.Visible = True
    Frame5.Visible = True
    Option3.Visible = False
    Frame2.Visible = False
    Frame4.Visible = True
    Frame3.Visible = False
    

End Sub
Private Sub ABRIR_TABLA_AUX()
' Set rs_datos = New ADODB.Recordset
'   If rs_datos.State = 1 Then rs_datos.Close
'   queryinicial = "select * from gc_beneficiario WHERE  tipoben_codigo < 20 "
'   'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
'   rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
'   rs_datos.Sort = "beneficiario_denominacion"
'   Set Ado_datos.Recordset = rs_datos
   
    
'    Set rs_clasif1 = New ADODB.Recordset
'    If rs_clasif1.State = 1 Then rs_clasif1.Close
'    'rs_clasif1.Open "SELECT * FROM rc_puestos where puesto_vacante = 'SI' ORDER BY puesto_descripcion  ", DB, adOpenStatic
'    rs_clasif1.Open "SELECT * FROM rv_puestos_solicitud where unidad_codigo_sol = '" & parametro & "' and solicitud_codigo = " & txt_codigo.Caption & " ORDER BY puesto_descripcion ", db, adOpenStatic
'    Set Ado_clasif1.Recordset = rs_clasif1
    
    Set rs_clasif2 = New ADODB.Recordset
    If rs_clasif2.State = 1 Then rs_clasif2.Close
    rs_clasif2.Open "SELECT * FROM gc_ocupacion_profesion ORDER BY ocup_descripcion ", db, adOpenStatic
    Set Ado_clasif2.Recordset = rs_clasif2
    
    Set rs_clasif3 = New ADODB.Recordset
    If rs_clasif3.State = 1 Then rs_clasif3.Close
    rs_clasif3.Open "SELECT * FROM rc_nivel_educacional ORDER BY nivel_educ_descripcion ", db, adOpenStatic
    Set Ado_clasif3.Recordset = rs_clasif3
    
    Set rs_clasif4 = New ADODB.Recordset
    If rs_clasif4.State = 1 Then rs_clasif4.Close
    rs_clasif4.Open "SELECT * FROM gc_municipio where region_codigo = 'SI' ORDER BY munic_descripcion ", db, adOpenStatic
    Set Ado_clasif4.Recordset = rs_clasif4
    
    Set rs_clasif5 = New ADODB.Recordset
    If rs_clasif5.State = 1 Then rs_clasif5.Close
    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
    rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE tipoben_codigo < '20' and estado_codigo <> 'ANL' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_clasif5.Recordset = rs_clasif5
    
'     Set rs_clasif5 = New ADODB.Recordset
'    If rs_clasif5.State = 1 Then rs_clasif5.Close
'    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
'    rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE tipoben_codigo < '20' and estado_codigo = 'APR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
'    Set Ado_clasif5.Recordset = rs_clasif5
    
'    Set rs_aux_1 = New ADODB.Recordset
'    If rs_aux_1.State = 1 Then rs_aux_1.Close
'    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
'    rs_aux_1.Open "SELECT * FROM gc_beneficiario WHERE tipoben_codigo < '20' and estado_codigo = 'APR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
'    Set Ado_clasif5.Recordset = rs_aux_1
    
    Set rs_aux_1 = New ADODB.Recordset
    If rs_aux_1.State = 1 Then rs_aux_1.Close
    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
    rs_aux_1.Open "SELECT * FROM rc_modalidad_contratacion WHERE compras_o_rrhh = 'R' and estado_codigo <> 'ANL'", db, adOpenStatic
    Set Ado_aux_1.Recordset = rs_aux_1
    
     Set rs_aux_2 = New ADODB.Recordset
    If rs_aux_2.State = 1 Then rs_aux_2.Close
    'rs_aux_2.Open "SELECT * FROM gc_tipo_solicitud WHERE (solicitud_num > '0' and solicitud_num < '9') and estado_codigo = 'APR'", db, adOpenStatic
    rs_aux_2.Open "SELECT * FROM gc_tipo_solicitud WHERE (unidad_codigo = 'DRRHH') and estado_codigo = 'APR'", db, adOpenStatic
    Set Ado_aux_2.Recordset = rs_aux_2
    
'      Set rs_aux_3 = New ADODB.Recordset
'    If rs_aux_3.State = 1 Then rs_aux_3.Close
'    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
'    rs_aux_3.Open "SELECT * FROM gc_unidad_ejecutora WHERE estado_codigo <> 'ANL' ORDER BY unidad_descripcion ", db, adOpenStatic
'    Set Ado_aux_3.Recordset = rs_aux_3

     Set rs_aux_4 = New ADODB.Recordset
    If rs_aux_4.State = 1 Then rs_aux_4.Close
    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
    rs_aux_4.Open "SELECT * FROM gc_unidad_ejecutora WHERE estado_codigo <> 'ANL' ORDER BY unidad_descripcion", db, adOpenStatic
    Set Ado_aux_4.Recordset = rs_aux_4
    
     Set rs_aux_6 = New ADODB.Recordset
    If rs_aux_6.State = 1 Then rs_aux_6.Close
    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
    rs_aux_6.Open "SELECT * FROM gc_departamento WHERE estado_codigo <> 'ANL'", db, adOpenStatic
    Set Ado_aux_6.Recordset = rs_aux_6

    
     Set rs_aux_7 = New ADODB.Recordset
    If rs_aux_7.State = 1 Then rs_aux_7.Close
    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
    rs_aux_7.Open "SELECT * FROM gc_genero WHERE estado_codigo = 'APR'", db, adOpenStatic
    Set Ado_aux_7.Recordset = rs_aux_7
    
    'PUESTOS
'    Set rs_aux_5 = New ADODB.Recordset
'    If rs_aux_5.State = 1 Then rs_aux_5.Close
'    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
'    rs_aux_5.Open "SELECT * FROM rc_puestos where estado_codigo = 'APR' and puesto_vacante = 'SI' ORDER BY puesto_descripcion", db, adOpenStatic
'    Set Ado_aux_5.Recordset = rs_aux_5
'rs_aux_8
    'Empresas
    Set rs_aux_8 = New ADODB.Recordset
    If rs_aux_8.State = 1 Then rs_aux_8.Close
    rs_aux_8.Open "SELECT * FROM gc_empresas where estado_codigo = 'APR' order by denominacion_empresa", db, adOpenStatic
    Set Ado_aux_8.Recordset = rs_aux_8
End Sub

Private Sub OptFilGral1_Click()
   Set rs_contratacion = New ADODB.Recordset
   If rs_contratacion.State = 1 Then rs_contratacion.Close
   queryinicial = "select * from rv_contrato_persona WHERE estado_codigo = 'REG' order by beneficiario_denominacion"
   modif = "select * from rv_contrato_persona"
   'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
   rs_contratacion.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
   rs_contratacion.Sort = "beneficiario_denominacion"
   Set Ado_contratacion.Recordset = rs_contratacion
   Set dg_datos.DataSource = Ado_contratacion.Recordset
   If Ado_contratacion.Recordset.RecordCount > 0 Then
        Ado_contratacion.Recordset.MoveFirst
   End If
     
End Sub

Private Sub OptFilGral2_Click()
   Set rs_contratacion = New ADODB.Recordset
   If rs_contratacion.State = 1 Then rs_contratacion.Close
   queryinicial = "select * from rv_contrato_persona order by beneficiario_denominacion"
   modif = "select * from rv_contrato_persona"
   'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
   rs_contratacion.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
   rs_contratacion.Sort = "beneficiario_denominacion"
   Set Ado_contratacion.Recordset = rs_contratacion
   Set dg_datos.DataSource = Ado_contratacion.Recordset
   Ado_contratacion.Recordset.MoveFirst
End Sub

Private Sub Option1_Click()
    Frame4.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
    
    If Base = "SI" Then
        txtCI.Text = ""
        dtc_depto_codigo.Text = ""
        txtPat.Text = ""
        txtMat.Text = ""
        txtNom.Text = ""
        dtc_genero.Text = ""
        txtTelefono.Text = ""
        txtDireccion.Text = ""
    End If
    
    Base = "NO"
    
'    txtSW = "1"
End Sub

Private Sub Option2_Click()
    Frame2.Visible = True
    Frame4.Visible = False
    Frame3.Visible = False
    Base = "SI"
'    txtSW = "2"
End Sub

Private Sub Option3_Click()
    Set rs_clasif1 = New ADODB.Recordset
    If rs_clasif1.State = 1 Then rs_clasif1.Close
    rs_clasif1.Open "SELECT * FROM rv_puestos_solicitud where unidad_codigo_sol = '" & parametro & "' and solicitud_codigo = " & txt_codigo.Caption & " ORDER BY puesto_descripcion ", db, adOpenStatic
    Set Ado_clasif1.Recordset = rs_clasif1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    'puesto2 = dtc_codigo1.Text
    dtc_desc1.Visible = True
    Option1.Visible = True
    Option2.Visible = True
    Frame5.Visible = True
    Option3.Visible = False
End Sub

Private Sub txt_monto2_KeyUp(KeyCode As Integer, Shift As Integer)
    txt_monto1.Text = Val(txt_monto2.Text) * Val(txt_tiempo.Text)
End Sub

Private Sub txt_monto2_LostFocus()
    If txt_monto2.Text <> "" And txt_tiempo.Text <> "" Then
        txt_monto1.Text = txt_monto2.Text * txt_tiempo.Text
    End If
End Sub

Private Sub txtFecha2_LostFocus()
    txtFecha.Value = txtFecha2.Value
    'Me.Print Format(DateDiff("y", Fecha_Inicial, Fecha_Final), Formato) & " dias"
    VAR_TIME = DateDiff("y", txtFecha3, txtFecha2)
    If Val(VAR_TIME) < 0 Then
        MsgBox "La Fecha Límite Postulación NO puede ser MENOR a la Fecha Inicio Convocatoria, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
        txtFecha2.SetFocus
    End If
End Sub





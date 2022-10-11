VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmLMayorAux 
   Caption         =   "Reportes Contables - Libro Mayor Auxiliar"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8835
   Icon            =   "FrmLMayorAux.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adoconvenio 
      Height          =   330
      Left            =   3540
      Top             =   7140
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "adoconvenio"
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
   Begin Crystal.CrystalReport CRyAux12 
      Left            =   660
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CryOrg 
      Left            =   240
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Fra_Busqueda 
      BackColor       =   &H00C0E0FF&
      Caption         =   "                           "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1545
      Left            =   1320
      TabIndex        =   27
      Top             =   5520
      Visible         =   0   'False
      Width           =   6435
      Begin VB.CommandButton cmdEjecutar 
         Caption         =   "Ejecutar"
         Height          =   300
         Left            =   840
         TabIndex        =   38
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_BSalir 
         Caption         =   "Salir"
         Height          =   300
         Left            =   4440
         TabIndex        =   37
         Top             =   1080
         Width           =   1050
      End
      Begin VB.CommandButton Cmd_Normal 
         Caption         =   "Normal"
         Height          =   300
         Left            =   2700
         TabIndex        =   36
         Top             =   1080
         Width           =   1125
      End
      Begin VB.ComboBox CboCampo 
         Height          =   315
         ItemData        =   "FrmLMayorAux.frx":0A02
         Left            =   240
         List            =   "FrmLMayorAux.frx":0A0C
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   480
         Width           =   1350
      End
      Begin VB.ComboBox CboOperador 
         Height          =   315
         ItemData        =   "FrmLMayorAux.frx":0A26
         Left            =   1920
         List            =   "FrmLMayorAux.frx":0A30
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox TxtValor 
         Height          =   336
         Left            =   3060
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   480
         Width           =   3060
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Búsqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   240
         TabIndex        =   39
         Top             =   0
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Búsqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   205
         TabIndex        =   40
         Top             =   0
         Width           =   1065
      End
   End
   Begin Crystal.CrystalReport CryLMayorCtaBancaria 
      Left            =   1500
      Top             =   4860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CryLMayorBenef 
      Left            =   1080
      Top             =   5040
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
      TabIndex        =   25
      Top             =   7500
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3810
      Left            =   60
      TabIndex        =   24
      Top             =   900
      Width           =   1125
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H8000000A&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   120
         Picture         =   "FrmLMayorAux.frx":0A3D
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Busca un Registro"
         Top             =   360
         Width           =   885
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H8000000A&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   120
         Picture         =   "FrmLMayorAux.frx":0FF5
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Imprime Formulario"
         Top             =   1560
         Width           =   885
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H8000000A&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   120
         Picture         =   "FrmLMayorAux.frx":15B2
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2760
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   3705
         Left            =   60
         Picture         =   "FrmLMayorAux.frx":17BC
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1020
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   60
      TabIndex        =   23
      Top             =   -60
      Width           =   8655
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reportes Contables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   360
         Left            =   5700
         TabIndex        =   32
         Top             =   240
         Width           =   2775
      End
      Begin VB.Image Image3 
         Height          =   840
         Left            =   0
         Picture         =   "FrmLMayorAux.frx":3DC6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15360
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3810
      Left            =   1260
      TabIndex        =   12
      Top             =   900
      Width           =   7485
      Begin MSDataListLib.DataCombo DtcCodAux3 
         Height          =   285
         Left            =   1980
         TabIndex        =   44
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ListField       =   "codigo_convenio"
         BoundColumn     =   "denominacion_convenio"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcDenomAux3 
         Height          =   285
         Left            =   3480
         TabIndex        =   45
         Top             =   2760
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   503
         _Version        =   393216
         ListField       =   "denominacion_convenio"
         BoundColumn     =   "codigo_convenio"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcCodAux2 
         DataField       =   "codigo_convenio"
         Height          =   285
         Left            =   1980
         TabIndex        =   42
         Top             =   2220
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ListField       =   "codigo_convenio"
         BoundColumn     =   "denominacion_convenio"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Crystal.CrystalReport CryConv_Conv 
         Left            =   540
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox cboCtaBancaria 
         Height          =   315
         Left            =   2040
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   2220
         Visible         =   0   'False
         Width           =   4275
      End
      Begin VB.ComboBox cbosubcta1 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   780
         Width           =   1140
      End
      Begin VB.ComboBox cbosubcta2 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   1230
         Width           =   1140
      End
      Begin VB.ComboBox cbocta 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   300
         Width           =   1140
      End
      Begin VB.CheckBox Chkaux1 
         Caption         =   "Auxiliar 1"
         Height          =   195
         Left            =   255
         TabIndex        =   4
         Top             =   1845
         Width           =   975
      End
      Begin VB.CheckBox Chkaux2 
         Caption         =   "Auxiliar 2"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   2325
         Width           =   1005
      End
      Begin VB.CheckBox Chkaux3 
         Caption         =   "Auxiliar 3"
         Height          =   270
         Left            =   240
         TabIndex        =   7
         Top             =   2775
         Width           =   1080
      End
      Begin VB.TextBox txtax1 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1344
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1740
         Width           =   585
      End
      Begin VB.TextBox Txtax2 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2220
         Width           =   585
      End
      Begin VB.TextBox txtax3 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2715
         Width           =   585
      End
      Begin MSComCtl2.DTPicker DTPfin 
         Height          =   360
         Left            =   3480
         TabIndex        =   9
         Top             =   3225
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   635
         _Version        =   393216
         Format          =   62849025
         CurrentDate     =   36614
      End
      Begin MSComCtl2.DTPicker DTPinicio 
         Height          =   345
         Left            =   1320
         TabIndex        =   8
         Top             =   3240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         _Version        =   393216
         Format          =   62849025
         CurrentDate     =   37257
         MaxDate         =   2958101
         MinDate         =   36892
      End
      Begin MSDataListLib.DataCombo DTCNomOrg 
         Height          =   315
         Left            =   3120
         TabIndex        =   29
         Top             =   2760
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtCOrg 
         Height          =   315
         Left            =   2040
         TabIndex        =   28
         Top             =   2760
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdodcOrganismo 
         Height          =   330
         Left            =   4920
         Top             =   2580
         Visible         =   0   'False
         Width           =   2115
         _ExtentX        =   3731
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
         Caption         =   "Adodc1"
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
      Begin MSDataListLib.DataCombo DtcDenomAux2 
         Height          =   285
         Left            =   3480
         TabIndex        =   43
         Top             =   2220
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   503
         _Version        =   393216
         ListField       =   "denominacion_convenio"
         BoundColumn     =   "codigo_convenio"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcCodAux1 
         DataField       =   "codigo_beneficiario"
         Height          =   285
         Left            =   1980
         TabIndex        =   46
         Top             =   1740
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo_beneficiario"
         BoundColumn     =   "denominacion_beneficiario"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcDenomAux1 
         DataField       =   "denominacion_beneficiario"
         Height          =   285
         Left            =   3420
         TabIndex        =   47
         Top             =   1740
         Visible         =   0   'False
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   503
         _Version        =   393216
         Style           =   2
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtbusca1 
         Height          =   330
         Left            =   2040
         TabIndex        =   5
         Top             =   1740
         Width           =   5280
      End
      Begin MSDataListLib.DataCombo DtCIdConvenio 
         Bindings        =   "FrmLMayorAux.frx":49A08
         Height          =   285
         Left            =   2040
         TabIndex        =   30
         Top             =   1740
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo_convenio"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtCDesConvenio 
         Bindings        =   "FrmLMayorAux.frx":49A22
         Height          =   285
         Left            =   3600
         TabIndex        =   31
         Top             =   1740
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   503
         _Version        =   393216
         Style           =   2
         ListField       =   "Denominacion_Convenio"
         BoundColumn     =   "codigo_Convenio"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Lblsub1 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2760
         TabIndex        =   22
         Top             =   840
         Width           =   30
      End
      Begin VB.Label lblcuenta 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2760
         TabIndex        =   21
         Top             =   360
         Width           =   30
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Subcuenta 2"
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Subcuenta 1"
         Height          =   195
         Left            =   210
         TabIndex        =   19
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lbsub2 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2760
         TabIndex        =   17
         Top             =   1260
         Width           =   30
      End
      Begin VB.Label Label4 
         Caption         =   "Desde:"
         Height          =   240
         Left            =   300
         TabIndex        =   16
         Top             =   3315
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Hasta:"
         Height          =   240
         Left            =   2880
         TabIndex        =   15
         Top             =   3330
         Width           =   645
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2640
      Left            =   120
      TabIndex        =   0
      Top             =   4740
      Width           =   8700
      Begin MSDataGridLib.DataGrid DTGBanco 
         Height          =   2295
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   4048
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
      Begin Crystal.CrystalReport CryBenefConvenios 
         Left            =   7200
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSDataGridLib.DataGrid DtGbenef 
         Height          =   2370
         Left            =   120
         TabIndex        =   11
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
      Left            =   1140
      Top             =   4740
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
'/****** REFORMULADO EL 20 de junio
'/************  RECORDSETS

Dim sql1 As String
Dim sql2 As String
Dim lcta As String
Dim nombenef As String
Dim combenef As ADODB.Command
Dim comAux12 As ADODB.Command
Dim comORG As ADODB.Command
Dim rsOrganismo As ADODB.Recordset
Dim comctabancaria As ADODB.Command
Dim rsplanctas As ADODB.Recordset
Dim rscuentas As ADODB.Recordset
Dim rsnombresub1 As ADODB.Recordset
Dim rssubcuenta As ADODB.Recordset
Dim rscta_bancaria As ADODB.Recordset
Dim rsbeneficiario As ADODB.Recordset
Dim rssaldos As ADODB.Recordset
Dim rsctabancaria As ADODB.Recordset
Dim rsConvenio As ADODB.Recordset
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
Public AUX2 As String
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
  Me.lblcuenta = rsplanctas!NombreCta
  If rscuentas.State = adStateOpen Then rscuentas.Close
  
  rscuentas.Open "SELECT Cuenta, SubCta1 FROM CC_Plan_Cuentas GROUP BY Cuenta, SubCta1 HAVING (SubCta1 <> '00') AND (Cuenta = '" & Trim(Me.cbocta.Text) & "')", db, adOpenKeyset, adLockReadOnly
  Do While Not rscuentas.EOF
    Me.cbosubcta1.AddItem rscuentas!subcta1
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
        Me.cbosubcta2.AddItem rssubcuenta!subcta2
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
On Error GoTo labelerr2
  txtbusca1.Visible = True
  DtCIdConvenio.Visible = True
  DtCDesConvenio.Visible = True
  With rssubcuenta
    .MoveFirst
    .Find "subcta2=" & "'" & Trim(Me.cbosubcta2) & "'"
    Me.lbsub2 = !NombreCta
    Me.txtax1 = !aux1
    Me.Txtax2 = !AUX2
    Me.txtax3 = !aux3
    Chkaux1.Enabled = True
    Chkaux2.Enabled = True
    Chkaux3.Enabled = True
    Chkaux1.Value = 1
    Chkaux2.Value = 1
    Chkaux3.Value = 1
    BtnBuscar.Enabled = False
    '--------
    txtbusca1.Visible = False
    cboCtaBancaria.Visible = False
    DtCDesConvenio.Visible = False
    DtCIdConvenio.Visible = False
    DtCOrg.Visible = False
    DTCNomOrg.Visible = False
    Select Case !aux1
      Case "00"
        'SSTabCuenta.TabEnabled(0) = False
        txtbusca1.Visible = False
        'cboCtaBancaria.Visible = False
        'DtCDesConvenio.Visible = False
        'DtCIdConvenio.Visible = False
        'DtcOrg.Visible = False
        'DTCNomOrg.Visible = False
        Chkaux1.Enabled = False
        Chkaux1.Value = 0
      Case "01"
        txtbusca1.Visible = True
        txtbusca1.Top = 1740
        'cboCtaBancaria.Visible = False
        'DtCDesConvenio.Visible = False
        'DtCIdConvenio.Visible = False
        'DtcOrg.Visible = False
        'DTCNomOrg.Visible = False
        Me.BtnBuscar.Enabled = True
      Case "02"
        cboCtaBancaria.Visible = True
        cboCtaBancaria.Top = 1740
        'txtbusca1.Visible = False
        'DtCDesConvenio.Visible = False
        'DtCIdConvenio.Visible = False
        'DtcOrg.Visible = False
        'DTCNomOrg.Visible = False
      Case "03"
        txtbusca1.Visible = True
        txtbusca1.Top = 1740
        Chkaux1.Value = 1
      Case "08"
        DTCNomOrg.Visible = True
        DTCNomOrg.Top = 1740
        DtCOrg.Visible = True
        DtCOrg.Top = 1740
        'txtbusca1.Visible = False
        'cboCtaBancaria.Visible = False
        'DtCDesConvenio.Visible = False
        'DtCIdConvenio.Visible = False
      Case "09"
        DtCIdConvenio.Visible = True
        DtCIdConvenio.Top = 1740
        DtCDesConvenio.Visible = True
        DtCDesConvenio.Top = 1740
        'txtbusca1.Visible = False
        'cboCtaBancaria.Visible = False
        'DtcOrg.Visible = False
        'DTCNomOrg.Visible = False
    End Select
    Select Case !AUX2
      Case "00"
        'SSTabCuenta.TabEnabled(0) = False
        txtbusca1.Visible = False
        'cboCtaBancaria.Visible = False
        'DtCDesConvenio.Visible = False
        'DtCIdConvenio.Visible = False
        'DtcOrg.Visible = False
        DTCNomOrg.Visible = False
        Chkaux2.Enabled = False
        Chkaux2.Value = 0
      Case "01"
        txtbusca1.Visible = True
        txtbusca1.Top = 2580
        'cboCtaBancaria.Visible = False
        'DtCDesConvenio.Visible = False
        'DtCIdConvenio.Visible = False
        'DtcOrg.Visible = False
        'DTCNomOrg.Visible = False
        Me.BtnBuscar.Enabled = True
      Case "02"
        cboCtaBancaria.Visible = True
        cboCtaBancaria.Top = 2220
        'txtbusca1.Visible = False
        'DtCDesConvenio.Visible = False
        'DtCIdConvenio.Visible = False
        'DtcOrg.Visible = False
        'DTCNomOrg.Visible = False
      Case "08"
        DTCNomOrg.Visible = True
        DTCNomOrg.Top = 2220
        DtCOrg.Visible = True
        DtCOrg.Top = 2220
        'txtbusca1.Visible = False
        'cboCtaBancaria.Visible = False
        'DtCDesConvenio.Visible = False
        'DtCIdConvenio.Visible = False
      Case "09"
        DtCIdConvenio.Visible = True
        DtCIdConvenio.Top = 2220
        DtCDesConvenio.Visible = True
        DtCDesConvenio.Top = 2220
        'txtbusca1.Visible = False
        'cboCtaBancaria.Visible = False
        'DtcOrg.Visible = False
        'DTCNomOrg.Visible = False
    End Select
    Select Case !aux3
      Case "00"
        'SSTabCuenta.TabEnabled(0) = False
        'txtbusca1.Visible = False
        'cboCtaBancaria.Visible = False
        'DtCDesConvenio.Visible = False
        'DtCIdConvenio.Visible = False
        'DtcOrg.Visible = False
        'DTCNomOrg.Visible = False
        Chkaux3.Enabled = False
        Chkaux3.Value = 0
      Case "01"
        txtbusca1.Visible = True
        txtbusca1.Top = 3120
        'cboCtaBancaria.Visible = False
        'DtCDesConvenio.Visible = False
        'DtCIdConvenio.Visible = False
        'DtcOrg.Visible = False
        'DTCNomOrg.Visible = False
        Me.BtnBuscar.Enabled = True
      Case "02"
        cboCtaBancaria.Visible = True
        cboCtaBancaria.Top = 3120
        'txtbusca1.Visible = False
        'DtCDesConvenio.Visible = False
        'DtCIdConvenio.Visible = False
        'DtcOrg.Visible = False
        'DTCNomOrg.Visible = False
      Case "08"
        DTCNomOrg.Visible = True
        DTCNomOrg.Top = 3120
        DtCOrg.Visible = True
        DtCOrg.Top = 3120
        'txtbusca1.Visible = False
        'cboCtaBancaria.Visible = False
        'DtCDesConvenio.Visible = False
        'DtCIdConvenio.Visible = False
      Case "09"
        DtCIdConvenio.Visible = True
        DtCIdConvenio.Top = 2760
        DtCDesConvenio.Visible = True
        DtCDesConvenio.Top = 2760
        'txtbusca1.Visible = False
        'cboCtaBancaria.Visible = False
        'DtcOrg.Visible = False
        'DTCNomOrg.Visible = False
    End Select
  End With
  'SSTabCuenta_Click (0)
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
          cboCtaBancaria.AddItem rscta_bancaria!Cta_Codigo
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
        If rsbeneficiario.State = 1 Then rsbeneficiario.Close
        sql2 = "SELECT codigo_beneficiario, denominacion_beneficiario From fc_beneficiario order by denominacion_beneficiario"
        rsbeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
        Set Me.DtGbenef.DataSource = rsbeneficiario
        Me.DtGbenef.Visible = True
        Me.DTGBanco.Visible = False
        Me.txtbusca1.Visible = True
        Me.BtnBuscar.Enabled = True
        Me.cboCtaBancaria.Visible = False
    End If
'****habilitamos boton de búsqueda
    
    If Me.txtax1 = "00" Or Me.txtax1 = "02" Then
        Me.BtnBuscar.Enabled = False
    Else
        Me.BtnBuscar.Enabled = True
    End If
    If Me.txtax1 = "03" Then
      Me.txtbusca1.Visible = True
    End If
    '-------- habilito datacombos para organismo financiadores
  If Trim(txtax1) = "01" And Trim(Txtax2) = "09" And Trim(txtax3) = "09" Then
    txtbusca1.Visible = False
    DtCIdConvenio.Visible = False
    DtCDesConvenio.Visible = False
    DtcCodAux1.Visible = True
    DtcCodAux2.Visible = True
    DtcCodAux3.Visible = True
    DtcDenomAux1.Visible = True
    DtcDenomAux2.Visible = True
    DtcDenomAux3.Visible = True
  Else
    'txtbusca1.Visible = True
    'DtCIdConvenio.Visible = True
    'DtCDesConvenio.Visible = True
    DtcCodAux1.Visible = False
    DtcCodAux2.Visible = False
    DtcCodAux3.Visible = False
    DtcDenomAux1.Visible = False
    DtcDenomAux2.Visible = False
    DtcDenomAux3.Visible = False
  End If

    
    Exit Sub
labelerr2:
    If Err.Number = 3021 Then
      MsgBox "Elija una subcuenta", vbCritical + vbDefaultButton1
      Me.cbosubcta2.SetFocus
    End If

'-------- habilito datacombos para organismo financiadores
  If Trim(txtax1) = "01" And Trim(Txtax2) = "09" And Trim(txtax3) = "09" Then
    txtbusca1.Visible = False
    DtCIdConvenio.Visible = False
    DtCDesConvenio.Visible = False
    DtcCodAux1.Visible = True
    DtcCodAux2.Visible = True
    DtcCodAux3.Visible = True
    DtcDenomAux1.Visible = True
    DtcDenomAux2.Visible = True
    DtcDenomAux3.Visible = True
  Else
    'txtbusca1.Visible = True
    'DtCIdConvenio.Visible = True
    'DtCDesConvenio.Visible = True
    DtcCodAux1.Visible = False
    DtcCodAux2.Visible = False
    DtcCodAux3.Visible = False
    DtcDenomAux1.Visible = False
    DtcDenomAux2.Visible = False
    DtcDenomAux3.Visible = False
  End If








'--------------

'    Me.Chkaux2.Value = False
'    Me.Chkaux3.Value = False
'    DtCIdConvenio.Visible = False
'    DtCDesConvenio.Visible = False
'    Me.DTCNomOrg.Visible = False
'    Me.DtcOrg.Visible = False
'    Me.Txtbusca2.Visible = True
'    Me.BtnBuscar.Enabled = True
'    Me.Chkaux1.Enabled = True
'    Me.Chkaux2.Enabled = True
'    Me.Chkaux3.Enabled = True
'    Me.txtax1.Enabled = True
'    Me.Txtax2.Enabled = True
'    Me.txtax3.Enabled = True
'    Me.txtbusca1.Enabled = True
'    Me.Txtbusca2.Enabled = True
'    Me.Txtbusca3.Enabled = True
'    With rssubcuenta
'      .MoveFirst
'      .Find "subcta2=" & "'" & Trim(Me.cbosubcta2) & "'"
'      Me.lbsub2 = !NombreCta
'      Me.txtax1 = !aux1
'      Me.Txtax2 = !aux2
'      Me.txtax3 = !aux3
'      If !aux1 = "00" Then
'        Me.Chkaux1.Enabled = False
'        Me.txtax1.Enabled = False
'        Me.txtbusca1.Enabled = False
'      End If
'      If !aux2 = "00" Then
'        Me.Chkaux2.Enabled = False
'        Me.Txtax2.Enabled = False
'        Me.Txtbusca2.Enabled = False
'      End If
'      If !aux3 = "00" Then
'        Me.Chkaux3.Enabled = False
'        Me.txtax3.Enabled = False
'        Me.Txtbusca3.Enabled = False
'      End If
'
'      If Me.Chkaux1.Enabled = True And Me.Chkaux2.Enabled = False And Me.Chkaux3.Enabled = False Then
'        Me.Chkaux1.Value = 1
'      End If
'      If Me.Chkaux1.Enabled = False And Me.Chkaux2.Enabled = True And Me.Chkaux3.Enabled = False Then
'        Me.Chkaux2.Value = 1
'      End If
'      If Me.Chkaux1.Enabled = False And Me.Chkaux2.Enabled = False And Me.Chkaux3.Enabled = True Then
'        Me.Chkaux3.Value = 1
'      End If
'    End With
'
'    If (Me.txtax1 <> "00" And Me.txtax1 <> "01" And Me.txtax1 <> "02" And txtax1 <> "09" And txtax1 <> "08") Then
'      f = 1
'      Me.Chkaux1.Enabled = False
'      Me.txtax1.Enabled = False
'      Me.txtbusca1.Enabled = False
'      Me.DTCNomOrg.Visible = False
'      Me.DtcOrg.Visible = False
'    End If
'    If (Me.Txtax2 <> "00" And Me.Txtax2 <> "01" And Me.Txtax2 <> "02" And Me.Txtax2 <> "09") Then
'      f = 2
'      Me.Chkaux2.Enabled = False
'      Me.Txtax2.Enabled = False
'      Me.Txtbusca2.Enabled = False
'      Me.DTCNomOrg.Visible = False
'      Me.DtcOrg.Visible = False
'    End If
'    'g--
'    If (Me.Txtax2 = "08") Then
'      f = 8
'      Me.Chkaux2.Enabled = True
'      Me.Txtax2.Enabled = True
'      Me.Txtbusca2.Enabled = True
'      Me.Txtbusca2.Visible = False
'      Me.DTCNomOrg.Visible = True
'      Me.DtcOrg.Visible = True
'    End If
'
'    If (Me.txtax1 = "09") Then
'      f = 9
'      Me.Chkaux2.Enabled = True
'      Me.Txtax2.Enabled = True
'      Me.txtbusca1.Enabled = True
'      Me.txtbusca1.Visible = False
'      'Me.DTCNomOrg.Visible = True
'      'Me.DtcOrg.Visible = True
'      DtCDesConvenio.Visible = True
'      DtCIdConvenio.Visible = True
'    End If
'
'
'
'    'g--
'    If (Me.txtax3 <> "00" And Me.txtax3 <> "01" And Me.txtax3 <> "02") Then
'      f = 3
'      Me.Chkaux3.Enabled = False
'      Me.txtax3.Enabled = False
'      Me.Txtbusca3.Enabled = False
'      Me.DTCNomOrg.Visible = False
'      Me.DtcOrg.Visible = False
'    End If
'    'If f = 1 Or f = 2 Or f = 3 Then
'      '  MsgBox "Por el momento solo se trabaja con Auxiliares de Beneficiarios y Ctas. Corrientes", vbInformation + vbDefaultButton1, "Atencion"
'        Me.cbocta.SetFocus
'    'End If
'    If Me.Chkaux1.Enabled = False And Me.Chkaux2.Enabled = False And Me.Chkaux3.Enabled = False Then
'    Me.BtnBuscar.Enabled = False
'    Else
''    Me.BtnGrabar.Enabled = False
'    End If
'If (Me.cbosubcta1.Text) = "00" And Me.cbosubcta2.Text = "00" Then
'    'Me.BtnGrabar.Enabled = True
'End If
''*******Se filtra si la cuenta es de bancos....
'If Me.cbocta = "1111" And Me.cbosubcta1 = "02" Then
'    Select Case Me.cbosubcta2
'        Case "01"
'            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_bancos.Bco_descripcion_larga FROM fc_cuenta_bancaria INNER JOIN " & _
'                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '41' or fc_cuenta_bancaria.Fte_codigo = '10' order by fc_cuenta_bancaria.Cta_codigo"
'        Case "02"
'            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_bancos.Bco_descripcion_larga FROM fc_cuenta_bancaria INNER JOIN " & _
'                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '43' order by fc_cuenta_bancaria.Cta_codigo"
'        Case "03"
'            sql1 = " SELECT fc_cuenta_bancaria.Cta_codigo, fc_cuenta_bancaria.Cta_descripcion_larga,  fc_bancos.Bco_descripcion_larga FROM fc_cuenta_bancaria INNER JOIN " & _
'                    "fc_bancos ON  fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo where  fc_cuenta_bancaria.Fte_codigo = '80' order by fc_cuenta_bancaria.Cta_codigo"
'     End Select
'    Me.cboCtaBancaria.Clear
'    If rscta_bancaria.State = 1 Then rscta_bancaria.Close
'    rscta_bancaria.Open sql1, db, adOpenKeyset, adLockReadOnly
'    If rscta_bancaria.RecordCount <> 0 Then
'        rscta_bancaria.MoveFirst
'    End If
'        Do While Not rscta_bancaria.EOF
'          cboCtaBancaria.AddItem rscta_bancaria!cta_codigo
'          rscta_bancaria.MoveNext
'        Loop
'    Me.cboCtaBancaria.Visible = True
'    Me.cboCtaBancaria.Text = Me.cboCtaBancaria.List(0)
'    Me.txtbusca1.Visible = False
'    Me.DTGBanco.Visible = True
'    Me.DtGbenef.Visible = False
'    Set Me.DTGBanco.DataSource = rscta_bancaria
'End If
'
''************Se habilita la tabla de beneficiarios
'    If Me.txtax1 = "01" Or Me.Txtax2 = "01" Or Me.txtax3 = "01" Then
'        If rsBeneficiario.State = 1 Then rsBeneficiario.Close
'        sql2 = "SELECT codigo_beneficiario, denominacion_beneficiario From fc_beneficiario order by denominacion_beneficiario"
'        rsBeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
'        Set Me.DtGbenef.DataSource = rsBeneficiario
'        Me.DtGbenef.Visible = True
'        Me.DTGBanco.Visible = False
'        Me.txtbusca1.Visible = True
'        Me.BtnBuscar.Enabled = True
'        Me.cboCtaBancaria.Visible = False
'    End If
''****habilitamos boton de búsqueda
'    If Me.txtax1 = "00" Or Me.txtax1 = "02" Then
'        Me.BtnBuscar.Enabled = False
'    Else
'        Me.BtnBuscar.Enabled = True
'    End If
'
'    Exit Sub
'labelerr2:
'    If err.Number = 3021 Then
'      MsgBox "Elija una subcuenta", vbCritical + vbDefaultButton1
'      Me.cbosubcta2.SetFocus
'    End If
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
Private Sub Cmd_BSalir_Click()
Me.Fra_Busqueda.Visible = False
End Sub

Private Sub Cmd_Normal_Click()
  If rsbeneficiario.State = 1 Then rsbeneficiario.Close
    sql2 = "SELECT codigo_beneficiario, denominacion_beneficiario From fc_beneficiario order by denominacion_beneficiario"
    rsbeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
    Set Me.DtGbenef.DataSource = rsbeneficiario
End Sub

'Private Sub BtnGrabar_Click()
'Call existecta(Trim(Me.cbocta), Trim(Me.cbosubcta1), Trim(Me.cbosubcta2))
'If lcta = "S" Then
'    If Me.cbocta.Text = "" Then
'        MsgBox "Elija una cuenta", vbCritical + vbDefaultButton1
'        Me.cbocta.SetFocus
'        Exit Sub
'    End If
'    If Me.cbosubcta1.Text = "" Then
'        MsgBox "Elija una subcuenta", vbCritical + vbDefaultButton1
'        Me.cbosubcta1.SetFocus
'        Exit Sub
'    End If
'    If Me.cbosubcta2.Text = "" Then
'        MsgBox "Elija una subcuenta", vbCritical + vbDefaultButton1
'        Me.cbosubcta2.SetFocus
'        Exit Sub
'    End If
'    If Me.txtax1 = "02" Then
'        If Me.cboCtaBancaria.Text = "" Then
'            MsgBox "Elija una cuenta bancaria", vbCritical + vbDefaultButton1
'            Me.cboCtaBancaria.SetFocus
'            Exit Sub
'        End If
'    End If
'    If Me.txtax1 = "01" And Me.Chkaux1.Value = 1 Then
'        If Me.txtbusca1.Text = "" Then
'            MsgBox "Escriba un beneficiario", vbCritical + vbDefaultButton1
'            Me.txtbusca1.SetFocus
'            Exit Sub
'        End If
'    End If
'    If Me.txtax1 = "02" Or Txtax2 = "02" Or txtax3 = "02" Then
'        If Me.cboCtaBancaria = "" Then
'            MsgBox "Seleccione una Cuenta Bancaria", vbCritical + vbDefaultButton1
'            Exit Sub
'        End If
'    End If
''    If Me.txtax1 = "01" Or Txtax2 = "01" Or txtax3 = "01" Then
''        If txtbusca1.Text = "" Then
''            MsgBox "Introduzca un Beneficiario", vbCritical + vbDefaultButton1
''            Exit Sub
''        End If
''    End If
'    If (DTPinicio.Value > DTPfin.Value) Or (DTPfin.Value < DTPinicio.Value) Then
'        MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
'        Exit Sub
'    End If
'    If Me.txtax1 = "00" And Me.Txtax2 = "00" And Me.txtax3 = "00" Then
'    '****si la cuenta no tiene auxiliares
'        Call Mayor000
'    Else
'        '****llamada al store procedure de Saldos para beneficiarios "SaldoBenef
'        If Chkaux1.Value = 0 And Chkaux2.Value = 0 And Chkaux3.Value = 0 Then
'          MsgBox "Seleccione una opción", vbExclamation + vbDefaultButton1, "REPORTES"
'          Exit Sub
'        End If
'        '***** si el aux es 1
'           If Chkaux1.Value = 1 And Chkaux2.Value = 0 And Chkaux3.Value = 0 Then
'              Select Case Trim(txtax1.Text)
'                Case "01"
'
'                  reporteBeneficiario  'procedimiento para reporte con beneficiarios
'                Case "02"
'                  ReporteCtaBancaria
'                Case "08"
'                  'ReporteOrg   ' procedimiento para reporte con organismos
'                Case "09"
'                  reporteconvenio
'              End Select
'
'
'           End If
'
'           If Chkaux1.Value = 0 And Chkaux2.Value = 1 And Chkaux3.Value = 0 Then
'              Select Case Trim(Txtax2.Text)
'                Case "01"
'                  reporteBeneficiario  'procedimiento para reporte con beneficiarios
'                Case "02"
'                  ReporteCtaBancaria
'                Case "08"
'                '  ReporteOrg   ' procedimiento para reporte con organismos
'              End Select
'           End If
'
'           If Chkaux1.Value = 0 And Chkaux2.Value = 0 And Chkaux3.Value = 1 Then
'              Select Case Trim(txtax3.Text)
'                Case "01"
'                  reporteBeneficiario  'procedimiento para reporte con beneficiarios
'                Case "02"
'                  ReporteCtaBancaria
'                Case "08"
'                  ReporteOrg   ' procedimiento para reporte con organismos
'              End Select
'           End If
'
'           If Chkaux1.Value = 1 And Chkaux2.Value = 1 And Chkaux3.Value = 0 Then
'              If Trim(txtax1.Text) = "01" And Trim(Txtax2.Text) = "08" Then
'                  If rsBeneficiario.State = 1 Then rsBeneficiario.Close
'                  rsBeneficiario.Open "select * from fc_beneficiario where codigo_beneficiario = '" & Trim(Me.txtbusca1.Text) & "'", db, adOpenKeyset, adLockReadOnly
'                  If rsBeneficiario.RecordCount <> 0 Then
'                    nombenef = rsBeneficiario!denominacion_beneficiario
'                  Else
'                    nombenef = ""
'                  End If
'                  ReporteAux1_2 Trim(txtbusca1.Text), Trim(DtCOrg.Text), Trim(txtax1.Text), Trim(Txtax2.Text), Trim(txtax3.Text), nombenef, Trim(DTCNomOrg.Text)
'              End If
'           End If
'     End If
'    End If
'
'    '---si el auxiliar es 2
''       If (Me.txtax1 = "02") Or (Me.Txtax2 = "02") Or (Me.txtax3 = "02") Then
''            ReporteCtaBancaria
''        End If
''    '---si el auxiliar es 8
''        If Me.Chkaux2.Value = 1 And ((Me.Txtax2 = "08")) Then 'Or (Me.Txtax2 = "02") Or (Me.txtax3 = "02"))
''          ReporteOrg
''        End If
''        If Me.Chkaux1.Value = 1 And Chkaux2.Value = 1 Then
''          ReporteAux1_2
''        End If
''    '---auxiliar 1 y  2
'    'End If
''End If
'End Sub



Private Sub cmdBusca_Click()
Me.Fra_Busqueda.Visible = True
End Sub



Private Sub BtnBuscar_Click()
    Me.Fra_Busqueda.Visible = True
    Me.CboCampo.Text = Me.CboCampo.List(0)
    'Me.CboOperador.Text = Me.CboCampo.List(0)
End Sub

Private Sub BtnCancelar_Click()
  'Me.txtbusca1 = ""
 ' Me.Txtbusca2 = ""
'  Me.Txtbusca3 = ""
'  Me.cbocta.SetFocus
  'Me.txtaux = ""
  'Me.cboaux.Text = Me.cboaux.List(0)
End Sub

Private Sub CmdEjecutar_Click()
Select Case Me.CboCampo
    Case "codigo"
        Select Case Me.CboOperador
            Case "="
                 sql2 = "SELECT codigo_beneficiario, denominacion_beneficiario From fc_beneficiario  where  codigo_beneficiario ='" & Trim(Me.TxtValor) & "' order by codigo_beneficiario"
            Case "como"
                 sql2 = " select codigo_beneficiario, denominacion_beneficiario from  fc_beneficiario WHERE Codigo_beneficiario like '" & Trim(Me.TxtValor) & "'+'%' order by codigo_beneficiario"
        End Select
    Case "denominacion"
        Select Case Me.CboOperador
            Case "="
                sql2 = "SELECT codigo_beneficiario, denominacion_beneficiario From fc_beneficiario  where  denominacion_beneficiario ='" & Trim(Me.TxtValor) & "' order by denominacion_beneficiario"
        Case "como"
                sql2 = " select codigo_beneficiario, denominacion_beneficiario from  fc_beneficiario WHERE denominacion_beneficiario like '" & Trim(Me.TxtValor) & "'+'%'  order by denominacion_beneficiario"
    End Select
End Select
    If rsbeneficiario.State = 1 Then rsbeneficiario.Close
    rsbeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
    Set Me.DtGbenef.DataSource = rsbeneficiario
End Sub

Private Sub BtnImprimir_Click()
Call existecta(Trim(Me.cbocta), Trim(Me.cbosubcta1), Trim(Me.cbosubcta2))
If lcta = "S" Then
    If Me.cbocta.Text = "" Then
        MsgBox "Elija una cuenta", vbExclamation + vbDefaultButton1
        Me.cbocta.SetFocus
        Exit Sub
    End If
    If Me.cbosubcta1.Text = "" Then
        MsgBox "Elija una subcuenta", vbExclamation + vbDefaultButton1
        Me.cbosubcta1.SetFocus
        Exit Sub
    End If
    If Me.cbosubcta2.Text = "" Then
        MsgBox "Elija una subcuenta", vbExclamation + vbDefaultButton1
        Me.cbosubcta2.SetFocus
        Exit Sub
    End If
    If (DTPinicio.Value > DTPfin.Value) Or (DTPfin.Value < DTPinicio.Value) Then
        MsgBox "Seleccione un rango de fechas correcto", vbExclamation + vbDefaultButton1
        Exit Sub
    End If
    '----preguntar si los tres chek estan en 1
    If Chkaux1.Value = 1 And Chkaux2.Value = 1 And Chkaux3.Value = 1 Then
      If Trim(txtax1) = "01" And Trim(Txtax2) = "09" And Trim(txtax3) = "09" Then
      '---reporte de 2 organismos
      With CryConv_Conv
         .Destination = crptToWindow
         .WindowState = crptMaximized
         .WindowShowPrintSetupBtn = True
         .WindowShowSearchBtn = True
         .ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMAux1_2_3.rpt"
         .StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
         .StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
         .StoredProcParam(2) = Trim(Me.cbocta.Text)
         .StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
         .StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
         .StoredProcParam(5) = Trim(DtcCodAux1.Text) 'Trim(Me.txtbusca1)
         .StoredProcParam(6) = Trim(DtcCodAux2.Text) 'Trim(DtCOrg.Text) 'Trim(Me.Txtbusca2)
         .StoredProcParam(7) = Trim(DtcCodAux3.Text)
         .StoredProcParam(8) = Trim(Me.txtax1)
         .StoredProcParam(9) = Trim(Me.Txtax2)
         .StoredProcParam(10) = Trim(Me.txtax3)
         .Formulas(2) = "nomaux1 = '" & Trim(DtcDenomAux1.Text) & "'"    'Trim(Me.DtCOrg.Text)& Trim(Me.Txtbusca2) & "'"
         .Formulas(3) = "nomaux2 = '" & Trim(DtcDenomAux2.Text) & "'"   'Trim(Me.DtCOrg.Text)& Trim(Me.Txtbusca2) & "'"
         .Formulas(4) = "nomaux3 = '" & Trim(DtcDenomAux3.Text) & "'"    'Trim(Me.DtCOrg.Text)& Trim(Me.Txtbusca2) & "'"
         .Formulas(5) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
         .Formulas(6) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
         .Formulas(7) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
         '.Formulas(12) = "SIBs = " & SaldoIBs
         '.Formulas(13) = "SISus = " & SaldoISus
         '.Formulas(14) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
         '.Formulas(15) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
         iResult = .PrintReport
        Exit Sub
        End With
        End If
    End If
    If Chkaux1.Value = 1 Then
    '--- reportes financiadores en gral
      If Chkaux1.Value = 1 And Chkaux2.Value = 0 And Chkaux3.Value = 0 Then
        If Trim(txtax1) = "01" And Trim(Txtax2) = "09" And Trim(txtax3) = "09" Then
         With CryConv_Conv
            .Destination = crptToWindow
            .WindowState = crptMaximized
             .WindowShowPrintSetupBtn = True
             .WindowShowSearchBtn = True
             .WindowShowGroupTree = True
             .ReportFileName = App.Path & "\REPORTES\Contabilidad\Libro_Mayor_Aux\CryLibroConv_Conv.rpt"
             .StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
             .StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
             .StoredProcParam(2) = Trim(Me.cbocta.Text)
             .StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
             .StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
             .StoredProcParam(5) = Trim(Me.txtax1)
             .StoredProcParam(6) = Trim(Me.Txtax2)
             .StoredProcParam(7) = Trim(Me.txtax3)
             .StoredProcParam(8) = Trim(DtcCodAux1.Text) '(Me.txtbusca1)
             .Formulas(0) = "cta = '" & Trim(Me.cbocta.Text) & "'"
             .Formulas(1) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
             .Formulas(2) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
             .Formulas(3) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
             .Formulas(5) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
             .Formulas(6) = "nomsubcta2 = '" & Trim(Me.lbsub2) & "'"
             .Formulas(7) = "organismo = '" & Trim(DtcCodAux1.Text) & "'" '& Trim(Me.txtbusca1) & "'"
             .Formulas(11) = "SIBs = " & Val(SaldoIBs)
             .Formulas(12) = "SISus= " & Val(SaldoISus)
             .Formulas(14) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
             .Formulas(15) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
            iresesult = .PrintReport
            Exit Sub
         End With
      End If
     End If
      Select Case Trim(txtax1)
        Case "01"
            If Me.txtbusca1.Text = "" Then
              MsgBox "Escriba un beneficiario", vbExclamation + vbDefaultButton1
              Me.txtbusca1.SetFocus
              Exit Sub
            End If
        Case "02"
            If Me.cboCtaBancaria = "" Then
              MsgBox "Seleccione una Cuenta Bancaria", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
       
        Case "08"
            If DtCOrg.Text = "" Then
              MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
            
        Case "09"
            If DtCDesConvenio.Text = "" Then
              MsgBox "Seleccione un Convenio", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
      End Select
    End If
    '**************
    If Chkaux1.Value = 2 Then
      Select Case Trim(txtax1)
        Case "01"
            If Me.txtbusca1.Text = "" Then
              MsgBox "Escriba un beneficiario", vbExclamation + vbDefaultButton1
              Me.txtbusca1.SetFocus
              Exit Sub
            End If
        Case "02"
            If Me.cboCtaBancaria = "" Then
              MsgBox "Seleccione una Cuenta Bancaria", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
       
        Case "08"
            If DtCOrg.Text = "" Then
              MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
            
        Case "09"
            If DtCDesConvenio.Text = "" Then
              MsgBox "Seleccione un Convenio", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
      End Select
    End If
    '*********
    If Chkaux3.Value = 1 Then
      Select Case Trim(txtax1)
        Case "01"
            If Me.txtbusca1.Text = "" Then
              MsgBox "Escriba un beneficiario", vbExclamation + vbDefaultButton1
              Me.txtbusca1.SetFocus
              Exit Sub
            End If
        Case "02"
            If Me.cboCtaBancaria = "" Then
              MsgBox "Seleccione una Cuenta Bancaria", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
       
        Case "08"
            If DtCOrg.Text = "" Then
              MsgBox "Seleccione un Organismo Financiador", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
            
        Case "09"
            If DtCDesConvenio.Text = "" Then
              MsgBox "Seleccione un Convenio", vbExclamation + vbDefaultButton1
              Exit Sub
            End If
      End Select
    End If
'***************
  If Chkaux1.Value = 1 And Chkaux2.Value = 0 And Chkaux3.Value = 0 Then
    Select Case Trim(txtax1.Text)
        Case "01", "03"
             If Chkaux2.Value = 0 And Trim(Txtax2) = "09" Then
                reporteBeneficiario_COnvenios
             Else
                reporteBeneficiario  'procedimiento para reporte con beneficiarios
             End If
        Case "02"
             ReporteCtaBancaria
        Case "08"
             ' procedimiento para reporte con organismos
             txtbusca1 = DtCOrg.Text
             reporteBeneficiario
        Case "09"
             reporteconvenio
    End Select
  End If
  If Chkaux1.Value = 0 And Chkaux2.Value = 1 And Chkaux3.Value = 0 Then
    Select Case Trim(Txtax2)
      Case "01"
        ReporteOrg Trim(txtax1), nombenef
      Case "02"
        ReporteOrg Trim(cboCtaBancaria), nomctabancaria
      Case "08"
        ReporteOrg Trim(DtCOrg.Text), Trim(DTCNomOrg.Text)
      Case "09"
         If (cbocta = "1121" And cbosubcta1 = "02") Or (cbocta = "2116" And cbosubcta1 = "04" And cbosubcta2 <> "03") Then
          ReporteOrg Trim(DtcCodAux2.Text), Trim(DtcDenomAux2.Text)
         Else
          ReporteOrg Trim(DtCIdConvenio.Text), Trim(DtCDesConvenio.Text)
         End If
    End Select
  End If
  If Chkaux1.Value = 0 And Chkaux2.Value = 0 And Chkaux3.Value = 1 Then
  
  End If
  If Chkaux1.Value = 1 And Chkaux2.Value = 1 And Chkaux3.Value = 0 Then
      'reporte benficiario con organismos
      If Trim(txtax1.Text) = "01" And Trim(Txtax2.Text) = "08" Then
        If rsbeneficiario.State = 1 Then rsbeneficiario.Close
          rsbeneficiario.Open "select * from fc_beneficiario where codigo_beneficiario = '" & Trim(Me.txtbusca1.Text) & "'", db, adOpenKeyset, adLockReadOnly
          If rsbeneficiario.RecordCount <> 0 Then
            nombenef = rsbeneficiario!denominacion_beneficiario
          Else
            nombenef = ""
          End If
          ReporteAux1_2 Trim(txtbusca1.Text), Trim(DtCOrg.Text), Trim(txtax1.Text), Trim(Txtax2.Text), Trim(txtax3.Text), nombenef, Trim(DTCNomOrg.Text)
    End If
    'reporte benficiario con convenios
    If Trim(txtax1.Text) = "01" And Trim(Txtax2.Text) = "09" Then
        If rsbeneficiario.State = 1 Then rsbeneficiario.Close
        rsbeneficiario.Open "select * from fc_beneficiario where codigo_beneficiario = '" & Trim(Me.txtbusca1.Text) & "'", db, adOpenKeyset, adLockReadOnly
        If rsbeneficiario.RecordCount <> 0 Then
           nombenef = rsbeneficiario!denominacion_beneficiario
        Else
           nombenef = ""
        End If
        ReporteAux1_2 Trim(txtbusca1.Text), Trim(DtCIdConvenio.Text), Trim(txtax1.Text), Trim(Txtax2.Text), Trim(txtax3.Text), nombenef, Trim(DtCDesConvenio.Text)
    End If
  End If
End If
End Sub
Private Sub BtnSalir_Click()
'Dtereportes.Connection1.Close
    Unload Me
End Sub

Private Sub DataGrid1_LostFocus()
parametro = DtEreportes.rsbenef!codigo_beneficiario
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DtcCodAux1_Click(Area As Integer)
  DtcDenomAux1.Text = DtcCodAux1.BoundText
End Sub

Private Sub DtcCodAux2_Click(Area As Integer)
  DtcDenomAux2.Text = DtcCodAux2.BoundText
End Sub

Private Sub DtcCodAux3_Click(Area As Integer)
  DtcDenomAux3.Text = DtcCodAux3.BoundText
End Sub

Private Sub DtcDenomAux1_Click(Area As Integer)
  DtcCodAux1.Text = DtcDenomAux1.BoundText
End Sub

Private Sub DtcDenomAux2_Click(Area As Integer)
  DtcCodAux2.Text = DtcDenomAux2.BoundText
End Sub

Private Sub DtcDenomAux3_Click(Area As Integer)
DtcCodAux3.Text = DtcDenomAux3.BoundText
End Sub

Private Sub DtCDesConvenio_Change()
  DtCIdConvenio.BoundText = DtCDesConvenio.BoundText
End Sub


Private Sub DtCIdConvenio_Change()
  DtCDesConvenio.BoundText = DtCIdConvenio.BoundText
End Sub

Private Sub DTCNomOrg_Click(Area As Integer)
  DtCOrg.BoundText = DTCNomOrg.BoundText
End Sub

Private Sub DtcOrg_Click(Area As Integer)
  DTCNomOrg.BoundText = DtCOrg.BoundText
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
'Me.BtnGrabar.Enabled = True
On Error GoTo error_conec
    Set rsOrganismo = New ADODB.Recordset
    Set rsplanctas = New ADODB.Recordset
    Set rscuentas = New ADODB.Recordset
    Set rsnombresub1 = New ADODB.Recordset
    Set rssubcuenta = New ADODB.Recordset
    Set rscta_bancaria = New ADODB.Recordset
    Set rsbeneficiario = New ADODB.Recordset
    Set rsConvenio = New ADODB.Recordset
    '-----------
     With rsConvenio
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        sql1 = "SELECT Codigo_Convenio, Denominacion_Convenio," & _
            " org_codigo From fc_convenios"
        .Open sql1, db, adOpenKeyset, adLockReadOnly
        Set Adoconvenio.Recordset = rsConvenio
    End With
    '-----------

    If rsplanctas.State = 1 Then rsplanctas.Close
    rsplanctas.Open "SELECT Cuenta, NombreCta FROM CC_Plan_Cuentas WHERE SubCta1 = '00' AND SubCta2 = '00' order by Cuenta", db, adOpenKeyset, adLockReadOnly
    rsplanctas.MoveFirst
    Do While Not rsplanctas.EOF
        Me.cbocta.AddItem rsplanctas!cuenta
        rsplanctas.MoveNext
    Loop
    If rsbeneficiario.State = 1 Then rsbeneficiario.Close
    sql2 = "SELECT codigo_beneficiario, denominacion_beneficiario From fc_beneficiario order by denominacion_beneficiario"
    rsbeneficiario.Open sql2, db, adOpenKeyset, adLockReadOnly
    Set Me.DtGbenef.DataSource = rsbeneficiario
    
    Me.cbocta.Text = Me.cbocta.List(0)
   ' Me.DTPfin.MaxDate = CDate(Date)
   ' Me.DTPinicio.MaxDate = CDate(Date)
    Me.DTPfin.Value = Date
    Me.DTPinicio.Value = CDate("01/01/2002")
    'Me.DTPinicio.MinDate = CDate("01/01/2001")
    Me.DTPfin.MinDate = CDate(Me.DTPinicio.Value)
    Me.PRB.Visible = False
    '----------
    If rsOrganismo.State = 1 Then rsOrganismo.Close
    rsOrganismo.CursorLocation = adUseClient
    rsOrganismo.Open "SELECT Org_codigo, Org_descripcion" & _
                      " FROM fc_organismo_financiamiento order by org_Codigo", db, adOpenKeyset, adLockReadOnly
    'MsgBox rsorganismo.RecordCount
    Set AdodcOrganismo.Recordset = rsOrganismo
    'Print AdodcOrganismo.Recordset.RecordCount
    AdodcOrganismo.Refresh '

    Set DtCOrg.RowSource = AdodcOrganismo.Recordset
    DtCOrg.ListField = "org_codigo"
    DtCOrg.BoundColumn = "org_codigo" 'AdodcOrganismo.Recordset!org_codigo
    
'DtCOrg.ReFill
    
    'Set DTCNomOrg.DataSource = AdodcOrganismo.Recordset
    Set DTCNomOrg.RowSource = AdodcOrganismo.Recordset
    DTCNomOrg.ListField = "org_descripcion"
    DTCNomOrg.BoundColumn = "org_codigo"
    Me.DTCNomOrg.Visible = False
    Me.DtCOrg.Visible = False
    If Not rsOrganismo.EOF And Not rsOrganismo.BOF Then
      rsOrganismo.MoveFirst
      DtCOrg.Text = rsOrganismo!org_codigo
      DtcOrg_Click (0)
    End If
    If Not rsConvenio.EOF And Not rsConvenio.BOF Then
      rsConvenio.MoveFirst
      DtCIdConvenio.Text = rsConvenio!codigo_convenio
      DtCIdConvenio_Change
    End If
    '----------1121 y 2116
    'Dim sql1 As String
    sql31 = "SELECT codigo_beneficiario, denominacion_beneficiario From fc_beneficiario " & _
          "WHERE (tipo_beneficiario = 'O') ORDER BY denominacion_beneficiario"
    Set DtcCodAux1.RowSource = db.Execute(sql31, , commantext)
    Set DtcDenomAux1.RowSource = db.Execute(sql31, , commantext)
    sql32 = "SELECT Codigo_Convenio, Denominacion_Convenio " & _
            "From fc_convenios ORDER BY Denominacion_Convenio"
    Set DtcCodAux2.RowSource = db.Execute(sql32, , commantext)
    Set DtcDenomAux2.RowSource = db.Execute(sql32, , commantext)
    Set DtcCodAux3.RowSource = db.Execute(sql32, , commantext)
    Set DtcDenomAux3.RowSource = db.Execute(sql32, , commantext)
    
    
 '   DtcDenomAux1
 '   DtcDenomAux2
                      
  '  Exit Sub
error_conec:
    If Err.Number = -2147220992 Then
      MsgBox "ERROR EN LA CONECCION, Revise su conección a la red", vbCritical + vbDefaultButton1, "Atencion"
      End
    End If

End Sub
Public Sub Mayor000()
  Dim iResult As Integer
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
        CryLMayor.WindowState = crptMaximized
        CryLMayor.WindowShowPrintSetupBtn = True
        CryLMayor.WindowShowSearchBtn = True
        CryLMayor.ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor\CryLMayor.rpt"
        CryLMayor.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
        CryLMayor.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
        CryLMayor.StoredProcParam(2) = Trim(Me.cbocta.Text)
        CryLMayor.StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
        CryLMayor.StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
        
        CryLMayor.Formulas(0) = "cta = '" & Trim(Me.cbocta.Text) & "'"
        CryLMayor.Formulas(1) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
        CryLMayor.Formulas(2) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
        CryLMayor.Formulas(4) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
        CryLMayor.Formulas(5) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
        CryLMayor.Formulas(6) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
        CryLMayor.Formulas(9) = "SIBs = " & SaldoIBs
        CryLMayor.Formulas(10) = "SISus = " & SaldoISus
        CryLMayor.Formulas(11) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
        CryLMayor.Formulas(12) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
        iResult = CryLMayor.PrintReport
        If iResult <> 0 Then
            MsgBox CryLMayor.LastErrorNumber & " : " & CryLMayor.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
End Sub

Private Sub txtbusca1_LostFocus()
Me.BtnImprimir.Enabled = True
End Sub
Public Sub existecta(cuenta As String, subcta1 As String, subcta2 As String)
    Dim rsexiste As ADODB.Recordset
    Set rsexiste = New ADODB.Recordset
    If rsexiste.State = 1 Then rsexiste.Close
    rsexiste.CursorLocation = adUseClient
    rsexiste.Open "SELECT * From CC_Plan_Cuentas WHERE (Cuenta='" & Trim(cuenta) & "') AND (SubCta1='" & Trim(subcta1) & "') AND (SubCta2='" & Trim(subcta2) & "')", db, adOpenKeyset, adLockReadOnly
    If rsexiste.RecordCount <> 0 Then
            If rsexiste!mov = "T" Then
                MsgBox "La cuenta es de título"
                lcta = "N"
            Else
                lcta = "S"
            End If
    Else
        MsgBox "La cuenta no existe"
        cbocta.SetFocus
        lcta = "N"
    End If
End Sub
Public Sub ReporteOrg(AUX2 As String, NOMAUX2 As String)
  Dim iResult As Integer
    Set comORG = New ADODB.Command ' para obtener los saldos
    With comORG
        .CommandType = adCmdStoredProc
        .CommandText = "SaldoOrganismo"
        .Parameters.Append comORG.CreateParameter("FFInicio", adVarChar, adParamInput, 10)
        .Parameters.Append comORG.CreateParameter("FFFinal", adVarChar, adParamInput, 10)
        .Parameters.Append comORG.CreateParameter("cuenta", adVarChar, adParamInput, 5)
        .Parameters.Append comORG.CreateParameter("subcta1", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("subcta2", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("organismo", adVarChar, adParamInput, 85)
        .Parameters.Append comORG.CreateParameter("aux1", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("aux2", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("aux3", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("SIBs", adDouble, adParamOutput)
        .Parameters.Append comORG.CreateParameter("SISus", adDouble, adParamOutput)
        .Parameters("FFInicio") = Me.DTPinicio.Value
        .Parameters("FFFinal") = Me.DTPfin.Value
        .Parameters("cuenta") = Trim(Me.cbocta.Text)
        .Parameters("subcta1") = Trim(Me.cbosubcta1.Text)
        .Parameters("subcta2") = Trim(Me.cbosubcta2.Text)
        .Parameters("organismo") = AUX2  'Trim(DtcOrg.Text) 'Trim(Me.cbosubcta2.Text)
        .Parameters("aux1") = Trim(Me.txtax1.Text)
        .Parameters("aux2") = Trim(Me.Txtax2.Text)
        .Parameters("aux3") = Trim(Me.txtax3.Text)
        .ActiveConnection = db
        .Execute
        SaldoIBs = .Parameters("SIBs")
        SaldoISus = .Parameters("SISus")
    End With
       CryOrg.Destination = crptToWindow
       CryOrg.WindowState = crptMaximized
       CryOrg.WindowShowPrintSetupBtn = True
       CryOrg.WindowShowSearchBtn = True
       CryOrg.ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMAuxORG.rpt"
       CryOrg.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
       CryOrg.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
       CryOrg.StoredProcParam(2) = Trim(Me.cbocta.Text)
       CryOrg.StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
       CryOrg.StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
       CryOrg.StoredProcParam(5) = AUX2 'Trim(DtCOrg.Text) 'Trim(Me.Txtbusca2.Text)
       CryOrg.StoredProcParam(6) = Trim(Me.txtax1.Text)
       CryOrg.StoredProcParam(7) = Trim(Me.Txtax2.Text)
       CryOrg.StoredProcParam(8) = Trim(Me.txtax3.Text)
      
       CryOrg.Formulas(0) = "cta = '" & Trim(Me.cbocta.Text) & "'"
       CryOrg.Formulas(1) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
       CryOrg.Formulas(2) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
       CryOrg.Formulas(3) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
       CryOrg.Formulas(4) = "nomorg = '" & NOMAUX2 & "'" 'Trim(DTCNomOrg.Text) & "'"
       CryOrg.Formulas(5) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
       CryOrg.Formulas(6) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
       CryOrg.Formulas(7) = "organismo ='" & AUX2 & "'" '& Trim(DtCOrg.Text) & "'"
       CryOrg.Formulas(10) = "SIBs = " & SaldoIBs
       CryOrg.Formulas(11) = "SISus = " & SaldoISus
       CryOrg.Formulas(12) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
       CryOrg.Formulas(13) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
        iResult = CryOrg.PrintReport
        If iResult <> 0 Then
            MsgBox CryOrg.LastErrorNumber & " : " & CryOrg.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If

End Sub
Public Sub ReporteAux1_2(busca1 As String, busca2 As String, ax1 As String, ax2 As String, ax3 As String, nombusca1 As String, nombusca2 As String)
   If rsbeneficiario.State = 1 Then rsbeneficiario.Close
   rsbeneficiario.Open "select * from fc_beneficiario where codigo_beneficiario = '" & Trim(Me.txtbusca1.Text) & "'", db, adOpenKeyset, adLockReadOnly
   If rsbeneficiario.RecordCount <> 0 Then
      nombenef = rsbeneficiario!denominacion_beneficiario
   Else
      nombenef = ""
   End If
 Dim iResult As Integer
   Set comAux12 = New ADODB.Command
   With comAux12
       .CommandType = adCmdStoredProc
       .CommandText = "Saldos_Aux1_2"
       .Parameters.Append comAux12.CreateParameter("FFInicio", adVarChar, adParamInput, 10)
       .Parameters.Append comAux12.CreateParameter("FFFinal", adVarChar, adParamInput, 10)
       .Parameters.Append comAux12.CreateParameter("cuenta", adVarChar, adParamInput, 5)
       .Parameters.Append comAux12.CreateParameter("subcta1", adVarChar, adParamInput, 3)
       .Parameters.Append comAux12.CreateParameter("subcta2", adVarChar, adParamInput, 3)
       .Parameters.Append comAux12.CreateParameter("busca1", adVarChar, adParamInput, 15)
       .Parameters.Append comAux12.CreateParameter("busca2", adVarChar, adParamInput, 15)
       .Parameters.Append comAux12.CreateParameter("aux1", adVarChar, adParamInput, 3)
       .Parameters.Append comAux12.CreateParameter("aux2", adVarChar, adParamInput, 3)
       .Parameters.Append comAux12.CreateParameter("aux3", adVarChar, adParamInput, 3)
       .Parameters.Append comAux12.CreateParameter("SIBs", adDouble, adParamOutput)
       .Parameters.Append comAux12.CreateParameter("SISus", adDouble, adParamOutput)
       .Parameters("FFInicio") = Me.DTPinicio.Value
       .Parameters("FFFinal") = Me.DTPfin.Value
       .Parameters("cuenta") = Trim(Me.cbocta.Text)
       .Parameters("subcta1") = Trim(Me.cbosubcta1.Text)
       .Parameters("subcta2") = Trim(Me.cbosubcta2.Text)
       .Parameters("busca1") = busca1 'Trim(Me.txtbusca1.Text)
       .Parameters("busca2") = busca2 'Trim(DtCOrg.Text) 'Trim(Me.Txtbusca2.Text)
       .Parameters("aux1") = ax1 'Trim(Me.txtax1)
       .Parameters("aux2") = ax2 'Trim(Me.txtax1)"00"
       .Parameters("aux3") = ax3 'Trim(Me.txtax1)"00"
       .ActiveConnection = db
       .Execute
       SaldoIBs = .Parameters("SIBs")
       SaldoISus = .Parameters("SISus")
   End With
   
   'Me.ProgressBar1.Visible = True
   'Me.ProgressBar1.Value = 0
'   CRyAux12
       CRyAux12.Destination = crptToWindow
       CRyAux12.WindowState = crptMaximized
       CRyAux12.WindowShowPrintSetupBtn = True
       CRyAux12.WindowShowSearchBtn = True
       CRyAux12.ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMAux1_2.rpt"
       CRyAux12.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
       CRyAux12.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
       CRyAux12.StoredProcParam(2) = Trim(Me.cbocta.Text)
       CRyAux12.StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
       CRyAux12.StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
       CRyAux12.StoredProcParam(5) = busca1 'Trim(Me.txtbusca1)
       CRyAux12.StoredProcParam(6) = busca2 'Trim(DtCOrg.Text) 'Trim(Me.Txtbusca2)
       CRyAux12.StoredProcParam(7) = ax1 'Trim(Me.txtax1)
       CRyAux12.StoredProcParam(8) = ax2 'Trim(Me.Txtax2)
       CRyAux12.StoredProcParam(9) = ax3 'Trim(Me.txtax3)
       
       CRyAux12.Formulas(0) = "aux2 = '" & busca2 & "'"    'Trim(Me.DtCOrg.Text)& Trim(Me.Txtbusca2) & "'"
       CRyAux12.Formulas(1) = "benef = '" & busca1 & "'" '& Trim(Me.txtbusca1) & "'"
       CRyAux12.Formulas(2) = "cta = '" & Trim(Me.cbocta.Text) & "'"
       CRyAux12.Formulas(3) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
       CRyAux12.Formulas(4) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
       CRyAux12.Formulas(5) = "nomaux2 = '" & nombusca2 & "'" 'Trim(DTCNomOrg.Text) & "'"
       CRyAux12.Formulas(6) = "nombenef = '" & nombusca1 & "'" 'nombenef & "'"
       CRyAux12.Formulas(7) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
       CRyAux12.Formulas(8) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
       CRyAux12.Formulas(9) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
       CRyAux12.Formulas(12) = "SIBs = " & SaldoIBs
       CRyAux12.Formulas(13) = "SISus = " & SaldoISus
       CRyAux12.Formulas(14) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
       CRyAux12.Formulas(15) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
       iResult = CRyAux12.PrintReport
'*****fin aux1
'Exit Sub
End Sub
Public Sub reporteBeneficiario()
If rsbeneficiario.State = 1 Then rsbeneficiario.Close
rsbeneficiario.Open "select * from fc_beneficiario where codigo_beneficiario = '" & Trim(Me.txtbusca1.Text) & "'", db, adOpenKeyset, adLockReadOnly
If rsbeneficiario.RecordCount <> 0 Then
  nombenef = rsbeneficiario!denominacion_beneficiario
Else
  nombenef = ""
End If
            Dim iResult As Integer
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
                CryLMayorBenef.WindowState = crptMaximized
                CryLMayorBenef.WindowShowPrintSetupBtn = True
                CryLMayorBenef.WindowShowSearchBtn = True
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
                If txtax1 = "03" Then
                  NombreCaja Trim(txtbusca1)
                End If
                CryLMayorBenef.Formulas(4) = "nombenef = '" & nombenef & "'"
                CryLMayorBenef.Formulas(5) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
                CryLMayorBenef.Formulas(6) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
                CryLMayorBenef.Formulas(7) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
                CryLMayorBenef.Formulas(10) = "SIBs = " & SaldoIBs
                CryLMayorBenef.Formulas(11) = "SISus = " & SaldoISus
                CryLMayorBenef.Formulas(12) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
                CryLMayorBenef.Formulas(13) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
                iResult = CryLMayorBenef.PrintReport
End Sub
Public Sub ReporteCtaBancaria()
  Set rsctabancaria = New ADODB.Recordset
            If rsctabancaria.State = 1 Then rsctabancaria.Close
            Dim SQLVar As String
            SQLVar = "SELECT fc_bancos.Bco_descripcion_larga,fc_cuenta_bancaria.Cta_codigo," & _
                     " fc_cuenta_bancaria.Cta_descripcion_larga FROM fc_bancos INNER JOIN " & _
                     " fc_cuenta_bancaria ON  fc_bancos.Bco_codigo = fc_cuenta_bancaria.Bco_codigo " & _
                     "WHERE fc_cuenta_bancaria.Cta_codigo='" & Trim(Me.cboCtaBancaria) & "'"
            rsctabancaria.Open SQLVar, db, adOpenKeyset, adLockReadOnly
            ctabancaria = Trim(rsctabancaria!Cta_Codigo)
            nombanco = Trim(rsctabancaria!bco_descripcion_larga)
            nomctabancaria = Trim(rsctabancaria!Cta_descripcion_larga)
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
                CryLMayorCtaBancaria.WindowState = crptMaximized
                CryLMayorCtaBancaria.WindowShowPrintSetupBtn = True
                CryLMayorCtaBancaria.WindowShowSearchBtn = True
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
                CryLMayorCtaBancaria.Formulas(5) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
                CryLMayorCtaBancaria.Formulas(6) = "nomctaBancaria = '" & nomctabancaria & "'"
                CryLMayorCtaBancaria.Formulas(7) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
                CryLMayorCtaBancaria.Formulas(8) = "nomsubcta2 = '" & Trim(Me.lbsub2) & "'"
                CryLMayorCtaBancaria.Formulas(11) = "SIBs = " & Val(SaldoIBs)
                CryLMayorCtaBancaria.Formulas(12) = "SISus= " & Val(SaldoISus)
                CryLMayorCtaBancaria.Formulas(13) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
                CryLMayorCtaBancaria.Formulas(14) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
                iResult = CryLMayorCtaBancaria.PrintReport
        'End If
        If iResult <> 0 Then
           MsgBox CryLMayorBenef.LastErrorNumber & " : " & CryLMayorBenef.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
End Sub

Public Sub reporteconvenio()
'funciona para todos los otros auxiliares
Dim iResult As Integer
    Set comORG = New ADODB.Command ' para obtener los saldos
    With comORG
        .CommandType = adCmdStoredProc
        .CommandText = "SaldoConvenio"
        .Parameters.Append comORG.CreateParameter("FFInicio", adVarChar, adParamInput, 10)
        .Parameters.Append comORG.CreateParameter("FFFinal", adVarChar, adParamInput, 10)
        .Parameters.Append comORG.CreateParameter("cuenta", adVarChar, adParamInput, 5)
        .Parameters.Append comORG.CreateParameter("subcta1", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("subcta2", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("organismo", adVarChar, adParamInput, 15)
        .Parameters.Append comORG.CreateParameter("aux1", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("aux2", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("aux3", adVarChar, adParamInput, 3)
        .Parameters.Append comORG.CreateParameter("SIBs", adDouble, adParamOutput)
        .Parameters.Append comORG.CreateParameter("SISus", adDouble, adParamOutput)
        .Parameters("FFInicio") = Me.DTPinicio.Value
        .Parameters("FFFinal") = Me.DTPfin.Value
        .Parameters("cuenta") = Trim(Me.cbocta.Text)
        .Parameters("subcta1") = Trim(Me.cbosubcta1.Text)
        .Parameters("subcta2") = Trim(Me.cbosubcta2.Text)
        .Parameters("organismo") = Trim(DtCIdConvenio)  'Trim(Me.cbosubcta2.Text)
        .Parameters("aux1") = Trim(Me.txtax1.Text)
        .Parameters("aux2") = Trim(Me.Txtax2.Text)
        .Parameters("aux3") = Trim(Me.txtax3.Text)
        .ActiveConnection = db
        .Execute
        SaldoIBs = .Parameters("SIBs")
        SaldoISus = .Parameters("SISus")
    End With
       CryOrg.Destination = crptToWindow
       CryOrg.WindowState = crptMaximized
       CryOrg.WindowShowPrintSetupBtn = True
       CryOrg.WindowShowSearchBtn = True
       CryOrg.ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMConvenio.rpt"
       ''"\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMAuxORG.rpt"
       CryOrg.StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
       CryOrg.StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
       CryOrg.StoredProcParam(2) = Trim(Me.cbocta.Text)
       CryOrg.StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
       CryOrg.StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
       CryOrg.StoredProcParam(5) = Trim(DtCIdConvenio)   'Trim(DtCOrg.Text) 'Trim(Me.Txtbusca2.Text)
       CryOrg.StoredProcParam(6) = Trim(Me.txtax1.Text)
       CryOrg.StoredProcParam(7) = Trim(Me.Txtax2.Text)
       CryOrg.StoredProcParam(8) = Trim(Me.txtax3.Text)
      
       CryOrg.Formulas(0) = "cta = '" & Trim(Me.cbocta.Text) & "'"
       CryOrg.Formulas(1) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
       CryOrg.Formulas(2) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
       CryOrg.Formulas(3) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
       CryOrg.Formulas(4) = "nomorg = '" & Trim(DtCDesConvenio) & "'" ' Trim(DTCNomOrg.Text) & "'"
       CryOrg.Formulas(5) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
       CryOrg.Formulas(6) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
       CryOrg.Formulas(7) = "organismo ='" & Trim(DtCIdConvenio) & "'" 'Trim(DtcOrg.Text) & "'"
       CryOrg.Formulas(10) = "SIBs = " & SaldoIBs
       CryOrg.Formulas(11) = "SISus = " & SaldoISus
       CryOrg.Formulas(12) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
       CryOrg.Formulas(13) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
        iResult = CryOrg.PrintReport
        If iResult <> 0 Then
            MsgBox CryOrg.LastErrorNumber & " : " & CryOrg.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If

End Sub

Public Sub reporteBeneficiario_COnvenios()
  If rsbeneficiario.State = 1 Then rsbeneficiario.Close
  rsbeneficiario.Open "select * from fc_beneficiario where codigo_beneficiario = '" & Trim(Me.txtbusca1.Text) & "'", db, adOpenKeyset, adLockReadOnly
  If rsbeneficiario.RecordCount <> 0 Then
    nombenef = rsbeneficiario!denominacion_beneficiario
  Else
    nombenef = ""
  End If
  With CryBenefConvenios
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        '.ReportFileName = App.path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLMBenef_Convenios.rpt"
        .ReportFileName = App.Path & "\Reportes\Contabilidad\Libro_Mayor_Aux\CryLMBenef_Convenios.rpt"
       ''"\Reportes\Contabilidad\Libro_Mayor_Aux\CryLibroMAuxORG.rpt"
        .StoredProcParam(0) = Format(Me.DTPinicio.Value, "dd/mm/yyyy")
        .StoredProcParam(1) = Format(Me.DTPfin.Value, "dd/mm/yyyy")
        .StoredProcParam(2) = Trim(Me.cbocta.Text)
        .StoredProcParam(3) = Trim(Me.cbosubcta1.Text)
        .StoredProcParam(4) = Trim(Me.cbosubcta2.Text)
        .StoredProcParam(5) = Trim(Me.txtax1.Text)
        .StoredProcParam(6) = Trim(Me.Txtax2.Text)
        .StoredProcParam(7) = Trim(Me.txtax3.Text)
        .StoredProcParam(8) = Trim(txtbusca1.Text)
        .Formulas(0) = "benef = '" & Trim(txtbusca1) & "'"
        .Formulas(1) = "cta = '" & Trim(Me.cbocta.Text) & "'"
        .Formulas(2) = "FFechaAInicio = '" & Me.DTPinicio.Value & "'"
        .Formulas(3) = "FFechaFinal = '" & Me.DTPfin.Value & "'"
        .Formulas(4) = "nombenef = '" & Trim(nombenef) & "'"
        .Formulas(5) = "nomcta = '" & Trim(Me.lblcuenta) & "'"
        .Formulas(6) = "nomsubcta1 = '" & Trim(Me.Lblsub1) & "'"
        .Formulas(7) = "nomsubcta2 ='" & Trim(Me.lbsub2) & "'"
        .Formulas(14) = "subcta1 = '" & Trim(Me.cbosubcta1.Text) & "'"
        .Formulas(15) = "subcta2 = '" & Trim(Me.cbosubcta2.Text) & "'"
        iResult = .PrintReport
        If iResult <> 0 Then
            MsgBox .LastErrorNumber & " : " & .LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
End With
End Sub
Public Sub NombreCaja(cajas As String)
Dim rsbuscaja As ADODB.Recordset
Set rsbuscaja = New ADODB.Recordset
rsbuscaja.Open "select denominacion_caja from cc_cajas where codigo_caja='" & cajas & "'", db, adOpenKeyset, adLockReadOnly
If rsbuscaja.RecordCount <> 0 Then
  nombenef = Trim(rsbuscaja!denominacion_caja)
End If
End Sub

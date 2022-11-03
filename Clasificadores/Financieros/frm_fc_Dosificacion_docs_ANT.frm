VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_fc_Dosificacion_docs_ANT 
   BackColor       =   &H00000000&
   Caption         =   "Clasificadores - Financieros - Dosificación Facturas"
   ClientHeight    =   6855
   ClientLeft      =   1065
   ClientTop       =   2415
   ClientWidth     =   13140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      ScaleHeight     =   960
      ScaleWidth      =   12840
      TabIndex        =   33
      Top             =   120
      Width           =   12900
      Begin VB.CommandButton cmdaceptar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3480
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdAdicionar 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdBorrar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdIMPRIMIR 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton cmd_busqueda 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOSIFICACION"
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
         Left            =   8340
         TabIndex        =   35
         Top             =   300
         Width           =   2235
      End
   End
   Begin VB.PictureBox FRADATOS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5520
      Left            =   6435
      ScaleHeight     =   5460
      ScaleWidth      =   6540
      TabIndex        =   21
      Top             =   1200
      Width           =   6600
      Begin MSDataListLib.DataCombo dtc_desc1 
         DataField       =   "doc_codigo"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   915
         TabIndex        =   0
         Top             =   525
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "doc_descripcion"
         BoundColumn     =   "doc_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcfue 
         DataField       =   "beneficiario_codigo"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   4830
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox Combo4 
         DataField       =   "correl"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   5160
         TabIndex        =   6
         Text            =   "0"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox Combo3 
         DataField       =   "correl_fin"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Text            =   "100000"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox Combo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         DataField       =   "estado_codigo"
         DataSource      =   "adoLista"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   525
         Width           =   855
      End
      Begin VB.TextBox Text4 
         DataField       =   "correl_ini"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "1"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         DataField       =   "dosifica_llave"
         DataSource      =   "adoLista"
         Height          =   495
         Left            =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "-"
         Top             =   2060
         Width           =   6255
      End
      Begin VB.TextBox Text1 
         DataField       =   "dosifica_autorizacion"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   3840
         TabIndex        =   2
         Text            =   "0"
         Top             =   1320
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo dtcfu 
         DataField       =   "beneficiario_codigo"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   4830
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "DataCombo1"
      End
      Begin MSComCtl2.DTPicker TxtFecha 
         DataField       =   "dosifica_fecha"
         DataSource      =   "Ado_datos16"
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   84869121
         CurrentDate     =   41791
         MinDate         =   32874
      End
      Begin MSComCtl2.DTPicker TxtFecha1 
         DataField       =   "dosifica_fecha_ini"
         DataSource      =   "Ado_datos16"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   4005
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   84869121
         CurrentDate     =   42156
         MinDate         =   32874
      End
      Begin MSComCtl2.DTPicker TxtFecha2 
         DataField       =   "dosifica_fecha_fin"
         DataSource      =   "Ado_datos16"
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   4005
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   84869121
         CurrentDate     =   42522
         MinDate         =   32874
      End
      Begin MSComCtl2.DTPicker TxtFecha3 
         DataField       =   "dosifica_fecha_limite"
         DataSource      =   "Ado_datos16"
         Height          =   285
         Left            =   4560
         TabIndex        =   9
         Top             =   4005
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   84869121
         CurrentDate     =   42522
         MinDate         =   32874
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         DataField       =   "doc_codigo"
         DataSource      =   "adoLista"
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   525
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "doc_codigo"
         BoundColumn     =   "doc_codigo"
         Text            =   ""
      End
      Begin VB.Label lbl_tipodoc 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tipo de Documento"
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
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Limite Emisión"
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
         Index           =   10
         Left            =   4520
         TabIndex        =   37
         Top             =   3720
         Width           =   1905
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin Emisión"
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
         Index           =   7
         Left            =   2400
         TabIndex        =   36
         Top             =   3720
         Width           =   1650
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable de Registro"
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
         Index           =   11
         Left            =   120
         TabIndex        =   30
         Top             =   4540
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.de Autorización"
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
         Left            =   3840
         TabIndex        =   29
         Top             =   1020
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Solicitud"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1020
         Width           =   1665
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correlativo Final"
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
         Index           =   2
         Left            =   2520
         TabIndex        =   27
         Top             =   2835
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   5520
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C"
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
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   2835
         Width           =   135
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "LLAVE"
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
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio Emisión"
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
         Index           =   8
         Left            =   120
         TabIndex        =   23
         Top             =   3720
         Width           =   1845
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correlativo Actual"
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
         Index           =   9
         Left            =   4800
         TabIndex        =   22
         Top             =   2835
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc adoLista 
      Height          =   330
      Left            =   120
      Top             =   6375
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483624
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
   Begin MSDataGridLib.DataGrid grdlista 
      Height          =   5115
      Left            =   120
      TabIndex        =   32
      Top             =   1200
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   9022
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
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "dosifica_autorizacion"
         Caption         =   "No.autirizacion"
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
         DataField       =   "dosifica_fecha"
         Caption         =   "Fecha_solicitud"
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
         DataField       =   "dosifica_fecha_limite"
         Caption         =   "Fecha_Limite"
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
         DataField       =   "doc_codigo"
         Caption         =   "Tipo.Doc"
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
         DataField       =   "estado_codigo"
         Caption         =   "Estado."
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
         DataField       =   "correl_ini"
         Caption         =   "Correl_ini"
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
      BeginProperty Column06 
         DataField       =   "correl_fin"
         Caption         =   "Correl_Fin"
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
      BeginProperty Column07 
         DataField       =   "correl"
         Caption         =   "Correl_actual"
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
      BeginProperty Column08 
         DataField       =   "dosifica_fecha_ini"
         Caption         =   "Fecha_ini"
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
      BeginProperty Column09 
         DataField       =   "dosifica_fecha_fin"
         Caption         =   "Fecha_Fin"
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
      BeginProperty Column10 
         DataField       =   "beneficiario_codigo_resp"
         Caption         =   "Responsable"
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
      BeginProperty Column11 
         DataField       =   "fecha_registro"
         Caption         =   "fecha_registro"
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
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   -1  'True
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adofuente 
      Height          =   375
      Left            =   0
      Top             =   6600
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Caption         =   "Adofuente"
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
   Begin MSAdodcLib.Adodc Ado_docs 
      Height          =   375
      Left            =   2160
      Top             =   6600
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Caption         =   "Ado_docs"
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_fc_Dosificacion_docs_ANT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstorg As New ADODB.Recordset
Dim rstfuente As New ADODB.Recordset
Dim rs_docs As New ADODB.Recordset
Dim CAMPOS As ADODB.Field
'Dim ClBuscaGrid As CompBusquedas.ClBuscaEnGridExterno
Dim sql_financiador As String

Private Sub Adolista_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'     If pRecordset.EOF Or pRecordset.BOF Then
''      cmdEditar.Enabled = False
''      cmdBorrar.Enabled = False
'      Text1.Text = Empty
''      Text2.Text = Empty
'      Text3.Text = Empty
'      Text4.Text = Empty
'      dtcfu.Text = ""
'      dtcfue.Text = ""
'      Exit Sub
'   End If
   
'   cmdEditar.Enabled = True
'   cmdBorrar.Enabled = True
'
'   Select Case pRecordset.EditMode
'      Case adEditInProgress
'      Case adEditNone
''         Text1.Text = IIf(IsNull(pRecordset("org_codigo")), "", pRecordset("org_codigo"))
'''         Text2.Text = IIf(IsNull(pRecordset("ges_gestion")), "", pRecordset("ges_gestion"))
''         Text3.Text = IIf(IsNull(pRecordset("org_descripcion")), "", pRecordset("org_descripcion"))
''         Text4.Text = IIf(IsNull(pRecordset("org_sigla")), "", pRecordset("org_sigla"))
''         Combo3.Text = IIf(IsNull(pRecordset("beneficiario_cargo_representante")), "", pRecordset("beneficiario_cargo_representante"))
''         Combo4.Text = IIf(IsNull(pRecordset("beneficiario_codigo")), "", pRecordset("beneficiario_codigo"))
'''         Combo1.Text = IIf(IsNull(pRecordset("pais_codigo")), "", pRecordset("pais_codigo"))
''         Combo2.Text = IIf(IsNull(pRecordset("estado_codigo")), "", pRecordset("estado_codigo"))
''        'If rstfue.State = 1 Then rstfue.Close
''         dtcfu.BoundText = pRecordset("fte_codigo")
''         rstfuente.MoveFirst
''         rstfuente.Find "Fte_codigo= '" & pRecordset!fte_codigo & "'"
''         If Not rstfuente.EOF Then dtcfue.Text = rstfuente!fte_descripcion & "" Else dtcfue.Text = ""
'''         TxtFecha.Text = IIf(IsNull(pRecordset("fecha_registro")), "", pRecordset("fecha_registro"))
''         'TxtHora.Text = IIf(IsNull(pRecordset("hora_registro")), "", pRecordset("hora_registro"))
'''         Txtusuario.Text = IIf(IsNull(pRecordset("usr_codigo")), "", pRecordset("usr_codigo"))
'      Case adEditDelete
'      Case adEditAdd
'   End Select
   adoLista.Caption = CStr(adoLista.Recordset.AbsolutePosition) & " de " & CStr(adoLista.Recordset.RecordCount)
End Sub
   
Private Sub CmdAceptar_Click()
On Error GoTo errorAceptar
Dim SW As Boolean
Dim SQL_FOR As String
Dim RSTORAUX As New ADODB.Recordset

   With adoLista
            If Text1 = "" Then
                  MsgBox "INTRODUZCA DATOS"
                  Text1.SetFocus
                  Exit Sub
            End If
            If dtc_codigo1 = "" Then
               MsgBox "Debe registrar ..." + lbl_tipodoc
               dtc_codigo1.SetFocus
               Exit Sub
            End If
            If Text3 = "" Then
                MsgBox "INTRODUZCA DATOS"
                Text3.SetFocus
                Exit Sub
            End If
            If Text4 = "" Then
                MsgBox "INTRODUZCA DATOS"
                Text4.SetFocus
                Exit Sub
            End If
                                      
    Set RSTORAUX = New ADODB.Recordset
    SQL_FOR = "select * from Fc_dosificacion_docs where dosifica_autorizacion = '" & Text1.Text & "'"
    RSTORAUX.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic, adCmdText
    If RSTORAUX.RecordCount > 0 And Text1.Enabled Then
      SW = True
      MsgBox " CODIGO DUPLICADO"
      Text1.SetFocus
      Exit Sub
    End If
    '
    'db.BeginTrans
    SW = False
'    If Text1.Enabled Then
'        .Recordset("org_codigo") = Text1.Text
'    End If
            .Recordset("doc_codigo").Value = dtc_codigo1.Text         '"R-101"
            .Recordset("dosifica_autorizacion").Value = Text1.Text
            .Recordset("dosifica_fecha").Value = TxtFecha.Value
            .Recordset("beneficiario_codigo").Value = "1003579028"
            .Recordset("beneficiario_codigo_resp").Value = dtcfu.Text
            .Recordset("correl_ini").Value = Text4.Text
            .Recordset("correl_fin").Value = Combo3.Text
            .Recordset("correl").Value = Combo4.Text
            
            .Recordset("dosifica_fecha_ini").Value = TxtFecha1.Value
            .Recordset("dosifica_fecha_fin").Value = TxtFecha2.Value
            .Recordset("dosifica_fecha_limite").Value = TxtFecha3.Value
            .Recordset("dosifica_codigo_control").Value = .Recordset.RecordCount + 1
            .Recordset("dosifica_llave").Value = Text3.Text
            
            .Recordset("estado_codigo").Value = "REG"
            .Recordset("usr_codigo").Value = glusuario
            .Recordset("fecha_registro").Value = Date
            
            .Recordset.Update
            .Recordset.Requery
     '       db.CommitTrans
            
      End With
      
'   Call Cmdadicionar_Click
    
'   Call cmdCancelar_Click
   
   Exit Sub

errorAceptar:
   
   Call pErrorRst(db.Errors)
   
   adoLista.Recordset.CancelUpdate
   
   db.RollbackTrans
End Sub
 Private Sub Cmdadicionar_Click()
   Text1.Enabled = True
   adoLista.Enabled = False
   'grdlista.Enabled = False
   FRADATOS.Enabled = True
   
  ' cmdBorrar.Visible = False
   cmd_busqueda.Visible = False
   CmdIMPRIMIR.Visible = False
'   cmdSalir.Visible = False
   cmdEditar.Visible = False
  cmdAdicionar.Visible = False

   cmdaceptar.Visible = True
   cmdCancelar.Visible = True
   adoLista.Recordset.AddNew
   
'   Text1.Text = Empty
''   Text2.Text = Empty
'   Text3.Text = Empty
'   Text4.Text = Empty
'   dtcfu.Text = ""
'   dtcfue.Text = ""
'   Combo3.Text = "" 'Combo3.List(0)
'   Combo4.Text = "" 'Combo4.List(0)
''   Combo1.Text = "" 'Combo1.List(0)
'   Combo2.Text = "" 'Combo2.List(0)
''   Text2.SetFocus
'
End Sub

'Private Sub Cmdborrar_Click()
'   Dim Mensaje As String
'
'On Error GoTo errorDelete
'
'   Mensaje = "¿Borrar: " & _
'               Text1.Text & " " & _
'               Trim(Text3.Text) & "?"
'   If MsgBox(Mensaje, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar:") = vbYes Then
'      db.BeginTrans
'      adoLista.Recordset.Delete
'      db.CommitTrans
'   End If
'
'   Exit Sub
'errorDelete:
'
'   Dim e As ADODB.Error
'
'   For Each e In db.Errors
'      MsgBox "Error No. " & e.Number & " " & e.Description
'   Next
'
'   db.RollbackTrans
'
'End Sub

Private Sub Cmd_Busqueda_Click()
''BUSQUEDA.Visible = True
''fradatos.Enabled = True
' Set ClBuscaGrid = New CompBusquedas.ClBuscaEnGridExterno
'    Set ClBuscaGrid.Conexión = DB
'    ClBuscaGrid.EsTdbGrid = False
'    Set ClBuscaGrid.GridTrabajo = grdlista
'    ClBuscaGrid.QueryUtilizado = sql_financiador
'    Set ClBuscaGrid.RecordsetTrabajo = adoLista.Recordset
'    'ClBuscaGrid.CamposVisibles = "11010011"
'    ClBuscaGrid.Ejecutar

End Sub

Private Sub CmdCancelar_Click()
  On Error Resume Next
   Text1.Enabled = True
   FRADATOS.Enabled = False
    adoLista.Recordset.Requery
   ' Grdlista.ReBind
  ' cmdBorrar.Visible = True
   cmd_busqueda.Visible = True
   CmdIMPRIMIR.Visible = True
'   cmdSalir.Visible = True
   cmdEditar.Visible = True
   cmdAdicionar.Visible = True
   cmdaceptar.Visible = False
   cmdCancelar.Visible = False
   adoLista.Enabled = True
   ' Grdlista.Enabled = True
   adoLista.Recordset.Requery
   'Grdlista.ReBind
'   Unload Me
End Sub

Private Sub Cmdeditar_Click()
   If adoLista.Recordset!estado_codigo = "REG" Then
       adoLista.Enabled = False
       ' Grdlista.Enabled = False
       FRADATOS.Enabled = True
       
       'cmdBorrar.Visible = False
       cmd_busqueda.Visible = False
       CmdIMPRIMIR.Visible = False
    '   cmdSalir.Visible = False
       cmdEditar.Visible = False
      cmdAdicionar.Visible = False
    
       cmdaceptar.Visible = True
       cmdCancelar.Visible = True
    '
       Text1.Enabled = False
'       Text2.Enabled = True
       Text3.Enabled = True
       Text4.Enabled = True
       
'       Text2.SetFocus
   Else
       MsgBox "No se puede modificar un registro Aprobado ...", , "Atencion"
   End If
End Sub

Private Sub Cmdimprimir_Click()
  Dim IResult As Integer
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.ReportFileName = App.Path & "\REPORTES\clasificadores\fr_organismo_financiador.rpt"
  IResult = CrystalReport1.PrintReport
  If IResult <> 0 Then
      MsgBox CrystalReport1.LastErrorNumber & " : " & CrystalReport1.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
  CrystalReport1.WindowState = crptMaximized
  
'  Dim IResult As Integer
'    'CrystalReport1.ReportFileName = App.Path & "\clasificadores\bancos\crybancos.rpt"
'     CrystalReport1.WindowShowPrintSetupBtn = True
'     CrystalReport1.WindowShowRefreshBtn = True
'  CrystalReport1.ReportFileName = "\SAF-2000\Clasificadores\presupuesto\organismo financiador\cryorgfin.rpt"
'  IResult = CrystalReport1.PrintReport
'  If IResult <> 0 Then
'      MsgBox CrystalReport1.LastErrorNumber & " : " & CrystalReport1.LastErrorString, vbExclamation + vbOKOnly, "Error"
'  End If
'
'CrystalReport1.WindowState = crptMaximized

'REPORGFIN.Show

'   rptModalidadSeleccion.Show vbModal
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub


Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtcfu_Click(Area As Integer)
    dtcfue.BoundText = dtcfu.BoundText
End Sub

Private Sub dtcfue_Click(Area As Integer)
    dtcfu.BoundText = dtcfue.BoundText
End Sub

Private Sub Form_Load()
   
   Dim sql_fuente As String
'   Label7.Caption = frmLogin.txtUserName.Text
'   Label9.Caption = Format(Time, "HH:mm:ss")
'   Label11.Caption = Date
   
   FRADATOS.Enabled = False
   cmdBorrar.Visible = True
   cmd_busqueda.Visible = True
   CmdIMPRIMIR.Visible = True
'   cmdSalir.Visible = True
   cmdaceptar.Visible = False
   cmdCancelar.Visible = False
   
   Set rstfuente = New ADODB.Recordset
   sql_fuente = "select * from rv_unidad_vs_responsable where unidad_codigo= 'DCONT' "     'beneficiario_codigo = '4314971' or beneficiario_codigo = '4333735'
   rstfuente.Open sql_fuente, db, adOpenKeyset, adLockOptimistic, adCmdText
   rstfuente.Sort = "beneficiario_denominacion"
  ' MsgBox rstfue.RecordCount
   Set Adofuente.Recordset = rstfuente

   Set rs_docs = New ADODB.Recordset
   rs_docs.Open "select * from gc_documentos_respaldo WHERE (clasif_codigo = 'ADM') AND (doc_original = 'CLIENTE') order by doc_descripcion", db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_docs.Recordset = rs_docs
   dtc_desc1.BoundText = dtc_codigo1.BoundText

   Set rstorg = New ADODB.Recordset
   sql_financiador = "select * from fc_dosificacion_docs" 'order by org_codigo"
   rstorg.Open sql_financiador, db, adOpenKeyset, adLockOptimistic, adCmdText
   rstorg.Sort = "dosifica_fecha"
   Set adoLista.Recordset = rstorg
   'Set ClBuscaGrid = Nothing
   
  
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (rstorg.State = adStateClosed) Then rstorg.Close
   'Set rstorg = Nothing

End Sub

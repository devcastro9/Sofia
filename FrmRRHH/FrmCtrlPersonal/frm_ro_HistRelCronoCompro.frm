VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ro_HistRelCronoCompro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion Anterior - Consultoría - Visor de la relación de Compr. Presupuestarios y Cronograma de Pagos"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFiltrar 
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   7320
      TabIndex        =   16
      Top             =   360
      Width           =   855
   End
   Begin VB.ComboBox cboCodigo_grupo 
      Height          =   315
      Left            =   5760
      TabIndex        =   15
      Text            =   "cboCodigo_grupo"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtGes_gestion 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "2001"
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   1935
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   6360
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   1935
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   4320
      Width           =   3735
   End
   Begin VB.CommandButton cmdDown 
      Enabled         =   0   'False
      Height          =   555
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Permite mover el registro del postulante una posición más abajo"
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdUp 
      Enabled         =   0   'False
      Height          =   555
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Permite mover el registro del postulante una posición más arriba"
      Top             =   6120
      Width           =   735
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   6960
      Width           =   7815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   4320
      Width           =   7815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5520
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
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
         DataField       =   "ges_gestion"
         Caption         =   "Gestion"
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
         DataField       =   "codigo_unidad"
         Caption         =   "Unidad"
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
      BeginProperty Column02 
         DataField       =   "codigo_grupo"
         Caption         =   "Grupo Liq."
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
      BeginProperty Column03 
         DataField       =   "numero_pago"
         Caption         =   "No. pago"
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
      BeginProperty Column04 
         DataField       =   "codigo_beneficiario"
         Caption         =   "Codigo Benef."
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
      BeginProperty Column05 
         DataField       =   "paterno"
         Caption         =   "Paterno"
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
      BeginProperty Column06 
         DataField       =   "materno"
         Caption         =   "Materno"
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
      BeginProperty Column07 
         DataField       =   "nombres"
         Caption         =   "Nombres"
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
      BeginProperty Column08 
         DataField       =   "monto_dolares_crono"
         Caption         =   "Monto $US cronograma"
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
      BeginProperty Column09 
         DataField       =   "monto_dolares_rel"
         Caption         =   "Monto $US rel. DEV o CYD"
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
      BeginProperty Column10 
         DataField       =   "monto_bolivianos_crono"
         Caption         =   "Monto BS cronograma"
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
      BeginProperty Column11 
         DataField       =   "monto_bolivianos_rel"
         Caption         =   "Monto BS rel. DEV o CYD"
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
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1184.882
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dtcCodigo_unidad 
      DataField       =   "Uni_codigo"
      DataSource      =   "adoUnidad"
      Height          =   315
      Left            =   3480
      TabIndex        =   11
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Uni_codigo"
      Text            =   "dtcCodigo_unidad"
      Object.DataMember      =   "dbo_edListaUnidadEjecutora"
   End
   Begin VB.Label Label1 
      Caption         =   "Bs.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   4560
      TabIndex        =   32
      Top             =   8280
      Width           =   255
   End
   Begin VB.Label lblMontoPorAsignar_BS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "labMontoPorAsignar_BS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   31
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "$US:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   30
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "$US:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   29
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Bs.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   28
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Bs.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   4200
      TabIndex        =   27
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label lblMontoAsignado_BS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblMontoAsignado_BS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   26
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lblMontoAsignadoMasErrados_BS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblMontoAsignadoMasErrados_BS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "$US:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   24
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Mas errados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   23
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label lblMontoAsignadoMasErrados_US 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblMontoAsignadoMasErrados_US"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   22
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label labCerrar 
      Alignment       =   2  'Center
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   615
      Left            =   10080
      TabIndex        =   21
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblMontoPorAsignar_US 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "labMontoPorAsignar_US"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Monto Total por asignar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   19
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Label lblMontoAsignado_US 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblMontoAsignado_US"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Monto total asignado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Grupo:"
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   14
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Unidad"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   13
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Gestión"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Cronogramas de pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Compromisos de pago o Devengados Pendientes de asignación"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   6720
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Devengados Asignados al Pago"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   4815
   End
End
Attribute VB_Name = "frm_ro_HistRelCronoCompro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xGes_Gestion$
Public xCodigo_Unidad$
Public xCodigo_Grupo%
Public xNumero_Pago%
Public Xcodigo_beneficiario As String

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Call CargaListas
End Sub

Sub CargaListas()
Dim Suma_US As Double
Dim Suma_US_E As Double
Dim Suma_BS As Double
Dim Suma_BS_E As Double

If Not (Me.Adodc1.Recordset.EOF And Me.Adodc1.Recordset.BOF) Then
    'carga lista de asignados
    Suma_US = 0
    Suma_US_E = 0
    Suma_BS = 0
    Suma_BS_E = 0

    If Me.Adodc1.Recordset.EOF Then Me.Adodc1.Recordset.MoveFirst
    DE.dbo_ap_Hist_ListaAsigados Me.Adodc1.Recordset!ges_gestion, Me.Adodc1.Recordset!codigo_unidad, Me.Adodc1.Recordset!codigo_grupo, Me.Adodc1.Recordset!NUMERO_PAGO, Me.Adodc1.Recordset!codigo_beneficiario, "DEV"  ' ,   IIf(Me.Option1(0).Value = 1, "COM", "DEV")
    Me.List1.Clear
    With DE.rsdbo_ap_Hist_ListaAsigados
        If .RecordCount > 0 Then
            Do While Not .EOF
                List1.AddItem !APROBOTESORERIA & " " & Left(!ges_gestion & "    ", 4) & "-" & Left(!tipocomprobante & "   ", 3) & "-" & Left(!org_codigo & "   ", 3) & "-" & Left(Trim(CStr(!codigo_pago)) & "          ", 10) & Format(!monto_dolares, "0000000.00") & " - [" & IIf(IsNull(!CODIGO_POA), "", !CODIGO_POA) & "] " & !justificacion
                If !APROBOTESORERIA <> "E" And !APROBOTESORERIA <> "R" Then
                    Suma_US = Suma_US + !monto_dolares
                    Suma_BS = Suma_BS + !monto_bolivianos
                End If
                
                If !APROBOTESORERIA <> "R" Then
                    Suma_US_E = Suma_US_E + !monto_dolares
                    Suma_BS_E = Suma_BS_E + !monto_bolivianos
                End If
                db.Execute "update ac_ben_Comprdeven set marcado='S' where ges_gestion='" & !ges_gestion & "' and org_codigo='" & !org_codigo & "' and codigo_pago=" & !codigo_pago & " and codigo_beneficiario ='" & Me.Adodc1.Recordset!codigo_beneficiario & "'"
                .MoveNext
            Loop
        End If
        .Close
    End With
    Me.lblMontoAsignado_US.Caption = Suma_US
    Me.lblMontoAsignadoMasErrados_US.Caption = Suma_US_E
    Me.lblMontoAsignado_BS.Caption = Suma_BS
    Me.lblMontoAsignadoMasErrados_BS.Caption = Suma_BS_E
    
    'carga lista de pendientes
    Suma_US = 0
    Suma_BS = 0
    
    DE.dbo_ap_Hist_ListaPendientes Me.Adodc1.Recordset!ges_gestion, Me.Adodc1.Recordset!codigo_unidad, Me.Adodc1.Recordset!codigo_grupo, Me.Adodc1.Recordset!NUMERO_PAGO, Me.Adodc1.Recordset!codigo_beneficiario, "COM"    ' IIf(Me.Option1(0).Value = 1, "COM", "DEV")
    Me.List2.Clear
    With DE.rsdbo_ap_Hist_ListaPendientes
        If .RecordCount > 0 Then
            Do While Not .EOF
                List2.AddItem !APROBOTESORERIA & " " & Left(!ges_gestion & "    ", 4) & "-" & Left(!tipocomprobante & "   ", 3) & "-" & Left(!org_codigo & "   ", 3) & "-" & Left(Trim(CStr(!codigo_pago)) & "          ", 10) & Format(!monto_dolares, "0000000.00") & " - [" & IIf(IsNull(!CODIGO_POA), "", !CODIGO_POA) & "] " & !justificacion
                db.Execute "update ac_ben_Comprdeven set marcado='S' where ges_gestion='" & !ges_gestion & "' and org_codigo='" & !org_codigo & "' and codigo_pago=" & !codigo_pago & " and codigo_beneficiario ='" & Me.Adodc1.Recordset!codigo_beneficiario & "'"
                If !APROBOTESORERIA <> "E" And !APROBOTESORERIA <> "R" Then
                    Suma_US = Suma_US + !monto_dolares
                    Suma_BS = Suma_BS + !monto_bolivianos
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    Me.lblMontoPorAsignar_US.Caption = Suma_US
    Me.lblMontoPorAsignar_BS.Caption = Suma_BS
    
End If
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cboCodigo_grupo_Change()
    Call RefrescaListaPRincipal
End Sub

Private Sub cmdDown_Click()
''''''''''' REVISAR POR QUE YA NO DEBE FUNCIONAR
'''''Dim I%
'''''If List1.SelCount > 0 Then
'''''    For I = 0 To List1.ListCount - 1
'''''        If List1.Selected(I) Then
''''''            MsgBox List1.List(I)
'''''            DE.dbo_edHistDwn CInt(Me.Adodc1.Recordset!numero_pago), CInt(Me.Adodc1.Recordset!idfuncionario), Trim(Me.Adodc1.Recordset!GES_GESTION), Trim(Me.Adodc1.Recordset!codigo_unidad), CInt(Me.Adodc1.Recordset!codigo_grupo), Trim(Trim(Left(List1.List(I), 4))), Trim(Mid(List1.List(I), 6, 3)), Trim(Mid(List1.List(I), 10, 3)), CInt(Mid(List1.List(I), 14, 10))
'''''        End If
'''''    Next
'''''    Call RefrescaListaPRincipal
'''''    Call CargaListas
'''''End If
End Sub

Private Sub cmdFiltrar_Click()
Call RefrescaListaPRincipal
End Sub

Private Sub cmdUp_Click()
'''''' REVISAR POR QUE YA NO DEBE FUNCIONAR
'''Dim I%
'''If List2.SelCount > 0 Then
'''    For I = 0 To List2.ListCount - 1
'''        If List2.Selected(I) Then
'''            DE.dbo_edHistUp CInt(Me.Adodc1.Recordset!numero_pago), CInt(Me.Adodc1.Recordset!idfuncionario), Trim(Me.Adodc1.Recordset!GES_GESTION), Trim(Me.Adodc1.Recordset!codigo_unidad), CInt(Me.Adodc1.Recordset!codigo_grupo), Trim(Trim(Left(List2.List(I), 4))), Trim(Mid(List2.List(I), 6, 3)), Trim(Mid(List2.List(I), 10, 3)), CInt(Mid(List2.List(I), 14, 10))
'''        End If
'''    Next
'''    Call RefrescaListaPRincipal
'''    Call CargaListas
'''End If
End Sub

Private Sub dtcCodigo_unidad_Click(Area As Integer)
    Call refrescaComboPlanillas
End Sub

Private Sub Form_Load()
If glProceso = "CONSULTORIA" Then
    Me.Caption = "SAF - Consultoría - Visor de la relación de Compr. Presupuestarios y Cronograma de Pagos"
Else
    Me.Caption = "SAF - Recursos Humanos - Visor de la relación de Compr. Presupuestarios y Cronograma de Pagos"
End If

If Len(Trim(af_LiquidaMain_c.lblGestion.Caption)) > 0 Then
    Me.txtGes_gestion = af_LiquidaMain_c.lblGestion.Caption
    Me.dtcCodigo_unidad = af_LiquidaMain_c.lblCodUniSol.Caption
    Me.cboCodigo_grupo = af_LiquidaMain_c.lblCodGrupo.Caption
End If

Call RefrescaListaPRincipal
Call refrescaComboPlanillas
Me.cboCodigo_grupo = af_LiquidaMain_c.lblCodGrupo.Caption
	Call SeguridadSet(Me)
End Sub

Sub RefrescaListaPRincipal()
Me.MousePointer = vbHourglass

'Xcodigo_beneficiario = 0
If Me.Adodc1.Caption = "good" Then
    If Not (Me.Adodc1.Recordset.EOF And Me.Adodc1.Recordset.BOF) Then
        xGes_Gestion = Me.Adodc1.Recordset!ges_gestion
        xCodigo_Unidad = Me.Adodc1.Recordset!codigo_unidad
        xCodigo_Grupo = Me.Adodc1.Recordset!codigo_grupo
        xNumero_Pago = Me.Adodc1.Recordset!NUMERO_PAGO
        Xcodigo_beneficiario = Me.Adodc1.Recordset!codigo_beneficiario
    End If
End If
'MsgBox IIf(Me.Option1(0).Value = 1, "COM", "DEV")
DE.dbo_ap_Hist_ListaDetalleCrono Me.txtGes_gestion, Me.dtcCodigo_unidad, Val(Me.cboCodigo_grupo), "DEV"           ' IIf(Me.Option1(0).Value = 1, "COM", "DEV")

Set Me.Adodc1.Recordset = DE.rsdbo_ap_Hist_ListaDetalleCrono.Clone
'MsgBox Me.Adodc1.Recordset.RecordCount

DE.rsdbo_ap_Hist_ListaDetalleCrono.Close
'If Xcodigo_beneficiario <> 0 Then
    Me.Adodc1.Recordset.Find "ges_Gestion='" & xGes_Gestion & "'"
    If Len(xCodigo_Unidad) > 0 Then Me.Adodc1.Recordset.Find "codigo_unidad='" & xCodigo_Unidad & "'"
    If xCodigo_Grupo > 0 Then Me.Adodc1.Recordset.Find "codigo_grupo=" & xCodigo_Grupo
    If xNumero_Pago > 0 Then Me.Adodc1.Recordset.Find "numero_pago=" & xNumero_Pago
    If Len(Trim(Xcodigo_beneficiario)) > 0 Then Me.Adodc1.Recordset.Find "codigo_beneficiario ='" & Xcodigo_beneficiario & "'"
'End If
Me.Adodc1.Caption = "good"
Me.MousePointer = vbDefault
End Sub

Sub refrescaComboPlanillas()
Dim rs As New ADODB.Recordset
rs.Open "select codigo_grupo from ao_pagos_grupos where ges_Gestion='" & Me.txtGes_gestion & "' and codigo_unidad='" & Me.dtcCodigo_unidad & "' order by codigo_grupo", db, adOpenStatic, adLockReadOnly
Me.cboCodigo_grupo.Clear
If rs.RecordCount > 0 Then
    Do While Not rs.EOF
        Me.cboCodigo_grupo.AddItem rs!codigo_grupo
        rs.MoveNext
    Loop
End If
End Sub

Private Sub labCerrar_Click()
Unload Me
End Sub

Private Sub List1_Click()
Me.Text1 = Me.List1
End Sub

Private Sub List2_Click()
Me.Text2 = Me.List2
End Sub

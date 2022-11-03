VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmBuscaFuncionario 
   Caption         =   "Busca Funcionario"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   5205
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdElegir 
      Caption         =   "&Elegir"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid dtgFuncionario 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5106
      _Version        =   393216
      BackColor       =   14220028
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "IdFuncionario"
         Caption         =   "Id."
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
         DataField       =   "Paterno"
         Caption         =   "Paterno"
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
         DataField       =   "Materno"
         Caption         =   "Materno"
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
         DataField       =   "Nombres"
         Caption         =   "Nombres"
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
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1365.165
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmBuscaFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFuncionarios As New ADODB.Recordset

Private Sub CmdCancelar_Click()
    GlElegido = ""
    Unload Me
End Sub

Private Sub cmdElegir_Click()
    GlElegido = rsFuncionarios!IdFuncionario
    Unload Me
End Sub

Private Sub Form_Load()
    rsFuncionarios.Open "Select * From rc_Personal Where Paterno<>'" & "ACEFALIA" & "' Order By Paterno", db, adOpenStatic
    Set dtgFuncionario.DataSource = rsFuncionarios
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Resize()
    dtgFuncionario.Width = FrmBuscaFuncionario.Width - 250
    dtgFuncionario.Height = FrmBuscaFuncionario.Height - 950
    CmdCancelar.Top = dtgFuncionario.Height + 150
    cmdElegir.Top = dtgFuncionario.Height + 150
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsFuncionarios.Close
End Sub

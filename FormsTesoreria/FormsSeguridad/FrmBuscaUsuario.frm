VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmBuscaUsuario 
   Caption         =   "Busca Usuario"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   4935
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdElegir 
      Caption         =   "&Elegir"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   2400
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2535
      TabIndex        =   0
      Top             =   3000
      Width           =   2400
   End
   Begin MSDataGridLib.DataGrid dtgUsuarios 
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   13564411
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "usr_usuario"
         Caption         =   "Usuario"
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
      BeginProperty Column04 
         DataField       =   "NivelAcceso"
         Caption         =   "Nivel Acceso"
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
         DataField       =   "Usr_Activo"
         Caption         =   "Activo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Si"
            FalseValue      =   "No"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmBuscaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUsuarios As New ADODB.Recordset

Private Sub CmdCancelar_Click()
    GlElegido = ""
    Unload Me
End Sub

Private Sub cmdElegir_Click()
    GlElegido = rsUsuarios!IdFuncionario
    Unload Me
End Sub

Private Sub Form_Load()
    rsUsuarios.Open "Select ud.IdFuncionario, ud.usr_usuario, pe.Paterno, pe.Materno, pe.Nombres,ud.NivelAcceso, ud.Usr_Activo From Usuarios_Udapre ud, rc_Personal pe Where ud.IdFuncionario=pe.IdFuncionario Order by pe.Paterno", db, adOpenStatic
    Set dtgUsuarios.DataSource = rsUsuarios
End Sub

Private Sub Form_Resize()
    dtgUsuarios.Width = FrmBuscaUsuario.Width - 150
    dtgUsuarios.Height = FrmBuscaUsuario.Height - 850
    CmdCancelar.Top = dtgUsuarios.Height + 50
    cmdElegir.Top = dtgUsuarios.Height + 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsUsuarios.Close
End Sub

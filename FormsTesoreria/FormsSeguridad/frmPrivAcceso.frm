VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPrivAcceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definicion de Privilegios de Acceso"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   8085
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   6915
      TabIndex        =   3
      Top             =   2220
      Width           =   1020
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   360
      Left            =   1245
      TabIndex        =   2
      Top             =   2235
      Width           =   1020
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   360
      Left            =   135
      TabIndex        =   1
      Top             =   2235
      Width           =   1020
   End
   Begin MSDataGridLib.DataGrid dtgPrivAcceso 
      Height          =   1935
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16571633
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
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
      Caption         =   "PRIVILEGIOS DE ACCESO"
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "IdPrivAcceso"
         Caption         =   "Id"
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
         DataField       =   "DesPrivAcceso"
         Caption         =   "Descripcion"
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
         DataField       =   "btnAñadir"
         Caption         =   "Add"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "btnModificar"
         Caption         =   "Edit"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "btnEliminar"
         Caption         =   "Del"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "btnBuscar"
         Caption         =   "Find"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "btnGrabar"
         Caption         =   "Save"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "btnCancelar"
         Caption         =   "Cancel"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "btnVer"
         Caption         =   "Ver"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "btnImprimir"
         Caption         =   "Print"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "btnDetalle"
         Caption         =   "Detail"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "btnCopiarReg"
         Caption         =   "Copy"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "btnAprobar"
         Caption         =   "Pass"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   434.835
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   434.835
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column12 
            Alignment       =   2
            ColumnWidth     =   510.236
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPrivAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPrivAcceso As New ADODB.Recordset
Dim rsNivelAcceso As New ADODB.Recordset
Dim Nuevo As Boolean

Private Sub cmdEliminar_Click()
    dtgPrivAcceso.AllowDelete = True
    If MsgBox("Esta seguro de eliminar este privilegio?", vbExclamation + vbYesNo, "Atencion") = vbYes Then
      If rsNivelAcceso.State = 1 Then rsNivelAcceso.Close
      rsNivelAcceso.Open "Select * From NivelAcceso Where IdPrivAcceso='" & dtgPrivAcceso.Columns(0).Value & "'", db, adOpenStatic
      If rsNivelAcceso.RecordCount = 0 Then
         rsPrivAcceso.Delete
      Else
         MsgBox "Existen " & rsNivelAcceso.RecordCount & " Niveles de Acceso con el Acceso: '" & dtgPrivAcceso.Columns(0).Value & "'." & vbCr & _
                "Cambie el Privilegio de Acceso y luego proceda a eliminarla.", vbExclamation + vbOKOnly, "Atencion"
      End If
    End If
    dtgPrivAcceso.AllowDelete = False
End Sub

Private Sub cmdNuevo_Click()
    Nuevo = True
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dtgPrivAcceso_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
       If Len(dtgPrivAcceso.Columns(0).Value) <> 3 Then
          MsgBox "Debe introducir tres caracteres para identificar el privilegio de acceso.", vbInformation + vbOKOnly, "Atencion"
       End If
    End If
    If ColIndex = 1 Then
       If Len(dtgPrivAcceso.Columns(1).Value) = 0 Then
          MsgBox "Debe introducir la descripcion del privilegio de acceso.", vbInformation + vbOKOnly, "Atencion"
       End If
    End If
End Sub

Private Sub dtgPrivAcceso_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If ColIndex <> 0 And ColIndex <> 1 Then
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      If KeyAscii <> Asc("S") Or KeyAscii <> Asc("N") Then
         KeyAscii = Asc("N")
      End If
    End If
End Sub

Private Sub dtgPrivAcceso_BeforeInsert(Cancel As Integer)
    If Not Nuevo Then Cancel = True
End Sub

Private Sub dtgPrivAcceso_BeforeUpdate(Cancel As Integer)
    Nuevo = False
End Sub

Private Sub dtgPrivAcceso_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
    Nuevo = False
    rsPrivAcceso.Open "Select * From Privilegioacceso", db, adOpenKeyset, adLockOptimistic
    Set dtgPrivAcceso.DataSource = rsPrivAcceso
    'dtgPrivAcceso.AllowAddNew = False
    dtgPrivAcceso.AllowDelete = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsPrivAcceso.EditMode <> adEditNone Then rsPrivAcceso.CancelUpdate
    rsPrivAcceso.Close
    If rsNivelAcceso.State = 1 Then rsNivelAcceso.Close
End Sub

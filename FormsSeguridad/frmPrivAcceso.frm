VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPrivAcceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Privilegios de Operación"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   Icon            =   "frmPrivAcceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmPrivAcceso.frx":0A02
   ScaleHeight     =   4095
   ScaleWidth      =   11175
   Begin VB.CommandButton BtnCancelar 
      BackColor       =   &H8000000A&
      Caption         =   "Cancelar"
      Height          =   675
      Left            =   120
      Picture         =   "frmPrivAcceso.frx":6CA34
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2640
      Width           =   765
   End
   Begin VB.CommandButton BtnGrabar 
      BackColor       =   &H8000000A&
      Caption         =   "Grabar"
      Height          =   675
      Left            =   120
      Picture         =   "frmPrivAcceso.frx":6CC3E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1980
      Width           =   765
   End
   Begin VB.CommandButton BtnSalir 
      BackColor       =   &H8000000A&
      Caption         =   "Cerrar"
      Height          =   675
      Left            =   120
      Picture         =   "frmPrivAcceso.frx":6CE48
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3300
      Width           =   765
   End
   Begin VB.CommandButton BtnEliminar 
      BackColor       =   &H8000000A&
      Caption         =   "Anular"
      Height          =   675
      Left            =   120
      Picture         =   "frmPrivAcceso.frx":6D052
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Anula Registro Activo"
      Top             =   1320
      Width           =   765
   End
   Begin VB.CommandButton BtnModificar 
      BackColor       =   &H8000000A&
      Caption         =   "Modificar"
      Height          =   675
      Left            =   120
      Picture         =   "frmPrivAcceso.frx":6DD1C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Modifica Registro Activo"
      Top             =   660
      Width           =   765
   End
   Begin VB.CommandButton BtnAñadir 
      BackColor       =   &H8000000A&
      Caption         =   "Nuevo"
      Height          =   675
      Left            =   120
      Picture         =   "frmPrivAcceso.frx":6E2FC
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Nuevo Registro"
      Top             =   0
      Width           =   765
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1035
      TabIndex        =   1
      Top             =   2520
      Width           =   10080
      Begin VB.CheckBox chkAprobar 
         Caption         =   "Aprobar"
         Height          =   240
         Left            =   9030
         TabIndex        =   16
         Top             =   315
         Width           =   855
      End
      Begin VB.CheckBox chkCopiar 
         Caption         =   "Copiar"
         Height          =   210
         Left            =   8010
         TabIndex        =   15
         Top             =   330
         Width           =   765
      End
      Begin VB.CheckBox chkIrDetalle 
         Caption         =   "Ir Detalle"
         Height          =   240
         Left            =   8025
         TabIndex        =   14
         Top             =   750
         Width           =   930
      End
      Begin VB.CheckBox chkImprimir 
         Caption         =   "Imprimir"
         Height          =   195
         Left            =   7050
         TabIndex        =   13
         Top             =   765
         Width           =   825
      End
      Begin VB.CheckBox chkVer 
         Caption         =   "Ver"
         Height          =   195
         Left            =   6060
         TabIndex        =   12
         Top             =   765
         Width           =   585
      End
      Begin VB.CheckBox chkCancelar 
         Caption         =   "Cancelar"
         Height          =   195
         Left            =   4995
         TabIndex        =   11
         Top             =   765
         Width           =   990
      End
      Begin VB.CheckBox chkGrabar 
         Caption         =   "Grabar"
         Height          =   195
         Left            =   4005
         TabIndex        =   10
         Top             =   765
         Width           =   855
      End
      Begin VB.CheckBox chkBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Left            =   7035
         TabIndex        =   9
         Top             =   315
         Width           =   810
      End
      Begin VB.CheckBox chkEliminar 
         Caption         =   "Eliminar"
         Height          =   240
         Left            =   6060
         TabIndex        =   8
         Top             =   315
         Width           =   840
      End
      Begin VB.CheckBox chkModificar 
         Caption         =   "Modificar"
         Height          =   270
         Left            =   5010
         TabIndex        =   7
         Top             =   285
         Width           =   945
      End
      Begin VB.CheckBox chkAdicionar 
         Caption         =   "Adicionar"
         Height          =   255
         Left            =   4005
         TabIndex        =   6
         Top             =   300
         Width           =   1020
      End
      Begin VB.TextBox txtDesPrivAcceso 
         Height          =   285
         Left            =   1305
         MaxLength       =   15
         TabIndex        =   5
         Top             =   720
         Width           =   2505
      End
      Begin VB.TextBox txtIdPrivAcceso 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1290
         MaxLength       =   3
         TabIndex        =   4
         Top             =   285
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   255
         TabIndex        =   3
         Top             =   735
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Id. Privilegio:"
         Height          =   195
         Left            =   255
         TabIndex        =   2
         Top             =   315
         Width           =   900
      End
   End
   Begin MSDataGridLib.DataGrid dtgPrivAcceso 
      Height          =   2415
      Left            =   1035
      TabIndex        =   0
      Top             =   105
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648384
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
      Caption         =   "PRIVILEGIOS DE OPERACION"
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
         Caption         =   "Descripción"
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
         Caption         =   "Modif"
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
         Caption         =   "Elim"
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
         Caption         =   "Busca"
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
         Caption         =   "Graba"
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
         Caption         =   "Imprim"
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
         Caption         =   "IrDetalle"
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
         Caption         =   "Copia"
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
         Caption         =   "Aprob"
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
            Locked          =   -1  'True
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column12 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   659.906
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
Dim rsAuxPrivAcceso As New ADODB.Recordset
Dim rsNivelAcceso As New ADODB.Recordset
Dim Nuevo As Boolean
Dim vIdPrivAcceso As String

Private Sub BtnCancelar_Click()
    Nuevo = False
    Me.Height = 2970
    If rsPrivAcceso.RecordCount > 0 Then
        BotonesNavegar
    Else
        BotonesInicio
    End If
    txtIdPrivAcceso.Enabled = True
    dtgPrivAcceso.AllowAddNew = False
End Sub

Private Sub BtnModificar_Click()
    Nuevo = False
    Me.Height = 4080
    txtIdPrivAcceso.Enabled = False
    RecuperaPrivilegio
    BotonesConfirma
End Sub

Private Sub BtnEliminar_Click()
    If rsPrivAcceso.RecordCount > 0 Then
        dtgPrivAcceso.AllowDelete = True
        If MsgBox("Esta seguro de eliminar este privilegio?", vbCritical + vbYesNo, "Atención") = vbYes Then
          If rsNivelAcceso.State = 1 Then rsNivelAcceso.Close
          rsNivelAcceso.Open "Select * Fromgc_nivelacceso Where IdPrivAcceso='" & dtgPrivAcceso.Columns(0).Value & "'", db, adOpenStatic
          If rsNivelAcceso.RecordCount = 0 Then
             rsPrivAcceso.Delete
          Else
             MsgBox "Existen " & rsNivelAcceso.RecordCount & " Niveles de Acceso con el Acceso: '" & dtgPrivAcceso.Columns(0).Value & "'." & vbCr & _
                    "Cambie el Privilegio de Acceso y luego proceda a eliminarla.", vbExclamation + vbOKOnly, "Atención"
          End If
        End If
        dtgPrivAcceso.AllowDelete = False
    End If
    If rsPrivAcceso.RecordCount > 0 Then
        BotonesNavegar
    Else
        BotonesInicio
    End If
End Sub

Private Sub BtnGrabar_Click()
On Error GoTo Error
    If ValidaCampos Then
        Screen.MousePointer = vbHourglass
        If Nuevo Then
            rsAuxPrivAcceso.Open "Select * From PrivilegioAcceso", db, adOpenKeyset, adLockOptimistic
            rsAuxPrivAcceso.AddNew
            rsAuxPrivAcceso!IdPrivAcceso = txtIdPrivAcceso
        Else
            rsAuxPrivAcceso.Open "Select * From PrivilegioAcceso Where IdPrivAcceso='" & vIdPrivAcceso & "'", db, adOpenKeyset, adLockOptimistic
        End If
        'rsAuxPrivAcceso!IdPrivAcceso = txtIdPrivAcceso
        rsAuxPrivAcceso!DesPrivAcceso = txtDesPrivAcceso
        rsAuxPrivAcceso!BtnAñadir = IIf(chkAdicionar = 1, True, False)
        rsAuxPrivAcceso!BtnModificar = IIf(chkModificar = 1, True, False)
        rsAuxPrivAcceso!BtnEliminar = IIf(chkEliminar = 1, True, False)
        rsAuxPrivAcceso!BtnBuscar = IIf(chkBuscar = 1, True, False)
        rsAuxPrivAcceso!BtnGrabar = IIf(chkGrabar = 1, True, False)
        rsAuxPrivAcceso!BtnCancelar = IIf(chkCancelar = 1, True, False)
        rsAuxPrivAcceso!BtnVer = IIf(chkVer = 1, True, False)
        rsAuxPrivAcceso!BtnImprimir = IIf(chkImprimir = 1, True, False)
        rsAuxPrivAcceso!BtnDetalle = IIf(chkIrDetalle = 1, True, False)
        rsAuxPrivAcceso!BtnCopiarReg = IIf(chkCopiar = 1, True, False)
        rsAuxPrivAcceso!BtnAprobar = IIf(chkAprobar = 1, True, False)
        rsAuxPrivAcceso.Update
        rsAuxPrivAcceso.Close
        
        rsPrivAcceso.Close
        rsPrivAcceso.Open "Select * From PrivilegioAcceso", db, adOpenKeyset, adLockOptimistic
        Set dtgPrivAcceso.DataSource = rsPrivAcceso
        Screen.MousePointer = vbDefault
        Me.Height = 2970
        txtIdPrivAcceso.Enabled = True
        BotonesNavegar
    End If
    Exit Sub
Error:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & ", " & Err.Description
    rsAuxPrivAcceso.CancelUpdate
    If rsAuxPrivAcceso.State = 1 Then rsAuxPrivAcceso.Close
    BotonesConfirma
End Sub

Private Sub BtnAñadir_Click()
    Nuevo = True
    Me.Height = 4080
    VaciaCampos
    BotonesConfirma
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Function ValidaCampos() As Boolean
    ValidaCampos = True
    If Len(txtIdPrivAcceso) = 0 Or Len(txtIdPrivAcceso) < 3 Then
       MsgBox "Debe introducir tres caracteres para identificar el privilegio de operación", vbInformation + vbOKOnly, "Atención"
       txtIdPrivAcceso.SetFocus
       ValidaCampos = False
       Exit Function
    End If
    If Len(txtDesPrivAcceso) = 0 Then
       MsgBox "Debe introducir la descripción del privilegio de acceso", vbInformation + vbOKOnly, "Atención"
       txtDesPrivAcceso.SetFocus
       ValidaCampos = False
       Exit Function
    End If
End Function

Private Sub Form_Load()
    Nuevo = False
    Me.Height = 2970
    rsPrivAcceso.Open "Select * From PrivilegioAcceso", db, adOpenKeyset, adLockOptimistic
    Set dtgPrivAcceso.DataSource = rsPrivAcceso
    If rsPrivAcceso.RecordCount = 0 Then
        BotonesInicio
    Else
        BotonesNavegar
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If rsPrivAcceso.EditMode <> adEditNone Then rsPrivAcceso.CancelUpdate
    rsPrivAcceso.Close
    If rsNivelAcceso.State = 1 Then rsNivelAcceso.Close
End Sub

Private Sub VaciaCampos()
    txtIdPrivAcceso = ""
    txtDesPrivAcceso = ""
    chkAdicionar = 0
    chkModificar = 0
    chkEliminar = 0
    chkBuscar = 0
    chkGrabar = 0
    chkCancelar = 0
    chkVer = 0
    chkImprimir = 0
    chkIrDetalle = 0
    chkCopiar = 0
    chkAprobar = 0
End Sub

Private Sub txtDesPrivAcceso_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Not (KeyAscii > 64 And KeyAscii < 91 Or KeyAscii = 32 Or KeyAscii = 8) Then
       Beep
       KeyAscii = 0
    End If
End Sub

Private Sub txtIdPrivAcceso_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Not (KeyAscii > 64 And KeyAscii < 91 Or KeyAscii = 32 Or KeyAscii = 8) Then
       Beep
       KeyAscii = 0
    End If
End Sub

Private Sub RecuperaPrivilegio()
    vIdPrivAcceso = rsPrivAcceso!IdPrivAcceso
    txtIdPrivAcceso = rsPrivAcceso!IdPrivAcceso
    txtDesPrivAcceso = rsPrivAcceso!DesPrivAcceso
    chkAdicionar = IIf(rsPrivAcceso!BtnAñadir, 1, 0)
    chkModificar = IIf(rsPrivAcceso!BtnModificar, 1, 0)
    chkEliminar = IIf(rsPrivAcceso!BtnEliminar, 1, 0)
    chkBuscar = IIf(rsPrivAcceso!BtnBuscar, 1, 0)
    chkGrabar = IIf(rsPrivAcceso!BtnGrabar, 1, 0)
    chkCancelar = IIf(rsPrivAcceso!BtnCancelar, 1, 0)
    chkVer = IIf(rsPrivAcceso!BtnVer, 1, 0)
    chkImprimir = IIf(rsPrivAcceso!BtnImprimir, 1, 0)
    chkIrDetalle = IIf(rsPrivAcceso!BtnDetalle, 1, 0)
    chkCopiar = IIf(rsPrivAcceso!BtnCopiarReg, 1, 0)
    chkAprobar = IIf(rsPrivAcceso!BtnAprobar, 1, 0)
End Sub

Private Sub BotonesConfirma()
On Error Resume Next
    BtnAñadir.Enabled = False
    BtnModificar.Enabled = False
    BtnGrabar.Enabled = True
    BtnCancelar.Enabled = True
    BtnEliminar.Enabled = False
    BtnSalir.Enabled = False
End Sub

Private Sub BotonesNavegar()
On Error Resume Next
    BtnAñadir.Enabled = True
    BtnModificar.Enabled = True
    BtnGrabar.Enabled = False
    BtnCancelar.Enabled = False
    BtnEliminar.Enabled = True
    BtnSalir.Enabled = True
End Sub

Private Sub BotonesInicio()
On Error Resume Next
    BtnAñadir.Enabled = True
    BtnModificar.Enabled = False
    BtnGrabar.Enabled = False
    BtnCancelar.Enabled = False
    BtnEliminar.Enabled = False
    BtnSalir.Enabled = True
End Sub

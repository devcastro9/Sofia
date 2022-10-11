VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form AlmCLFrmProveedor 
   Caption         =   "Proveedores"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11685
      TabIndex        =   9
      Top             =   3735
      Width           =   11685
      Begin VB.Frame Frame3 
         Height          =   60
         Left            =   1275
         TabIndex        =   10
         Top             =   255
         Width           =   8310
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clasificador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   1
         Left            =   9660
         TabIndex        =   11
         Top             =   75
         Width           =   1845
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clasificador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   9675
         TabIndex        =   12
         Top             =   90
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   990
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   11625
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   11685
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "UNIDAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   60
         TabIndex        =   8
         Top             =   660
         Width           =   885
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   195
         Left            =   1050
         TabIndex        =   7
         Top             =   660
         Width           =   2310
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "USUARIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8670
         TabIndex        =   6
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   9780
         TabIndex        =   5
         Top             =   660
         Width           =   540
      End
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "CLASIFICADOR DE PROVEEDORES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   3045
         TabIndex        =   4
         Top             =   135
         Width           =   5175
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         Caption         =   "."
         ForeColor       =   &H0000C000&
         Height          =   180
         Left            =   3915
         TabIndex        =   3
         Top             =   675
         Width           =   2655
      End
   End
   Begin VB.PictureBox PicBoton 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   0
      ScaleHeight     =   2745
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   990
      Width           =   1215
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   855
         Left            =   75
         Picture         =   "AlmCLFrmProveedor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1095
         Width           =   945
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar"
         Height          =   855
         Left            =   75
         Picture         =   "AlmCLFrmProveedor.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   945
      End
   End
   Begin TrueOleDBGrid60.TDBGrid tdbgProv 
      Align           =   3  'Align Left
      Height          =   2745
      Left            =   1215
      OleObjectBlob   =   "AlmCLFrmProveedor.frx":0AAC
      TabIndex        =   1
      Top             =   990
      Width           =   10425
   End
End
Attribute VB_Name = "AlmCLFrmProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsProv As ADODB.Recordset

Private Sub CmdEliminar_Click()
    If RsProv.RecordCount > 0 Then
        If RsProv.BOF Or RsProv.EOF Then Exit Sub
        If MsgBox("Esta seguro de eliminar a el Proveedor seleccionado.", vbQuestion + vbYesNo, "Atención") = vbYes Then
            RsProv.Delete
            RsProv.Requery
        End If
    End If
End Sub

Private Sub CmdSalir_Click()
    If Not (RsProv.BOF And RsProv.EOF) Then
        If RsProv.EditMode <> adEditNone Then RsProv.Update
    End If
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 3315
    Me.Width = 11805
    '
    GlSqlAux = "SELECT * FROM ALProveedores ORDER BY CodProveedor"
    Set RsProv = New ADODB.Recordset
    RsProv.Open GlSqlAux, db, adOpenStatic, adLockOptimistic
    tdbgProv.DataSource = RsProv
End Sub
Private Sub Form_Resize()
On Error Resume Next
    tdbgProv.Width = Me.ScaleWidth - picBoton.Width
End Sub


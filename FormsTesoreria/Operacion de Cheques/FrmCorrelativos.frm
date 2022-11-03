VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmCorrelativos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información de los Correlativos de Cheques y Transfrencias"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10995
   Icon            =   "FrmCorrelativos.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CryCorrel 
      Left            =   2100
      Top             =   7410
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   10935
      TabIndex        =   0
      Top             =   0
      Width           =   10995
      Begin VB.Label Label2 
         Caption         =   "UNIDAD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   60
         TabIndex        =   5
         Top             =   675
         Width           =   1110
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   4
         Top             =   705
         Width           =   2460
      End
      Begin VB.Label Label6 
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   9210
         TabIndex        =   3
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   2
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "CORRELATIVOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   360
         Left            =   4890
         TabIndex        =   1
         Top             =   135
         Width           =   2565
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8160
      Left            =   1395
      TabIndex        =   8
      Top             =   1035
      Width           =   9165
      Begin VB.Frame FraCambia 
         Height          =   1920
         Left            =   945
         TabIndex        =   12
         Top             =   3195
         Visible         =   0   'False
         Width           =   5955
         Begin VB.CommandButton CmdSale 
            Caption         =   "Salir"
            Height          =   570
            Left            =   4455
            TabIndex        =   21
            Top             =   1305
            Width           =   1335
         End
         Begin VB.CommandButton CmdGrabar 
            Caption         =   "Grabar"
            Height          =   570
            Left            =   4455
            TabIndex        =   20
            Top             =   735
            Width           =   1335
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Buscar"
            Height          =   525
            Left            =   4455
            TabIndex        =   19
            Top             =   210
            Width           =   1335
         End
         Begin VB.TextBox TxtEstadoAprobacion 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   495
            TabIndex        =   17
            Top             =   1290
            Width           =   1830
         End
         Begin VB.TextBox TxTOrganismo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2535
            TabIndex        =   14
            Top             =   570
            Width           =   1485
         End
         Begin VB.TextBox TxtCodigoPago 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   465
            TabIndex        =   13
            Top             =   570
            Width           =   1845
         End
         Begin VB.Label Label7 
            Caption         =   "Estado de Aprobación"
            Height          =   210
            Left            =   510
            TabIndex        =   18
            Top             =   1020
            Width           =   1950
         End
         Begin VB.Label Label5 
            Caption         =   "Organismo"
            Height          =   225
            Left            =   2535
            TabIndex        =   16
            Top             =   255
            Width           =   1170
         End
         Begin VB.Label Label4 
            Caption         =   "Cmpte."
            Height          =   225
            Left            =   465
            TabIndex        =   15
            Top             =   270
            Width           =   1515
         End
      End
      Begin MSDataGridLib.DataGrid DtGCorrelativos 
         Height          =   7680
         Left            =   240
         TabIndex        =   9
         Top             =   255
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   13547
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Enabled         =   -1  'True
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
   End
   Begin VB.Frame FraOpciones 
      Height          =   8175
      Left            =   90
      TabIndex        =   6
      Top             =   990
      Width           =   1245
      Begin VB.CommandButton CmdHabilitar 
         Caption         =   "Habilitar"
         Height          =   825
         Left            =   150
         TabIndex        =   11
         Top             =   1245
         Width           =   945
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   885
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2070
         Width           =   945
      End
      Begin VB.CommandButton CmdImpresion 
         Caption         =   "Imprime"
         Height          =   885
         Left            =   150
         Picture         =   "FrmCorrelativos.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   945
      End
   End
End
Attribute VB_Name = "FrmCorrelativos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCorrelativos As New ADODB.Recordset
Dim rsbusca  As New ADODB.Recordset

Private Sub CmdBuscar_Click()
Set rsbusca = New ADODB.Recordset
    rsbusca.Open "select * from pago_detalle where codigo_pago=" & (Val(TxtCodigoPago.Text)) & " and org_codigo='" & txtorganismo.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsbusca.RecordCount > 0 Then
        TxtEstadoAprobacion.Text = rsbusca("Estado_Aprobacion")
    Else
       MsgBox "No lo encontró"
    End If
End Sub

Private Sub CmdGrabar_Click()
Resp = MsgBox("Esta seguro de colocar el estado de aprobacion en S", vbYesNo)
If Resp = vbYes Then
    db.Execute "update pago_detalle set estado_aprobacion='N' where codigo_pago=" & (Val(TxtCodigoPago.Text)) & " and org_codigo='" & txtorganismo.Text & "'"
End If
    FraCambia.Visible = False
    DtGCorrelativos.Enabled = True
End Sub

Private Sub CmdHabilitar_Click()
Dim Resp As String
    Resp = InputBox("Introducir Clave")
    If Resp = "CELIA" Then
        DtGCorrelativos.Enabled = True
        DtGCorrelativos.AllowUpdate = True
        FraCambia.Visible = True
    End If
End Sub

Private Sub CmdImpresion_Click()
            CryCorrel.ReportFileName = App.Path & "\FormsTesoreria\Operacion de Cheques\Rpt_Correlativos.rpt"
            iResult = CryCorrel.PrintReport
            If iResult <> 0 Then
                MsgBox CryCorrel.LastErrorNumber & " : " & CryCorrel.LastErrorString, vbCritical + vbOKOnly, "Error..."
            End If
                        
End Sub

Private Sub CmdSale_Click()
    FraCambia.Visible = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Set rsCorrelativos = New ADODB.Recordset
    rsCorrelativos.Open "select cta_codigo1 as [Cta. Codigo],cta_codigo2 as TGN,numero_correlativo as Correlativo,descripcion from fc_correl", db, adOpenKeyset, adLockOptimistic
    If rsCorrelativos.RecordCount > 0 Then
        Set DtGCorrelativos.DataSource = rsCorrelativos
    End If
	Call SeguridadSet(Me)
End Sub


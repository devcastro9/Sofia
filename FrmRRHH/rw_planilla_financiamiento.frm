VERSION 5.00
Begin VB.Form rw_planilla_financiamiento 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnLimpiar 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton btnGenerar 
      Caption         =   "Generar "
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame frmDatos 
      Caption         =   "Parametros Reporte"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Top             =   1920
         Width           =   2295
      End
      Begin VB.ComboBox cmb_financiamiento 
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Top             =   1440
         Width           =   2295
      End
      Begin VB.ComboBox cmb_mes 
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cmb_gestion 
         Height          =   315
         ItemData        =   "rw_planilla_financiamiento.frx":0000
         Left            =   2280
         List            =   "rw_planilla_financiamiento.frx":002E
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblplanilla 
         Caption         =   "Planilla codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblfinanciamiento 
         Caption         =   "Financiamiento"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblmes 
         Caption         =   "Mes"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblgestion 
         Caption         =   "Gestion"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "rw_planilla_financiamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CargarControles()
    Dim rsMes As New ADODB.Recordset
    Dim rsFinan As New ADODB.Recordset
    rsMes.Open " SELECT DISTINCT mes_grupo, CASE mes_grupo WHEN 1 THEN 'Enero' WHEN 2 THEN 'Febrero' WHEN 3 THEN 'Marzo' WHEN 4 THEN 'Abril' WHEN 5 THEN 'Mayo' WHEN 6 THEN 'Junio' WHEN 7 THEN 'Julio' WHEN 8 THEN 'Agosto' WHEN 9 THEN 'Septiembre' WHEN 10 THEN 'Octubre' WHEN 11 THEN 'Noviembre' WHEN 12 THEN 'Diciembre' ELSE 'Enero' END As mes FROM ro_pagos_cronograma_Detalle ", db, adOpenStatic
    rsMes.MoveFirst
    With Me.cmb_mes
        .Clear
        Do
            .AddItem rsMes![mes]
            rsMes.MoveNext
        Loop Until rsMes.EOF
    End With
    
    rsFinan.Open " SELECT DISTINCT org_codigo, org_descripcion FROM fc_organismo_financiamiento ", db, adOpenStatic
    rsFinan.MoveFirst
    With Me.cmb_financiamiento
        .Clear
        Do
            .AddItem rsFinan![org_descripcion]
            rsFinan.MoveNext
        Loop Until rsFinan.EOF
    End With
'UserForm_Initialize_Exit:
    On Error Resume Next
    rsMes.Close
End Sub

Private Sub Form_Load()
  Call CargarControles
	Call SeguridadSet(Me)
End Sub

VERSION 5.00
Begin VB.Form mw_opcion_importacion 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form2"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   2010
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtnFL 
      Caption         =   "FACTURACION LOCAL"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
   End
   Begin VB.CommandButton BtnID 
      Caption         =   "IMPORTACION DIRECTA (Cliente)"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "mw_opcion_importacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VAR_TIPOC2 As String
Dim VAR_TIT As String

Private Sub BtnFL_Click()
    VAR_TIPOC2 = "V"
    Call ABRIR_OPCION
End Sub

Private Sub BtnID_Click()
    VAR_TIPOC2 = "L"
    Call ABRIR_OPCION
End Sub

Private Sub ABRIR_OPCION()
    Select Case Glaux
        Case "PROVI"
            'Glaux = "PROVI"
            VAR_TIPOC = VAR_TIPOC2
            fw_compras_gral.lbl_titulo = frmMain.Mnu_ProveedoresEquipos.Caption
            fw_compras_gral.FraNavega = frmMain.Mnu_ProveedoresEquipos.Caption
            fw_compras_gral.lbl_titulo2 = frmMain.Mnu_ProveedoresEquipos.Caption
            fw_compras_gral.Show
        Case "TRANS"
            'Glaux = "TRANS"
            VAR_TIPOC = VAR_TIPOC2
            fw_compras_gral.lbl_titulo = frmMain.Mnu_Transporte.Caption
            fw_compras_gral.FraNavega = frmMain.Mnu_Transporte.Caption
            fw_compras_gral.lbl_titulo2 = frmMain.Mnu_Transporte.Caption
            fw_compras_gral.Show
        Case "ADUAN"
            'Glaux = "ADUAN"
            VAR_TIPOC = VAR_TIPOC2
            fw_compras_gral.lbl_titulo = frmMain.Mnu_Nacionalizacion.Caption
            fw_compras_gral.FraNavega = frmMain.Mnu_Nacionalizacion.Caption
            fw_compras_gral.lbl_titulo2 = frmMain.Mnu_Nacionalizacion.Caption
            fw_compras_gral.Show
        Case "DESCA"
            'Glaux = "DESCA"
            VAR_TIPOC = VAR_TIPOC2
            fw_compras_gral.lbl_titulo = frmMain.Mnu_Descarguio.Caption
            fw_compras_gral.FraNavega = frmMain.Mnu_Descarguio.Caption
            fw_compras_gral.lbl_titulo2 = frmMain.Mnu_Descarguio.Caption
            fw_compras_gral.Show
        Case "CONTR"
            'Glaux = "CONTR"
            VAR_TIPOC = VAR_TIPOC2
            fw_compras_gral.lbl_titulo = frmMain.Mnu_ContratacionTecnicos.Caption
            fw_compras_gral.FraNavega = frmMain.Mnu_ContratacionTecnicos.Caption
            fw_compras_gral.lbl_titulo2 = frmMain.Mnu_ContratacionTecnicos.Caption
            fw_compras_gral.Show
        Case Else
            'Glaux = "PROVI"
            VAR_TIPOC = VAR_TIPOC2
            fw_compras_gral.lbl_titulo = frmMain.Mnu_ProveedoresEquipos.Caption
            fw_compras_gral.FraNavega = frmMain.Mnu_ProveedoresEquipos.Caption
            fw_compras_gral.lbl_titulo2 = frmMain.Mnu_ProveedoresEquipos.Caption
            fw_compras_gral.Show
    End Select
    Unload Me
End Sub

Private Sub Form_Load()
    Aux = "COMEX"
    Select Case Glaux
        Case "PROVI"
            VAR_TIT = frmMain.Mnu_ProveedoresEquipos.Caption
        Case "TRANS"
            VAR_TIT = frmMain.Mnu_Transporte.Caption
        Case "ADUAN"
            VAR_TIT = frmMain.Mnu_Nacionalizacion.Caption
        Case "DESCA"
            VAR_TIT = frmMain.Mnu_Descarguio.Caption
        Case "CONTR"
            VAR_TIT = frmMain.Mnu_ContratacionTecnicos.Caption
        Case Else
            VAR_TIT = frmMain.Mnu_ProveedoresEquipos.Caption
    End Select
    mw_opcion_importacion.Caption = VAR_TIT
	Call SeguridadSet(Me)
End Sub

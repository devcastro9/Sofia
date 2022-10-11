VERSION 5.00
Begin VB.MDIForm MDIFrmPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Almacenes GTZ - Udapre"
   ClientHeight    =   8400
   ClientLeft      =   -45
   ClientTop       =   525
   ClientWidth     =   12000
   Icon            =   "MDIFrmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnuClasificadores 
      Caption         =   "Clasificadores"
      Begin VB.Menu mnuGrupos 
         Caption         =   "Grupos"
      End
      Begin VB.Menu mnuLineaA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaterial 
         Caption         =   "Detalle"
      End
      Begin VB.Menu LineaD 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDestinos 
         Caption         =   "Destinos"
      End
   End
   Begin VB.Menu mnuRegistro 
      Caption         =   "Registro"
      Begin VB.Menu mnuIngreso 
         Caption         =   "Ingreso"
         Begin VB.Menu mnuManual 
            Caption         =   "Ingreso Manual"
         End
         Begin VB.Menu mnuCompras 
            Caption         =   "Ingreso de Compras"
         End
      End
      Begin VB.Menu mnuLineaB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalida 
         Caption         =   "Entrega"
      End
   End
   Begin VB.Menu mnuControl 
      Caption         =   "Control"
      Begin VB.Menu mnuInventario 
         Caption         =   "Inventario"
      End
      Begin VB.Menu mnuLineaC 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlmacen 
         Caption         =   "Estado Almacen"
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "Salir"
      Begin VB.Menu mnuSalirApp 
         Caption         =   "Salir de la Aplicación"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "MDIFrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9090
    Me.Width = 12120
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    TerminaTodo
End Sub
Private Sub mnuAlmacen_Click()
    ALFrmAlmacen.Show
End Sub
Private Sub mnuCompras_Click()
    With ALFrmIngDeLici
        .Show vbModal
        If .QResp Then
            With AlFrmIngresoMaterial
                .ALPrincipal 2, IngresoDeLicitacion(ALFrmIngDeLici.NoLicitacion)
            End With
        End If
    End With
End Sub
Private Sub mnuDestinos_Click()
    ALFrmCLDestinos.Show
End Sub
Private Sub mnuGrupos_Click()
    AlmFrmCLGrupos.Show
End Sub
Private Sub mnuInventario_Click()
    AlmFrmInventario.Show
End Sub
Private Sub mnuManual_Click()
    AlFrmIngresoMaterial.ALPrincipal 0
End Sub
Private Sub mnuMaterial_Click()
    AlFrmCreaMaterial.ALPrincipal 0
End Sub
Private Sub mnuSalida_Click()
    AlmFrmSalidaMaterial.ALPrincipal 0
End Sub
Private Sub mnuSalirApp_Click()
    End
End Sub

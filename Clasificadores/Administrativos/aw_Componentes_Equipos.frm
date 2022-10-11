VERSION 5.00
Begin VB.Form aw_Componentes_Equipos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clasificadores - Administrativos - Caracteristicas de Equipos"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8940
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "COMPONENTES DE LOS EQUIPOS"
      ForeColor       =   &H00FF0000&
      Height          =   5175
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   4335
      Begin VB.OptionButton Option17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SEÑALIZACION (EQUIPO)"
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   4680
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.OptionButton Option16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BOTONERIA (EQUIPO)"
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   4200
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.OptionButton Option15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MOTOR DEL EQUIPO"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   3720
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ESTETICA DE LA CABINA"
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   3375
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CONDICION DE LA CABINA"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   3375
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CONTROL DE MAQUINA"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   1320
         Width           =   3375
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CUARTO DE CONTROL"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   1800
         Width           =   3375
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "GRUPO DE COCHES"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   2280
         Width           =   3375
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SISTEMA DE PUERTAS"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   2760
         Width           =   3375
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TIPO DE PUERTA PISO"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   3240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PROPIEDADES DE LOS EQUIPOS"
      ForeColor       =   &H00FF0000&
      Height          =   5175
      Left            =   4440
      TabIndex        =   4
      Top             =   960
      Width           =   4455
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CONDICION DE VENTA"
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   2760
         Width           =   3375
      End
      Begin VB.OptionButton Option14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CAPACIDAD DEL EQUIPO (Peso, Pasajeros)"
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   3240
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "VELOCIDAD DE LOS EQUIPOS"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   2280
         Width           =   3375
      End
      Begin VB.OptionButton Option13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TIPOS DE EQUIPOS"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   3375
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MARCAS DE EQUIPOS Y OTROS BIENES"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   3735
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MODELOS DE EQUIPOS"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "LINEAS DE EQUIPOS (TECNOLOGIA)"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   3375
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H80000011&
      FillColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   8835
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Height          =   600
         Left            =   7440
         MaskColor       =   &H00000000&
         Picture         =   "aw_Componentes_Equipos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancelar"
         Top             =   100
         Width           =   1365
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H80000015&
         Height          =   600
         Left            =   0
         Picture         =   "aw_Componentes_Equipos.frx":07C2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   100
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARACTERISTICAS DE EQUIPOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   360
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   7275
      End
   End
End
Attribute VB_Name = "aw_Componentes_Equipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnCancelar_Click()
    aw_Componentes_Equipos.Visible = False
End Sub

Private Sub Option1_Click()
    frm_ac_bienes_tecnologia_linea.lbl_titulo = Option1.Caption
    frm_ac_bienes_tecnologia_linea.FraNavega = Option1.Caption
    frm_ac_bienes_tecnologia_linea.lbl_titulo2 = Option1.Caption
    frm_ac_bienes_tecnologia_linea.Show
End Sub

Private Sub Option10_Click()
    aw_p_ac_bienes_equipo_velocidad.lbl_titulo = Option10.Caption
    aw_p_ac_bienes_equipo_velocidad.FraNavega = Option10.Caption
    aw_p_ac_bienes_equipo_velocidad.lbl_titulo2 = Option10.Caption
    aw_p_ac_bienes_equipo_velocidad.Show
End Sub

Private Sub Option11_Click()
    aw_p_ac_bienes_marcas.lbl_titulo = Option11.Caption
    aw_p_ac_bienes_marcas.FraNavega = Option11.Caption
    aw_p_ac_bienes_marcas.lbl_titulo2 = Option11.Caption
    aw_p_ac_bienes_marcas.Show
End Sub

Private Sub Option12_Click()
    frm_ac_bienes_modelos.lbl_titulo = Option12.Caption
    frm_ac_bienes_modelos.FraNavega = Option12.Caption
    frm_ac_bienes_modelos.lbl_titulo2 = Option12.Caption
    frm_ac_bienes_modelos.Show
End Sub

Private Sub Option13_Click()
    aw_p_ac_bienes_equipo_tipos.lbl_titulo = Option13.Caption
    aw_p_ac_bienes_equipo_tipos.FraNavega = Option13.Caption
    aw_p_ac_bienes_equipo_tipos.lbl_titulo2 = Option13.Caption
    aw_p_ac_bienes_equipo_tipos.Show
End Sub

Private Sub Option2_Click()
    aw_p_ac_bienes_equipo_cabina_estetica.lbl_titulo = Option2.Caption
    aw_p_ac_bienes_equipo_cabina_estetica.FraNavega = Option2.Caption
    aw_p_ac_bienes_equipo_cabina_estetica.lbl_titulo2 = Option2.Caption
    aw_p_ac_bienes_equipo_cabina_estetica.Show
End Sub

Private Sub Option3_Click()
    aw_p_ac_bienes_equipo_condicion_cabina.lbl_titulo = Option3.Caption
    aw_p_ac_bienes_equipo_condicion_cabina.FraNavega = Option3.Caption
    aw_p_ac_bienes_equipo_condicion_cabina.lbl_titulo2 = Option3.Caption
    aw_p_ac_bienes_equipo_condicion_cabina.Show
End Sub

Private Sub Option4_Click()
    aw_p_ac_bienes_equipo_condicion_ventas.lbl_titulo = Option4.Caption
    aw_p_ac_bienes_equipo_condicion_ventas.FraNavega = Option4.Caption
    aw_p_ac_bienes_equipo_condicion_ventas.lbl_titulo2 = Option4.Caption
    aw_p_ac_bienes_equipo_condicion_ventas.Show
End Sub

Private Sub Option5_Click()
    aw_p_ac_bienes_equipo_ctrl_maquina.lbl_titulo = Option5.Caption
    aw_p_ac_bienes_equipo_ctrl_maquina.FraNavega = Option5.Caption
    aw_p_ac_bienes_equipo_ctrl_maquina.lbl_titulo2 = Option5.Caption
    aw_p_ac_bienes_equipo_ctrl_maquina.Show
End Sub

Private Sub Option6_Click()
    aw_p_ac_bienes_equipo_cuadro_ctrl.lbl_titulo = Option6.Caption
    aw_p_ac_bienes_equipo_cuadro_ctrl.FraNavega = Option6.Caption
    aw_p_ac_bienes_equipo_cuadro_ctrl.lbl_titulo2 = Option6.Caption
    aw_p_ac_bienes_equipo_cuadro_ctrl.Show
End Sub

Private Sub Option7_Click()
    aw_p_ac_bienes_equipo_grupo_coches.lbl_titulo = Option7.Caption
    aw_p_ac_bienes_equipo_grupo_coches.FraNavega = Option7.Caption
    aw_p_ac_bienes_equipo_grupo_coches.lbl_titulo2 = Option7.Caption
    aw_p_ac_bienes_equipo_grupo_coches.Show
End Sub

Private Sub Option8_Click()
    aw_p_ac_bienes_equipo_sistema_puertas.lbl_titulo = Option8.Caption
    aw_p_ac_bienes_equipo_sistema_puertas.FraNavega = Option8.Caption
    aw_p_ac_bienes_equipo_sistema_puertas.lbl_titulo2 = Option8.Caption
    aw_p_ac_bienes_equipo_sistema_puertas.Show
End Sub

Private Sub Option9_Click()
    aw_p_ac_bienes_equipo_tipo_puerta_piso.lbl_titulo = Option9.Caption
    aw_p_ac_bienes_equipo_tipo_puerta_piso.FraNavega = Option9.Caption
    aw_p_ac_bienes_equipo_tipo_puerta_piso.lbl_titulo2 = Option9.Caption
    aw_p_ac_bienes_equipo_tipo_puerta_piso.Show
End Sub

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tw_tecnico_venta 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Tecnico - Venta de Servicios"
   ClientHeight    =   11100
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   15480
   ForeColor       =   &H00FFFFC0&
   Icon            =   "tw_tecnico_venta.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   13065
   ScaleMode       =   0  'User
   ScaleWidth      =   3.98351e5
   WindowState     =   2  'Maximized
   Begin VB.Frame FraAnula 
      BackColor       =   &H00404040&
      Caption         =   "Registra Justificacion para Anulación de Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   2535
      Left            =   8520
      TabIndex        =   200
      Top             =   4440
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton BtnGrabar2 
         BackColor       =   &H00C0FFFF&
         Height          =   635
         Left            =   2040
         Picture         =   "tw_tecnico_venta.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   1560
         Width           =   1365
      End
      Begin VB.CommandButton BtnCancelar2 
         BackColor       =   &H00C0FFFF&
         Height          =   635
         Left            =   4080
         MaskColor       =   &H00000000&
         Picture         =   "tw_tecnico_venta.frx":11F0
         Style           =   1  'Graphical
         TabIndex        =   202
         ToolTipText     =   "Cancelar"
         Top             =   1560
         Width           =   1365
      End
      Begin VB.TextBox TxtAnula 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos16"
         ForeColor       =   &H00000000&
         Height          =   885
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   201
         Text            =   "tw_tecnico_venta.frx":1ADC
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   204
         Top             =   1425
         Width           =   2025
      End
   End
   Begin VB.Frame frm_benef 
      BackColor       =   &H00404040&
      Caption         =   "Registra Datos del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   2535
      Left            =   8520
      TabIndex        =   193
      Top             =   4440
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox TxtCelular 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos16"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   199
         Text            =   "0"
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox TxtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos16"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   198
         Text            =   "0"
         Top             =   480
         Width           =   3975
      End
      Begin VB.CommandButton BtnCancelarBen 
         BackColor       =   &H00C0FFFF&
         Height          =   635
         Left            =   4080
         MaskColor       =   &H00000000&
         Picture         =   "tw_tecnico_venta.frx":1ADE
         Style           =   1  'Graphical
         TabIndex        =   195
         ToolTipText     =   "Cancelar"
         Top             =   1440
         Width           =   1365
      End
      Begin VB.CommandButton BtnGrabarBen 
         BackColor       =   &H00C0FFFF&
         Height          =   635
         Left            =   2040
         Picture         =   "tw_tecnico_venta.frx":23CA
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Correo Electrónico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   720
         TabIndex        =   197
         Top             =   465
         Width           =   2025
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono Celular:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   720
         TabIndex        =   196
         Top             =   960
         Width           =   2025
      End
   End
   Begin VB.Frame fra_reportes 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Elija una de las opciones ..."
      ForeColor       =   &H00FF0000&
      Height          =   4455
      Left            =   6720
      TabIndex        =   143
      Top             =   1200
      Visible         =   0   'False
      Width           =   11775
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ninguno"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   208
         Top             =   3480
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton opt_vigentes 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Solo datos de Contratos VIGENTES"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   151
         Top             =   960
         Width           =   4215
      End
      Begin VB.OptionButton opt_vigentes_eqp1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Contratos VIGENTES con detalle de EQUIPOS por Zona Geográfica"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   150
         Top             =   600
         Width           =   5415
      End
      Begin VB.OptionButton opt_todos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Solo datos de Contratos (TODOS)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   149
         Top             =   1320
         Width           =   3615
      End
      Begin VB.OptionButton opt_vigentes_eqp2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Contratos VIGENTES con detalle de EQUIPOS (Zona Piloto)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   148
         Top             =   1680
         Width           =   5175
      End
      Begin VB.OptionButton opt_salir 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Salir"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   147
         Top             =   3120
         Width           =   1215
      End
      Begin VB.OptionButton opt_vigentes_eqp3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Contratos VIGENTES con detalle de EQUIPOS (Zona Piloto) MIGRAR"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   146
         Top             =   2040
         Width           =   5295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Certificado de Cumplimiento de Contrato por Mantenimiento Integral"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   145
         Top             =   2760
         Width           =   5295
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Contratos VIGENTES con detalle de BIENES"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   144
         Top             =   2400
         Width           =   5175
      End
      Begin Crystal.CrystalReport Cr_otros 
         Left            =   120
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
   End
   Begin VB.PictureBox BtnImprimir1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   360
      Picture         =   "tw_tecnico_venta.frx":2BB8
      ScaleHeight     =   615
      ScaleWidth      =   1575
      TabIndex        =   127
      ToolTipText     =   "Imprime Orden de Servicio"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20520
      TabIndex        =   100
      Top             =   0
      Width           =   20520
      Begin VB.PictureBox BtnAprobar3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   12240
         Picture         =   "tw_tecnico_venta.frx":36BC
         ScaleHeight     =   615
         ScaleWidth      =   1440
         TabIndex        =   212
         ToolTipText     =   "Aprueba el Contrato Elegido (ya NO podrá ser modificado)"
         Top             =   20
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.PictureBox BtnVer2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   10800
         Picture         =   "tw_tecnico_venta.frx":416A
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   211
         ToolTipText     =   "Registra Adenda o Modificación al Contrato"
         Top             =   20
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox BtnDesAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7800
         Picture         =   "tw_tecnico_venta.frx":4F4B
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   126
         ToolTipText     =   "Cambiar Contrato a Provisional o Viceversa"
         Top             =   20
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6375
         Picture         =   "tw_tecnico_venta.frx":5A95
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   125
         ToolTipText     =   "Cerrar Tramite y Archivarlo"
         Top             =   20
         Width           =   1395
      End
      Begin VB.PictureBox BtnVer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   9315
         Picture         =   "tw_tecnico_venta.frx":654F
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   124
         ToolTipText     =   "Registra Adenda o Modificación al Contrato"
         Top             =   20
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5040
         Picture         =   "tw_tecnico_venta.frx":6E06
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   106
         ToolTipText     =   "Imprimir el Listado de los Registros"
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   18480
         Picture         =   "tw_tecnico_venta.frx":76D3
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   107
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3840
         Picture         =   "tw_tecnico_venta.frx":7E95
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   105
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2520
         Picture         =   "tw_tecnico_venta.frx":864A
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   104
         ToolTipText     =   "Aprueba el Registro Elegido"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1320
         Picture         =   "tw_tecnico_venta.frx":8E80
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   103
         ToolTipText     =   "Anula Todo el Tramite"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   -15
         Picture         =   "tw_tecnico_venta.frx":95CC
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   102
         ToolTipText     =   "Modifica datos del Contrato elegido"
         Top             =   0
         Width           =   1430
      End
      Begin VB.Label lbl_titulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRONOGRAMA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   13695
         TabIndex        =   108
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   1875
      TabIndex        =   78
      Top             =   7275
      Width           =   1935
      Begin VB.PictureBox BtnImprimir2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   360
         Picture         =   "tw_tecnico_venta.frx":9EE1
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   131
         ToolTipText     =   "Imprimir el Solicitud de Facturación"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnAprobar2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         Picture         =   "tw_tecnico_venta.frx":A921
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   130
         ToolTipText     =   "Aprueba la Cuota Elegida para Facturación"
         Top             =   480
         Width           =   1320
      End
      Begin VB.PictureBox BtnModDetalle2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         Picture         =   "tw_tecnico_venta.frx":B157
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   129
         ToolTipText     =   "Modifica la Cuota Identifiacada"
         Top             =   0
         Width           =   1430
      End
      Begin VB.CommandButton BtnAnlDetalle2 
         BackColor       =   &H80000015&
         Height          =   525
         Left            =   240
         Picture         =   "tw_tecnico_venta.frx":BA6C
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Solicita Anulación de la FACTURA de la Cuota elegida"
         Top             =   1275
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton BtnAddDetalle2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nuevo->"
         Height          =   640
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Registra una Nueva Cobranza"
         Top             =   45
         Visible         =   0   'False
         Width           =   765
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   1425
      Left            =   120
      Negotiate       =   -1  'True
      ScaleHeight     =   5.688
      ScaleMode       =   4  'Character
      ScaleWidth      =   15.625
      TabIndex        =   75
      Top             =   5805
      Width           =   1935
      Begin VB.PictureBox BtnImprimir4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   192
         Picture         =   "tw_tecnico_venta.frx":C1B8
         ScaleHeight     =   615
         ScaleWidth      =   1575
         TabIndex        =   128
         ToolTipText     =   "Imprime Nota de Devolucion"
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton BtnAnlDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Anular-->"
         Enabled         =   0   'False
         Height          =   640
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Anula la Cobranza Identificada"
         Top             =   825
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnAddDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Repuestos Devueltos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   255
         MaskColor       =   &H00FFFF00&
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Adiciona Detalle"
         Top             =   150
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5010
      Left            =   6600
      TabIndex        =   9
      Top             =   765
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   8837
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12632256
      ForeColor       =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Registro DATOS CONTRATO"
      TabPicture(0)   =   "tw_tecnico_venta.frx":CDD7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrmCabecera"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Registro DATOS CRONO.MTTO."
      TabPicture(1)   =   "tw_tecnico_venta.frx":CDF3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmEdita"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Registro DE CUOTAS (Cobranza)"
      TabPicture(2)   =   "tw_tecnico_venta.frx":CE0F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrmCobros"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Registro ALCANCE CONTRATO"
      TabPicture(3)   =   "tw_tecnico_venta.frx":CE2B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrmAlcance"
      Tab(3).Control(1)=   "FraGrabarCancelar1"
      Tab(3).Control(2)=   "FrmABMDet1"
      Tab(3).ControlCount=   3
      Begin VB.PictureBox FrmABMDet1 
         BackColor       =   &H80000015&
         FillColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   -74880
         Negotiate       =   -1  'True
         ScaleHeight     =   3.313
         ScaleMode       =   4  'Character
         ScaleWidth      =   97.625
         TabIndex        =   154
         Top             =   3960
         Width           =   11775
         Begin VB.PictureBox BtnAddDetalle1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3960
            Picture         =   "tw_tecnico_venta.frx":CE47
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   156
            ToolTipText     =   "Genera nuevos items del Alcance del Conrato"
            Top             =   120
            Width           =   1430
         End
         Begin VB.PictureBox BtnModDetalle1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5760
            Picture         =   "tw_tecnico_venta.frx":D606
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   155
            ToolTipText     =   "Modifica datos del Alcance del Contrato"
            Top             =   120
            Width           =   1430
         End
      End
      Begin VB.PictureBox FraGrabarCancelar1 
         BackColor       =   &H80000015&
         FillColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   -74880
         ScaleHeight     =   705
         ScaleWidth      =   11715
         TabIndex        =   157
         Top             =   4080
         Visible         =   0   'False
         Width           =   11775
         Begin VB.PictureBox BtnGrabar1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   9360
            Picture         =   "tw_tecnico_venta.frx":DF1B
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   159
            Top             =   50
            Width           =   1280
         End
         Begin VB.PictureBox BtnFlecha01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   590
            Left            =   7320
            Picture         =   "tw_tecnico_venta.frx":E709
            ScaleHeight     =   585
            ScaleWidth      =   825
            TabIndex        =   158
            ToolTipText     =   "Genera nuevos items del Alcance del Conrato"
            Top             =   0
            Width           =   825
         End
         Begin VB.Label LblAyuda01 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Para Modificar la ""Fecha.Inicio y Fecha.Fin"", edite la casilla y digite Enter..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   240
            Left            =   600
            TabIndex        =   161
            Top             =   240
            Width           =   6675
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "... Luego -->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   240
            Left            =   8280
            TabIndex        =   160
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.Frame FrmAlcance 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3375
         Left            =   -74880
         TabIndex        =   152
         Top             =   480
         Width           =   11775
         Begin MSDataGridLib.DataGrid DtgAlcance 
            Bindings        =   "tw_tecnico_venta.frx":EE2F
            Height          =   2985
            Left            =   120
            Negotiate       =   -1  'True
            TabIndex        =   153
            Top             =   240
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   5265
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   13
            FormatLocked    =   -1  'True
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "venta_codigo"
               Caption         =   "Cod_Venta"
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
               DataField       =   "solicitud_tipo"
               Caption         =   "Codigo"
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
               DataField       =   "solicitud_tipo_descripcion"
               Caption         =   "Descripcion del Alcance"
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
            BeginProperty Column03 
               DataField       =   "unidad_codigo_tec"
               Caption         =   "Unidad.Ejecutora"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   4105
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "fecha_inicio_alcance"
               Caption         =   "Fecha.Inicio"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "fecha_fin_alcance"
               Caption         =   "Fecha.Fin"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "venta_tiempo_dias"
               Caption         =   "Dias Calendario"
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
            BeginProperty Column07 
               DataField       =   "estado_codigo"
               Caption         =   "Estado"
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
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   615.118
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   5040
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1365.165
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  DividerStyle    =   1
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  DividerStyle    =   1
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column06 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1184.882
               EndProperty
               BeginProperty Column07 
                  Alignment       =   2
                  ColumnWidth     =   645.165
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrmCobros 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4590
         Left            =   -75000
         TabIndex        =   27
         Top             =   380
         Width           =   11895
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   11380
            TabIndex        =   164
            Top             =   2895
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_codigo2A 
            Bindings        =   "tw_tecnico_venta.frx":EE48
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   9840
            TabIndex        =   59
            Top             =   2880
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   14737632
            ForeColor       =   0
            ListField       =   "beneficiario_nit"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.CommandButton CmdEmail 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9480
            Picture         =   "tw_tecnico_venta.frx":EE61
            Style           =   1  'Graphical
            TabIndex        =   192
            Top             =   2880
            Width           =   375
         End
         Begin VB.TextBox TxtDscto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "cobranza_fecha_prog"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            DataSource      =   "Ado_datos16"
            Height          =   300
            Left            =   7680
            TabIndex        =   5
            Text            =   "0"
            Top             =   195
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo dtc_email2A 
            Bindings        =   "tw_tecnico_venta.frx":F863
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   6720
            TabIndex        =   162
            Top             =   2880
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   14737632
            ForeColor       =   0
            ListField       =   "beneficiario_email"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.ComboBox CmbEmision 
            DataSource      =   "Ado_datos16"
            Height          =   315
            ItemData        =   "tw_tecnico_venta.frx":F87C
            Left            =   6360
            List            =   "tw_tecnico_venta.frx":F889
            TabIndex        =   141
            Text            =   "FACTURA FISICA"
            Top             =   2400
            Width           =   2985
         End
         Begin MSDataListLib.DataCombo dtc_desc2A 
            Bindings        =   "tw_tecnico_venta.frx":F8BB
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   1575
            TabIndex        =   60
            Top             =   2880
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.PictureBox Frame7 
            BackColor       =   &H80000015&
            FillColor       =   &H00FFFFFF&
            Height          =   705
            Left            =   60
            ScaleHeight     =   645
            ScaleWidth      =   11775
            TabIndex        =   119
            Top             =   3840
            Width           =   11840
            Begin VB.PictureBox CmdCancelaCobro 
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   6120
               Picture         =   "tw_tecnico_venta.frx":F8D4
               ScaleHeight     =   615
               ScaleWidth      =   1455
               TabIndex        =   137
               Top             =   0
               Width           =   1455
            End
            Begin VB.PictureBox CmdGrabaCobro 
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   4560
               Picture         =   "tw_tecnico_venta.frx":101C0
               ScaleHeight     =   615
               ScaleWidth      =   1275
               TabIndex        =   136
               Top             =   0
               Width           =   1280
            End
         End
         Begin VB.TextBox txt_fojas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "nro_fojas"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Ado_datos16"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   9885
            Locked          =   -1  'True
            TabIndex        =   88
            Text            =   "0"
            Top             =   1990
            Width           =   1725
         End
         Begin MSComCtl2.DTPicker DTPFechaProg 
            DataField       =   "cobranza_fecha_prog"
            DataSource      =   "Ado_datos16"
            Height          =   285
            Left            =   7680
            TabIndex        =   87
            Top             =   120
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   109969409
            CurrentDate     =   44621
            MinDate         =   36526
         End
         Begin VB.TextBox txtDoc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "doc_numero"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Ado_datos16"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   9885
            Locked          =   -1  'True
            TabIndex        =   85
            Text            =   "0"
            Top             =   1320
            Width           =   1725
         End
         Begin VB.CheckBox Chk_plazo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Es requisito para el Plazo de entrega ?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   255
            TabIndex        =   73
            Top             =   3120
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.TextBox txt_plazo 
            CausesValidation=   0   'False
            DataField       =   "cobranza_concepto_plazo"
            DataSource      =   "Ado_datos16"
            Height          =   345
            Left            =   1575
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   71
            Top             =   3360
            Width           =   10035
         End
         Begin VB.TextBox TxtCobrador 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            CausesValidation=   0   'False
            DataField       =   "nombre_cobrador"
            DataSource      =   "Ado_datos16"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2175
            Locked          =   -1  'True
            MaxLength       =   60
            MultiLine       =   -1  'True
            TabIndex        =   30
            Top             =   1905
            Width           =   5175
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   9000
            TabIndex        =   39
            Top             =   1920
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_codigo4A 
            Bindings        =   "tw_tecnico_venta.frx":10996
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   7320
            TabIndex        =   62
            Top             =   1905
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   14737632
            ForeColor       =   0
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin VB.TextBox TxtDsctoTot 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "cobranza_programada_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Ado_datos16"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5685
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "0"
            Top             =   660
            Width           =   1545
         End
         Begin VB.TextBox TxtMontoDol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "cobranza_total_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Ado_datos16"
            Enabled         =   0   'False
            Height          =   285
            Left            =   8520
            TabIndex        =   28
            Text            =   "0"
            Top             =   720
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox TxtMonto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "cobranza_programada_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Ado_datos16"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2835
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "0"
            Top             =   660
            Width           =   1455
         End
         Begin VB.TextBox TxtObs 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            CausesValidation=   0   'False
            DataField       =   "cobranza_observaciones"
            DataSource      =   "Ado_datos16"
            Height          =   465
            Left            =   1575
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1155
            Width           =   7695
         End
         Begin MSDataListLib.DataCombo dtc_desc4A 
            Bindings        =   "tw_tecnico_venta.frx":109B0
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   2175
            TabIndex        =   61
            Top             =   1905
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSComCtl2.DTPicker DTPFechaCobro 
            DataField       =   "cobranza_fecha_cobro"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   2175
            TabIndex        =   66
            Top             =   2400
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CheckBox        =   -1  'True
            Format          =   109969409
            CurrentDate     =   44600
            MaxDate         =   47848
            MinDate         =   36526
         End
         Begin MSComCtl2.DTPicker DTPFechaConf 
            DataField       =   "cobranza_fecha_conformidad"
            DataSource      =   "Ado_datos16"
            Height          =   285
            Left            =   9840
            TabIndex        =   86
            Top             =   675
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   109969409
            CurrentDate     =   44621
            MinDate         =   36526
         End
         Begin MSDataListLib.DataCombo dtc_benef2A 
            Bindings        =   "tw_tecnico_venta.frx":109CA
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   9720
            TabIndex        =   138
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail               NIT"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9435
            TabIndex        =   163
            Top             =   2640
            Width           =   1395
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_fecha_conformidad"
            DataSource      =   "Ado_datos16"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   9885
            TabIndex        =   64
            Top             =   450
            Width           =   1725
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de Emisión de la Factura:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4040
            TabIndex        =   140
            Top             =   2420
            Width           =   2265
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Estimada a Cobrar:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   139
            Top             =   2420
            Width           =   1830
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de Fojas"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9840
            TabIndex        =   89
            Top             =   1755
            Width           =   1755
         End
         Begin VB.Label lblfechaCertif 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Certificado"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9840
            TabIndex        =   84
            Top             =   195
            Width           =   1755
         End
         Begin VB.Label TxtNroVentaC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "venta_codigo"
            DataSource      =   "Ado_datos16"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   1200
            TabIndex        =   74
            Top             =   195
            Width           =   1365
         End
         Begin VB.Label lbl_plazo 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto Factura:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   225
            TabIndex        =   72
            Top             =   3360
            Width           =   1560
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "USD (Dol)"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   4695
            TabIndex        =   70
            Top             =   660
            Width           =   915
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00800000&
            X1              =   0
            X2              =   9600
            Y1              =   1755
            Y2              =   1755
         End
         Begin VB.Label lblccertif 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "No.Doc.Certificado"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9840
            TabIndex        =   69
            Top             =   1030
            Width           =   1755
         End
         Begin VB.Label Txt_cod_cobro 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_prog_codigo"
            DataSource      =   "Ado_datos16"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4020
            TabIndex        =   68
            Top             =   195
            Width           =   885
         End
         Begin VB.Label Lbl_nombre_fac 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente a Facturar:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   225
            TabIndex        =   67
            Top             =   2895
            Width           =   1350
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Cuota:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   3000
            TabIndex        =   65
            Top             =   210
            Width           =   990
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00800000&
            X1              =   9600
            X2              =   9600
            Y1              =   -60
            Y2              =   2640
         End
         Begin VB.Label lbl_fechas 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Programada de la Cuota:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5360
            TabIndex        =   36
            Top             =   210
            Width           =   2265
         End
         Begin VB.Label Lbl_Cobrador 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Encargado de la Cobranza:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   35
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label48 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "BOB (Bs)"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1935
            TabIndex        =   34
            Top             =   660
            Width           =   855
         End
         Begin VB.Label lbl_monto 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Importe de la Cuota:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   33
            Top             =   660
            Width           =   1425
         End
         Begin VB.Label lbl_obs 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto Cuota:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   225
            TabIndex        =   32
            Top             =   1200
            Width           =   1320
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Venta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   225
            TabIndex        =   31
            Top             =   210
            Width           =   1110
         End
      End
      Begin VB.Frame FrmEdita 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4590
         Left            =   -74940
         TabIndex        =   26
         Top             =   380
         Width           =   11895
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   1335
            Left            =   240
            TabIndex        =   205
            Top             =   2520
            Width           =   11415
            Begin VB.OptionButton Option10 
               BackColor       =   &H00C0C0C0&
               Caption         =   "1. Programar en meses IMPARES (ENE, MAR, MAY, JUL, SEP, NOV), los insumos: 3.(Aceite.680) y/o 4.(Aceite.20/50)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   210
               Left            =   480
               TabIndex        =   207
               Top             =   360
               Width           =   9975
            End
            Begin VB.OptionButton Option11 
               BackColor       =   &H00C0C0C0&
               Caption         =   "2. Programar en meses PARES (FEB, ABR, JUN, AGO, OCT, DIC), los insumos: 3.(Aceite.680) y/o 4.(Aceite.20/50)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   225
               Left            =   480
               TabIndex        =   206
               Top             =   840
               Width           =   10095
            End
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "tw_tecnico_venta.frx":109E3
            DataField       =   "beneficiario_codigo_tec"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5100
            TabIndex        =   183
            Top             =   660
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "0"
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   $"tw_tecnico_venta.frx":109FC
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1035
            Left            =   240
            TabIndex        =   174
            Top             =   1320
            Width           =   11415
            Begin VB.ComboBox cmb_mes_ini_tec 
               DataField       =   "mes_inicio_crono_tec"
               DataSource      =   "Ado_datos"
               Height          =   315
               ItemData        =   "tw_tecnico_venta.frx":10A83
               Left            =   7200
               List            =   "tw_tecnico_venta.frx":10AAB
               TabIndex        =   178
               Text            =   "SEPTIEMBRE"
               Top             =   360
               Width           =   1900
            End
            Begin VB.TextBox txt_cant 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               DataField       =   "cantidad_periodos_tec"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   360
               TabIndex        =   177
               Text            =   "0"
               Top             =   360
               Width           =   855
            End
            Begin VB.ComboBox cmd_unimed_tec 
               BackColor       =   &H00FFFFFF&
               DataField       =   "unimed_codigo_tec"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "tw_tecnico_venta.frx":10B14
               Left            =   1560
               List            =   "tw_tecnico_venta.frx":10B2A
               TabIndex        =   176
               Text            =   "ANUAL"
               Top             =   360
               Width           =   1215
            End
            Begin VB.ComboBox cmb_dia 
               DataField       =   "dia_nombre"
               DataSource      =   "Ado_datos"
               Height          =   315
               ItemData        =   "tw_tecnico_venta.frx":10B52
               Left            =   360
               List            =   "tw_tecnico_venta.frx":10B6E
               TabIndex        =   175
               Text            =   "AUTOMATICO"
               Top             =   720
               Visible         =   0   'False
               Width           =   1575
            End
            Begin MSComCtl2.DTPicker lbl_fecha_ini 
               DataField       =   "fecha_inicio_tec"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd-MMM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   3
               EndProperty
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   3120
               TabIndex        =   179
               Top             =   360
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   503
               _Version        =   393216
               CalendarBackColor=   -2147483646
               CheckBox        =   -1  'True
               Format          =   109969409
               CurrentDate     =   44197
               MinDate         =   36526
            End
            Begin MSComCtl2.DTPicker lbl_fecha_fin 
               DataField       =   "fecha_fin_tec"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd-MMM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   3
               EndProperty
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   5160
               TabIndex        =   180
               Top             =   360
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   503
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   109969409
               CurrentDate     =   44561
               MinDate         =   36526
            End
            Begin VB.Label LblParImpar 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "NO ASIGNADO"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   9480
               TabIndex        =   181
               Top             =   360
               Width           =   1605
            End
         End
         Begin VB.PictureBox FraGrabarDet 
            BackColor       =   &H80000015&
            FillColor       =   &H00FFFFFF&
            Height          =   825
            Left            =   0
            ScaleHeight     =   765
            ScaleWidth      =   11835
            TabIndex        =   116
            Top             =   3720
            Visible         =   0   'False
            Width           =   11895
            Begin VB.CommandButton CmdGrabaDet 
               BackColor       =   &H80000015&
               Height          =   555
               Left            =   4680
               Picture         =   "tw_tecnico_venta.frx":10BBA
               Style           =   1  'Graphical
               TabIndex        =   117
               Top             =   120
               Width           =   1245
            End
            Begin VB.CommandButton CmdCancelaDet 
               BackColor       =   &H80000015&
               Height          =   555
               Left            =   5900
               MaskColor       =   &H00000000&
               Picture         =   "tw_tecnico_venta.frx":11390
               Style           =   1  'Graphical
               TabIndex        =   118
               ToolTipText     =   "Cancelar"
               Top             =   120
               Width           =   1365
            End
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   7800
            TabIndex        =   38
            Top             =   1455
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "tw_tecnico_venta.frx":11C7C
            DataField       =   "beneficiario_codigo_tec"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   240
            TabIndex        =   182
            Top             =   660
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_aux4 
            Bindings        =   "tw_tecnico_venta.frx":11C95
            DataField       =   "beneficiario_codigo_tec"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5520
            TabIndex        =   184
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "beneficiario_iniciales"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo dtc_desc7 
            Bindings        =   "tw_tecnico_venta.frx":11CAE
            DataField       =   "zpiloto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7320
            TabIndex        =   186
            Top             =   660
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "zpiloto_descripcion"
            BoundColumn     =   "zpiloto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo7 
            Bindings        =   "tw_tecnico_venta.frx":11CC7
            DataField       =   "zpiloto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6720
            TabIndex        =   187
            Top             =   660
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            ListField       =   "zpiloto_codigo"
            BoundColumn     =   "zpiloto_codigo"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo dtc_aux7 
            Bindings        =   "tw_tecnico_venta.frx":11CE0
            DataField       =   "zpiloto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   10680
            TabIndex        =   191
            Top             =   360
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            ListField       =   "mes_par_impar"
            BoundColumn     =   "zpiloto_codigo"
            Text            =   "0"
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Zona Piloto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6840
            TabIndex        =   188
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Responsable/Supervisor del Servicio Técnico:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   185
            Top             =   360
            Width           =   4200
         End
      End
      Begin VB.Frame FrmCabecera 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4590
         Left            =   60
         TabIndex        =   14
         Top             =   380
         Width           =   11895
         Begin VB.TextBox Text10 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   290
            Left            =   6915
            TabIndex        =   57
            Top             =   990
            Width           =   270
         End
         Begin MSDataListLib.DataCombo dtc_aux3 
            Bindings        =   "tw_tecnico_venta.frx":11CF9
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5880
            TabIndex        =   54
            Top             =   975
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ListField       =   "edif_codigo_corto"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin VB.TextBox Txt_Adenda 
            DataField       =   "literal_a"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   960
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   123
            Top             =   2520
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.TextBox Txt_campo2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "unidad_codigo_ant"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7755
            TabIndex        =   98
            Text            =   "0"
            Top             =   320
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "tw_tecnico_venta.frx":11D12
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5835
            TabIndex        =   55
            Top             =   660
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "tw_tecnico_venta.frx":11D2B
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   165
            TabIndex        =   56
            Top             =   975
            Width           =   6075
            _ExtentX        =   10716
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin VB.TextBox Text13 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   280
            Left            =   7275
            TabIndex        =   63
            Top             =   330
            Width           =   260
         End
         Begin VB.TextBox Text11 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   290
            Left            =   11490
            TabIndex        =   58
            Top             =   990
            Width           =   260
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "tw_tecnico_venta.frx":11D44
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   10305
            TabIndex        =   52
            Top             =   660
            Visible         =   0   'False
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.Frame Fra_datos 
            BackColor       =   &H00C0C0C0&
            Caption         =   "- EMPRESA ------------------------------------------------------------------------------------------ Cobrador "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1965
            Left            =   60
            TabIndex        =   22
            Top             =   1335
            Width           =   11775
            Begin VB.ComboBox cmd_unimed2 
               DataField       =   "unimed_codigo_cobr"
               DataSource      =   "Ado_datos"
               Height          =   315
               ItemData        =   "tw_tecnico_venta.frx":11D5D
               Left            =   10160
               List            =   "tw_tecnico_venta.frx":11D73
               TabIndex        =   121
               Text            =   "ANUAL"
               Top             =   1560
               Width           =   1420
            End
            Begin VB.ComboBox cmb_mes_ini 
               DataField       =   "mes_inicio_crono"
               DataSource      =   "Ado_datos"
               Height          =   315
               ItemData        =   "tw_tecnico_venta.frx":11D9B
               Left            =   2400
               List            =   "tw_tecnico_venta.frx":11DC3
               TabIndex        =   120
               Text            =   "ENERO"
               Top             =   1560
               Width           =   1620
            End
            Begin VB.TextBox txtCantCobr 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               DataField       =   "venta_cantidad_cobr"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   7005
               TabIndex        =   94
               Text            =   "0"
               Top             =   1560
               Width           =   1140
            End
            Begin MSComCtl2.DTPicker DTPFechaIni 
               DataField       =   "venta_fecha_inicio"
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   7080
               TabIndex        =   90
               Top             =   735
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   503
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   109969409
               CurrentDate     =   44348
               MaxDate         =   401768
               MinDate         =   2
            End
            Begin VB.TextBox TxtConcepto 
               DataField       =   "venta_descripcion"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1020
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   2
               Top             =   1120
               Width           =   10515
            End
            Begin VB.TextBox TxtPlazo 
               Alignment       =   2  'Center
               DataField       =   "venta_plazo_dias_calendario"
               DataSource      =   "Ado_datos"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4200
               TabIndex        =   1
               Text            =   "0"
               Top             =   1560
               Visible         =   0   'False
               Width           =   1215
            End
            Begin MSDataListLib.DataCombo dtc_desc11 
               Bindings        =   "tw_tecnico_venta.frx":11E2C
               DataField       =   "venta_tipo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1320
               TabIndex        =   0
               Top             =   705
               Width           =   4200
               _ExtentX        =   7408
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "venta_tipo_descripcion"
               BoundColumn     =   "venta_tipo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_codigo11 
               Bindings        =   "tw_tecnico_venta.frx":11E46
               DataField       =   "venta_tipo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   4440
               TabIndex        =   37
               Top             =   495
               Visible         =   0   'False
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "venta_tipo"
               BoundColumn     =   "venta_tipo"
               Text            =   ""
            End
            Begin MSComCtl2.DTPicker DTPFechaFin 
               DataField       =   "venta_fecha_fin"
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   10060
               TabIndex        =   91
               Top             =   735
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   503
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   109969409
               CurrentDate     =   44348
               MinDate         =   36526
            End
            Begin MSDataListLib.DataCombo dtc_desc5 
               Bindings        =   "tw_tecnico_venta.frx":11E60
               DataField       =   "beneficiario_codigo_cobr"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   6720
               TabIndex        =   114
               Top             =   270
               Width           =   4845
               _ExtentX        =   8546
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_denominacion"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_codigo5 
               Bindings        =   "tw_tecnico_venta.frx":11E79
               DataField       =   "beneficiario_codigo_cobr"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   10080
               TabIndex        =   99
               Top             =   120
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_codigo"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "0"
            End
            Begin MSDataListLib.DataCombo dtc_desc8 
               Bindings        =   "tw_tecnico_venta.frx":11E92
               DataField       =   "codigo_empresa"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   240
               TabIndex        =   189
               Top             =   240
               Width           =   6000
               _ExtentX        =   10583
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "denominacion_empresa"
               BoundColumn     =   "codigo_empresa"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_codigo8 
               Bindings        =   "tw_tecnico_venta.frx":11EAB
               DataField       =   "codigo_empresa"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   5160
               TabIndex        =   190
               Top             =   0
               Visible         =   0   'False
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "codigo_empresa"
               BoundColumn     =   "codigo_empresa"
               Text            =   ""
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Venta"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   165
               TabIndex        =   113
               Top             =   765
               Width           =   1275
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Forma de Pago"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   8865
               TabIndex        =   96
               Top             =   1575
               Width           =   1305
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Número de Cuotas"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   5565
               TabIndex        =   95
               Top             =   1575
               Width           =   1320
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de Fin"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   9000
               TabIndex        =   93
               Top             =   765
               Width           =   1050
            End
            Begin VB.Label lbl_mes_ini 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Mes Inicio del Plan de Cuotas"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   165
               TabIndex        =   92
               Top             =   1575
               Width           =   2100
            End
            Begin VB.Label lbl_campo4 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Inicio"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   6065
               TabIndex        =   45
               Top             =   765
               Width           =   960
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Concepto:"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   165
               TabIndex        =   23
               Top             =   1155
               Width           =   900
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Fra_Total 
            BackColor       =   &H00C0C0C0&
            Caption         =   $"tw_tecnico_venta.frx":11EC4
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   1215
            Left            =   60
            TabIndex        =   15
            Top             =   3300
            Width           =   11775
            Begin VB.TextBox Text17 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_saldo_p_cobrar_dol"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "Ado_datos"
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
               Height          =   360
               Left            =   9885
               TabIndex        =   172
               Text            =   "0"
               Top             =   720
               Width           =   1545
            End
            Begin VB.TextBox Text16 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_cobrado_dol"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "Ado_datos"
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
               Height          =   360
               Left            =   7800
               TabIndex        =   171
               Text            =   "0"
               Top             =   720
               Width           =   1545
            End
            Begin VB.TextBox Text15 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_total_dol"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "Ado_datos"
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
               Height          =   360
               Left            =   5640
               TabIndex        =   170
               Text            =   "0"
               Top             =   720
               Width           =   1665
            End
            Begin VB.TextBox Text14 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_adenda_dol"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "Ado_datos"
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
               Height          =   360
               Left            =   3600
               TabIndex        =   169
               Text            =   "0"
               Top             =   720
               Width           =   1545
            End
            Begin VB.TextBox Text12 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_origen_dol"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "Ado_datos"
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
               Height          =   360
               Left            =   1440
               TabIndex        =   168
               Text            =   "0"
               Top             =   720
               Width           =   1545
            End
            Begin VB.TextBox Text7 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_origen_bs"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "Ado_datos"
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
               Height          =   360
               Left            =   1440
               TabIndex        =   133
               Text            =   "0"
               Top             =   300
               Width           =   1545
            End
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_adenda_bs"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "Ado_datos"
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
               Height          =   360
               Left            =   3600
               TabIndex        =   132
               Text            =   "0"
               Top             =   300
               Width           =   1545
            End
            Begin VB.TextBox txtTDC 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               DataField       =   "venta_tipo_cambio"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   3360
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   44
               Top             =   360
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox TxtCobrado 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_cobrado_bs"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "Ado_datos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               Left            =   7800
               TabIndex        =   19
               Text            =   "0"
               Top             =   300
               Width           =   1545
            End
            Begin VB.TextBox txtCantTotal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_cantidad_total"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "Ado_datos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1320
               TabIndex        =   18
               Text            =   "0"
               Top             =   300
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox TxtMontoBs 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_total_bs"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "Ado_datos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   360
               Left            =   5640
               TabIndex        =   17
               Text            =   "0"
               Top             =   300
               Width           =   1665
            End
            Begin VB.TextBox TxtBstotal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_saldo_p_cobrar_bs"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "Ado_datos"
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
               Height          =   360
               Left            =   9885
               TabIndex        =   16
               Text            =   "0"
               Top             =   300
               Width           =   1545
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "DOLARES:"
               ForeColor       =   &H00000040&
               Height          =   195
               Left            =   240
               TabIndex        =   167
               Top             =   840
               Width           =   1005
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "BOLIVIANOS:"
               ForeColor       =   &H00000040&
               Height          =   195
               Left            =   240
               TabIndex        =   166
               Top             =   360
               Width           =   1005
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label38 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "+/-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   3000
               TabIndex        =   135
               Top             =   360
               Width           =   525
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "="
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   5160
               TabIndex        =   134
               Top             =   360
               Width           =   405
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   7335
               TabIndex        =   21
               Top             =   345
               Width           =   405
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "="
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   9405
               TabIndex        =   20
               Top             =   345
               Width           =   405
            End
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "tw_tecnico_venta.frx":11F5D
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7365
            TabIndex        =   46
            Top             =   975
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "tw_tecnico_venta.frx":11F76
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6660
            TabIndex        =   49
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "unidad_codigo"
            BoundColumn     =   "unidad_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "tw_tecnico_venta.frx":11F8F
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1500
            TabIndex        =   50
            Top             =   315
            Width           =   6045
            _ExtentX        =   10663
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "unidad_descripcion"
            BoundColumn     =   "unidad_codigo"
            Text            =   "Todos"
         End
         Begin MSComCtl2.DTPicker DTPfechasol 
            DataField       =   "venta_fecha"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   8880
            TabIndex        =   101
            Top             =   960
            Visible         =   0   'False
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   109969409
            CurrentDate     =   44348
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo Dtc_aux2 
            Bindings        =   "tw_tecnico_venta.frx":11FA8
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   9480
            TabIndex        =   142
            Top             =   660
            Visible         =   0   'False
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   -2147483624
            ListField       =   "tipoben_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_deudor2 
            Bindings        =   "tw_tecnico_venta.frx":11FC1
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   8640
            TabIndex        =   165
            Top             =   660
            Visible         =   0   'False
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   255
            ForeColor       =   0
            ListField       =   "beneficiario_deudor"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.Label LblEmpresa 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   420
            Left            =   11040
            TabIndex        =   173
            Top             =   120
            Width           =   765
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Observaciones:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   200
            TabIndex        =   122
            Top             =   3480
            Visible         =   0   'False
            Width           =   1110
            WordWrap        =   -1  'True
         End
         Begin VB.Label lbl_cerrado 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "TRAMITE CERRADO !!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   3480
            TabIndex        =   115
            Top             =   -30
            Width           =   4875
         End
         Begin VB.Label lbl_cite 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Contrato/O.S."
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7755
            TabIndex        =   97
            Top             =   75
            Width           =   1845
         End
         Begin VB.Label txt_venta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "venta_codigo"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   9780
            TabIndex        =   83
            Top             =   315
            Width           =   1125
         End
         Begin VB.Label txt_codigo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "doc_codigo"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5115
            TabIndex        =   82
            Top             =   660
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label txt_campo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "doc_numero"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4395
            TabIndex        =   81
            Top             =   660
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Edificio"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   180
            TabIndex        =   53
            Top             =   755
            Width           =   660
         End
         Begin VB.Label txt_codigo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "solicitud_codigo"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   180
            TabIndex        =   51
            Top             =   320
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cod.Trámite"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   48
            Top             =   75
            Width           =   1110
         End
         Begin VB.Label lbl_campo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad Ejecutora"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1545
            TabIndex        =   47
            Top             =   75
            Width           =   1560
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Venta"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9780
            TabIndex        =   25
            Top             =   75
            Width           =   1005
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7380
            TabIndex        =   24
            Top             =   745
            Width           =   615
         End
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTA"
      ForeColor       =   &H00800000&
      Height          =   5040
      Left            =   15
      TabIndex        =   40
      Top             =   720
      Width           =   6585
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   3960
         TabIndex        =   43
         Top             =   4635
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pendientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1320
         TabIndex        =   42
         Top             =   4635
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   4290
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   7567
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "venta_codigo"
            Caption         =   "Nro.Venta"
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
            DataField       =   "unidad_codigo"
            Caption         =   "Unidad"
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
            DataField       =   "solicitud_codigo"
            Caption         =   "Tramite"
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
            DataField       =   "edif_codigo_corto"
            Caption         =   "Cod.Edificio"
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
         BeginProperty Column04 
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cod.Adm/O.S."
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
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
         BeginProperty Column06 
            DataField       =   "edif_descripcion"
            Caption         =   "Nombre.de.Edificio"
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
         BeginProperty Column07 
            DataField       =   "venta_fecha"
            Caption         =   "Fecha.Venta"
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
         BeginProperty Column08 
            DataField       =   "usr_codigo"
            Caption         =   "Usuario"
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
         BeginProperty Column09 
            DataField       =   "doc_numero"
            Caption         =   "Correl.Doc."
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
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   -1  'True
               ColumnWidth     =   2459.906
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   929.764
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   4560
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   16777215
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.Frame FrmDetalle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE POR EQUIPO y OTROS BB.SS."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1425
      Left            =   2160
      TabIndex        =   12
      Top             =   5820
      Width           =   16455
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "tw_tecnico_venta.frx":11FDA
         Height          =   1140
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   16215
         _ExtentX        =   28601
         _ExtentY        =   2011
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
         FormatLocked    =   -1  'True
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "venta_codigo"
            Caption         =   "Nro.Venta"
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
            DataField       =   "bien_codigo"
            Caption         =   "Codigo.Bien/Eqp"
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
            DataField       =   "concepto_venta"
            Caption         =   "Descripcion y Características del Bien"
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
         BeginProperty Column03 
            DataField       =   "venta_det_cantidad"
            Caption         =   "Cantidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4105
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "venta_precio_unitario_bs"
            Caption         =   "Prec.Unitario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "venta_descuento_bs"
            Caption         =   "Descuento"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "venta_precio_total_bs"
            Caption         =   "Precio Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo.Equipo"
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
         BeginProperty Column08 
            DataField       =   "bien_cantidad_por_empaque"
            Caption         =   "Hrs.X Dia"
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
         BeginProperty Column09 
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
         BeginProperty Column10 
            DataField       =   "observaciones"
            Caption         =   "Descripción.para.el.Cliente"
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
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   4380.095
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   4140.284
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrmCobranza 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PLAN DE CUOTAS PARA CONTROL DE COBRANZA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2145
      Left            =   2160
      TabIndex        =   10
      Top             =   7275
      Width           =   16455
      Begin MSDataGridLib.DataGrid DtgCobro 
         Bindings        =   "tw_tecnico_venta.frx":11FF4
         Height          =   1860
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   16230
         _ExtentX        =   28628
         _ExtentY        =   3281
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
         FormatLocked    =   -1  'True
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "cobranza_prog_codigo"
            Caption         =   "No.Cuota"
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
            DataField       =   "cobranza_fecha_prog"
            Caption         =   "Mes.Programado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "mmm-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cobranza_programada_bs"
            Caption         =   "Monto Programado Bs."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Beneficiario"
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
         BeginProperty Column04 
            DataField       =   "doc_numero"
            Caption         =   "Nro.Doc.Resp."
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
         BeginProperty Column05 
            DataField       =   "cobranza_fecha_conformidad"
            Caption         =   "Fecha.Certif."
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
         BeginProperty Column06 
            DataField       =   "cobranza_observaciones"
            Caption         =   "Concepto de la Cuota"
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
         BeginProperty Column07 
            DataField       =   "cobranza_concepto_plazo"
            Caption         =   "Plazo a Cumplir"
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
         BeginProperty Column08 
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
         BeginProperty Column09 
            DataField       =   "estado_ac"
            Caption         =   "Aviso Cob."
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
         BeginProperty Column10 
            DataField       =   "correl_ac"
            Caption         =   "Nro. Aviso Cob"
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
         BeginProperty Column11 
            DataField       =   "cobranza_programada_dol"
            Caption         =   "Monto a Pagar Dol."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "cobranza_codigo"
            Caption         =   "Cod.Cobranza"
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
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   6284.977
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1184.882
            EndProperty
         EndProperty
      End
   End
   Begin Crystal.CrystalReport CryV01 
      Left            =   0
      Top             =   10680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6720
      Top             =   9960
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2160
      Top             =   9960
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc ado_datos14 
      Height          =   330
      Left            =   11280
      Top             =   10320
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "ado_datos14"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc ado_datos17 
      Height          =   330
      Left            =   9000
      Top             =   10320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "ado_datos17"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_datos11 
      Height          =   330
      Left            =   -120
      Top             =   10320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos11"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_datos16 
      Height          =   330
      Left            =   13560
      Top             =   10320
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos16"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc ado_datos15 
      Height          =   330
      Left            =   6720
      Top             =   10320
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "ado_datos15"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   11280
      Top             =   9960
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_Datos12 
      Height          =   330
      Left            =   2160
      Top             =   10320
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_Datos12"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_datos13 
      Height          =   330
      Left            =   4440
      Top             =   10320
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos13"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   13560
      Top             =   9960
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoAux"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_datos3 
      Height          =   330
      Left            =   4440
      Top             =   9960
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
      Top             =   9960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc ado_datos4A 
      Height          =   330
      Left            =   9000
      Top             =   9960
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "ado_datos4A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Crystal.CrystalReport CryR01 
      Left            =   480
      Top             =   10680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox FraGrabarCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      FillStyle       =   2  'Horizontal Line
      ForeColor       =   &H80000008&
      Height          =   676
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   20400
      TabIndex        =   109
      Top             =   0
      Visible         =   0   'False
      Width           =   20400
      Begin VB.PictureBox BtnEliminar2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1320
         Picture         =   "tw_tecnico_venta.frx":1200E
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   209
         ToolTipText     =   "Borra el ""Detalle del Cronograma por Contrato"""
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnAñadir2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "tw_tecnico_venta.frx":1275A
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   210
         ToolTipText     =   "Genera Nuevo ""Detalle del Cronograma por Contrato"""
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "tw_tecnico_venta.frx":12F19
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   111
         Top             =   0
         Width           =   1455
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         Picture         =   "tw_tecnico_venta.frx":13805
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   110
         Top             =   0
         Width           =   1280
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENTAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   13215
         TabIndex        =   112
         Top             =   180
         Width           =   1005
      End
   End
   Begin Crystal.CrystalReport CryV02 
      Left            =   960
      Top             =   10680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport cry_ac 
      Left            =   2160
      Top             =   10680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport cry_deuda 
      Left            =   2640
      Top             =   10680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryR02 
      Left            =   3120
      Top             =   10680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   4440
      Top             =   10680
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_datos7 
      Height          =   330
      Left            =   6720
      Top             =   10680
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   9000
      Top             =   10680
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   11280
      Top             =   10680
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado_datos9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label LblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label LblUni_descripcion_larga 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   225
      Left            =   3360
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.Label lblUni_codigo 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "tw_tecnico_venta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************
'Ventas
Dim rs_datos As New ADODB.Recordset     'VENTAS
Dim rs_datos1 As New ADODB.Recordset    'UNIDAD EJECUTORA
Dim rs_datos2 As New ADODB.Recordset    'Beneficiario Personas Nat. y Juridicas (menos de CGI)
Dim rs_datos3 As New ADODB.Recordset    'Proyecto de Edificacion
Dim rs_datos4 As New ADODB.Recordset    'Beneficiario Funcionario de CGI (Vendedor, Cobrador, Admin, etc.)
Dim rs_datos6 As New ADODB.Recordset    'Alcance del Contrato
Dim rs_datos7 As New ADODB.Recordset    'Zonas Piloto
Dim rs_datos8 As New ADODB.Recordset    'EMPRESA CGI-CGE
Dim rs_datos9 As New ADODB.Recordset    'TO_Cronograma
Dim rs_datos10 As New ADODB.Recordset   'tc_zona_piloto_edif
Dim rs_datos11 As New ADODB.Recordset
Dim rs_datos12 As New ADODB.Recordset
Dim rs_datos13 As New ADODB.Recordset
Dim rs_datos14 As New ADODB.Recordset   'Ventas_detalle
Dim rs_datos15 As New ADODB.Recordset
Dim rs_datos16 As New ADODB.Recordset   'Ventas cobranzas
Dim rs_datos17 As New ADODB.Recordset
Dim rs_datos18 As New ADODB.Recordset
Dim rs_aviso_cob As New ADODB.Recordset
Dim rs_datos19 As New ADODB.Recordset   'Acumula Cobranzas
'AUXILIARES
'Dim rs_Ventas_lista As New ADODB.Recordset
Dim rs_aux99 As New ADODB.Recordset
Dim rs_aux0 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset
Dim rs_aux9 As New ADODB.Recordset
Dim rs_aux10 As New ADODB.Recordset     'Para Solicitud de Anulacion Factura ao_ventas_cobranzas
Dim rs_aux11 As New ADODB.Recordset
Dim rs_aux12 As New ADODB.Recordset     'Datos Personales Auxiliares gc_beneficiario
Dim rs_aux13 As New ADODB.Recordset
Dim rs_aux14 As New ADODB.Recordset
Dim rs_aux15 As New ADODB.Recordset
Dim rs_aux16 As New ADODB.Recordset     'Cronograma Mensual to_Cronograma_Mensual
Dim rs_aux17 As New ADODB.Recordset
Dim rs_aux18 As New ADODB.Recordset     'Correlativo tc_zona_piloto_edif
Dim rs_aux19 As New ADODB.Recordset
Dim rs_aux20 As New ADODB.Recordset
Dim rs_aux21 As New ADODB.Recordset
Dim rs_aux22 As New ADODB.Recordset

Dim rstdestino As New ADODB.Recordset
Dim rstcorrel_ing As New ADODB.Recordset
'OTROS
'Dim rstdetsalalm As New ADODB.Recordset
Dim RS_BENEF As New ADODB.Recordset
Dim rs_TipoCambio As New ADODB.Recordset
Dim rs_almacen2 As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim rsAuxDetalle As New ADODB.Recordset

'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir As String
'Dim queryinicial As String
Dim queryinicial2, consulta As String
'Almacenes
Dim descri_bien As String
Dim Cant_Alm, VAR_CANT As Integer
Dim correlativo1 As Long
'VARIABLES
Dim marca1 As Variant

Dim swgrabar, swnuevo, deta2, CONT_MED As Integer
Dim correldet2, corrprog As Integer
Dim VAR_PARTIDA, VAR_PROY, correldetalle As Integer
Dim CONT1, CONT2, CONT3, CONT4, VAR_TIPO As Integer
Dim fdia, fmes, fanio, Dias_Mes, TimeD  As Integer
Dim VAR_COBR1, VAR_COBR2, VAR_CONTR, aviso_cob As Integer
Dim VAR_ALMACEN, VAR_SOLTIPO As Integer
Dim VAR_IMPAR, VAR_MESINI2, VAR_ORDEN As Integer
Dim VAR_CANTCRO, VAR_ZONA, VAR_EMPRESA As Integer
Dim VAR_EDIFC As Integer

Dim nroventa, correlv As Long
Dim VAR_CORREL, VAR_CORRELM As Long
Dim VAR_COMPM, VAR_TECPLAN As Long
Dim VAR_CODANT, Var_Comp, VAR_SOL, CANTOT As Long
Dim VAR_IDFAC As Long

Dim Cobrobs, VAR_COBR, VAR_AUX, VAR_AUX2 As Double
Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2, VAR_MBS2, VAR_MDOL2 As Double

Dim VAR_DCORR, VAR_HCORR As String
Dim gestion0, var_literal, VAR_PROY2, VAR_CITE, VAR_CTA As String
Dim VAR_CODTIPO, VAR_ORG, VAR_FTE, VAR_BENEF, VAR_GLOSA, VAR_GLOSA2, VAR_MONEDA As String
Dim VAR_BEND, VAR_EDIFD, VARG_ORGD, VAR_CTAD, VAR_UNID, VAR_DPTO, VAR_DPTOD As String
Dim VAR_COD1, VAR_COD2, VAR_COD3, VAR_COD4 As String
Dim VAR_MED, VAR_MED2 As String
Dim VAR_TIPOV, VAR_VAL As String
Dim VAR_FEC2, MControl, VAR_MES2 As String
Dim VAR_BEN2, VAR_TRANS As String
Dim VAR_DA, VAR_OS As String
Dim VAR_NOMD, VAR_NOMH As String
Dim VAR_UORIGEN, VAR_COBRANZA As String
Dim VAR_CODDOC, VAR_EMISION As String
Dim VAR_EXPOR, VAR_VALD As String
Dim VAR_NEWZ, VAR_ESTADO As String
Dim VAR_PRO, VAR_SUB, VAR_ETAPA, VAR_CLASIF, VAR_DOCS As String

Dim FInicio, FFin, FControl As Date

Public Function Literal(Cadena As String) As String
Dim SW As Integer
Dim sw1 As Integer
Dim swc As Integer
Dim VEC(20) As Long
SW = 0
      '*********PARTE DECIMAL*********
            If Cadena < 0 Then Cadena = Cadena * (-1)
            Cadena = Round(Cadena, 2)
             x = Len(Cadena)
              For k = 1 To x
                  Z = Mid(Cadena, k, 1)
                  If (Z = ".") Or SW = 1 Then
                    d = d + Mid(Cadena, k, 1)
                    SW = 1
                  End If
              Next k
              
              d = Mid(d, 2, Len(d))
              
              'Para la parte decimal del monto
              If d = "00" Or d = "" Then
                 d = d & " 00/100"
              Else
                 If d >= 0 And d <= 9 And Len(d) = 1 Then
                    d = " " & d & "0" & "/100"
                 Else
                    d = " " & d & "/100 "
                 End If
              End If
      '*********PARTE ENTERA*********
 If Cadena <> "" Then
    Cadena = Int(Cadena)
 Else
    MsgBox "No existe monto"
 End If
   s = ""
   Z = ""
   c = 0
   k = 0
   sw1 = 0
   swc = 0
   
   
   x = Len(Cadena)
   For i = 1 To x
       a = Mid(Cadena, i, 1)
       VEC(i) = Mid(Cadena, i, 1)
   Next i
j = x
While j <> 0
k = k + 1
If k <> 8 Then
  If c <> 3 Then
       c = c + 1
      
       If c = 1 And (VEC(j - 1) <> 1 And VEC(j - 1) <> 2) Then
            Select Case VEC(j)
                Case 0: s = " " + s
                Case 1:
                   If sw1 <> 1 Then
                      s = "UNO " + Z + s
                   End If
                   If sw1 = 1 Then
                      s = "UN " + Z + s
                   End If
                   
                Case 2: s = "DOS " + Z + s
                Case 3: s = "TRES " + Z + s
                Case 4: s = "CUATRO " + Z + s
                Case 5: s = "CINCO " + Z + s
                Case 6: s = "SEIS " + Z + s
                Case 7: s = "SIETE " + Z + s
                Case 8: s = "OCHO " + Z + s
                Case 9: s = "NUEVE " + Z + s
          End Select
          
           'If J + 1 <> "" And sw1 <> 1 And VEC(J - 1) <> 0 And VEC(J) <> 0 Then
           If VEC(j - 1) <> 0 And VEC(j) <> 0 Then
                 s = "Y " + s
           End If
           
        End If
        
         If c = 2 And VEC(j) = 1 Then
               swc = 1
                Select Case VEC(j + 1)
                      Case 0: s = "DIEZ " + Z + s
                      Case 1: s = "ONCE " + Z + s
                      Case 2: s = "DOCE " + Z + s
                      Case 3: s = "TRECE " + Z + s
                      Case 4: s = "CATORCE " + Z + s
                      Case 5: s = "QUINCE " + Z + s
                      Case 6: s = "DIECISEIS " + Z + s
                      Case 7: s = "DIECISIETE " + Z + s
                      Case 8: s = "DIECIOCHO " + Z + s
                      Case 9: s = "DIECINUEVE " + Z + s
                End Select
          End If
          
          If c = 2 And VEC(j) = 2 Then
                Select Case VEC(j + 1)
                      Case 0: s = "VEINTE " + Z + s
                      Case 1: s = "VEINTIUNO " + Z + s
                      Case 2: s = "VEINTIDOS " + Z + s
                      Case 3: s = "VEINTITRES " + Z + s
                      Case 4: s = "VEINTICUATRO " + Z + s
                      Case 5: s = "VEINTICINCO " + Z + s
                      Case 6: s = "VEINTISEIS " + Z + s
                      Case 7: s = "VEINTISIETE " + Z + s
                      Case 8: s = "VEINTIOCHO " + Z + s
                      Case 9: s = "VEINTINUEVE " + Z + s
                End Select
          End If
   
        If c = 2 Then
            Select Case VEC(j)
                Case 3: s = "TREINTA " + Z + s
                Case 4: s = "CUARENTA " + Z + s
                Case 5: s = "CINCUENTA " + Z + s
                Case 6: s = "SESENTA " + Z + s
                Case 7: s = "SETENTA " + Z + s
                Case 8: s = "OCHENTA " + Z + s
                Case 9: s = "NOVENTA " + Z + s
            End Select
            
        End If
        
        If c = 3 Then
            Select Case VEC(j)
                Case 1:
                If j = 1 Then
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       s = "CIEN " + Z + s
                    Else
                       s = "CIENTO " + Z + s
                    End If
                Else
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       s = "CIEN " + Z + s
                    Else
                       s = "CIENTO " + Z + s
                    End If
                       'S = "CIENTO " + z + S
                End If
                Case 2: s = "DOSCIENTOS " + Z + s
                Case 3: s = "TRESCIENTOS " + Z + s
                Case 4: s = "CUATROCIENTOS " + Z + s
                Case 5: s = "QUINIENTOS " + Z + s
                Case 6: s = "SEISCIENTOS " + Z + s
                Case 7: s = "SETECIENTOS " + Z + s
                Case 8: s = "OCHOCIENTOS " + Z + s
                Case 9: s = "NOVECIENTOS " + Z + s
            End Select
        End If
   Else
     If j >= 3 Then
            If VEC(j) = 0 And VEC(j - 1) = 0 And VEC(j - 2) = 0 Then
            Else
              s = "MIL " + s
            End If
    Else
              s = "MIL " + s
    End If
        j = j + 1
        c = 0
        sw1 = 1
   End If
 Else
    If VEC(j) <> 1 Then
      s = "MILLONES " + s
    Else
'      If K > 7 Then
      If k <> 8 Then
        s = "MILLONES " + s
      Else
        s = "MILLON " + s
      End If
    End If
      j = j + 1
      c = 0
      sw1 = 1
 End If
   j = j - 1
   
Wend

Literal = s + d
End Function


Private Sub CmdDetalle_Click()
    FrmCobranza.Visible = True
End Sub

'Private Sub adosalalm_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    If pRecordset.EOF Or pRecordset.BOF Then Exit Sub
'        Select Case pRecordset.EditMode
'        Case adEditNone
'            If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'            rstdetsalalm.Open "Select * from ao_detallesalidaalmacen where correlativo_salida = '" & pRecordset("correlativo_salida") & "'", db, adOpenDynamic, adLockOptimistic
'            Set DataGrid2.DataSource = Nothing
'            Set DataGrid2.DataSource = rstdetsalalm
'            DataGrid2.ReBind
'        End Select
'End Sub

Private Sub Adodetallesolicitud_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If (Not adoDetalleSolicitud.Recordset.BOF) And (Not adoDetalleSolicitud.Recordset.EOF) Then
        If Not IsNull(adoDetalleSolicitud.Recordset("correlativo_solicitud")) Then
            txtnosolicitud1.Text = adoDetalleSolicitud.Recordset("correlativo_solicitud")
            txtcorrdet.Text = adoDetalleSolicitud.Recordset("correlativo_detalle")
        Else
            txtnosolicitud1.Text = Ado_datos.Recordset("codigo_solicitud")
            txtcorrdet.Text = " "
            dtccodpar.Text = " "
            dtcdescripar.Text = " "
            txtsolpeso.Text = 0
        End If
    End If
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim descri_bien As String
Dim Cant_Alm As Integer
If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
   If Not IsNull(Ado_datos.Recordset!venta_codigo) Then
        If Ado_datos.Recordset!codigo_empresa = "2" Then
            LblEmpresa.Caption = "CGE"
        Else
            LblEmpresa.Caption = "CGI"
        End If
        If glusuario = "CARIZACA" Or glusuario = "KBETANCOURTH" Or glusuario = "DTORRICO" Or glusuario = "VBELLIDO" Or glusuario = "ADMIN" Or glusuario = "CSALINAS" Or glusuario = "LNAVA" Then
            btnEliminar.Visible = True
        Else
            btnEliminar.Visible = False
        End If
        If buscados = 0 Then
           OptFilGral1.Visible = True
           OptFilGral2.Visible = True
        Else
           OptFilGral1.Visible = False
           OptFilGral2.Visible = False
        End If
        If VAR_UORIGEN = "DNREP" Or VAR_UORIGEN = "DNINS" Or VAR_UORIGEN = "DREPS" Or VAR_UORIGEN = "DREPB" Or VAR_UORIGEN = "DREPC" Or VAR_UORIGEN = "DINSS" Or VAR_UORIGEN = "DINSB" Then
            BtnImprimir1.Visible = True
            BtnImprimir4.Visible = True
        Else
            BtnImprimir1.Visible = False
            BtnImprimir4.Visible = False
        End If
            
        If (Ado_datos.Recordset!estado_codigo = "REG") Then
            BtnAprobar.Visible = True
            BtnModificar.Visible = True
            BtnVer2.Visible = True
'            BtnEliminar.Visible = True                  ' ANULAR TRAMITE
            BtnAñadir.Visible = False   'Cerrar Tramite
            BtnDesAprobar.Visible = False     'ProvisionalOr VAR_UORIGEN = "DREPC"
            lbl_cerrado.Caption = ""
            BtnImprimir2.Visible = True
            BtnVer.Visible = False
            If IsNull(Ado_datos.Recordset("venta_tipo")) Then
                FrmABMDet.Visible = False
                FrmABMDet2.Visible = False
                FrmCobranza.Visible = False
            Else
                FrmABMDet.Visible = True
                FrmABMDet2.Visible = True
                FrmCobranza.Visible = True
            End If
            If VAR_UORIGEN = "DNREP" Or VAR_UORIGEN = "DNINS" Or VAR_UORIGEN = "DREPS" Or VAR_UORIGEN = "DREPB" Or VAR_UORIGEN = "DREPC" Or VAR_UORIGEN = "DINSS" Or VAR_UORIGEN = "DINSB" Then
                FrmABMDet.Visible = True
                BtnImprimir4.Visible = True
                SSTab1.TabEnabled(1) = False
                SSTab1.TabEnabled(3) = True
            Else
                'BtnModDetalle.Visible = False
                BtnImprimir4.Visible = False
                SSTab1.TabEnabled(3) = False
                SSTab1.TabEnabled(1) = True
                FrmEdita.Enabled = False
            End If
        Else
            BtnAprobar.Visible = False
            BtnModificar.Visible = False
            BtnVer2.Visible = False
            Select Case Ado_datos.Recordset!estado_cancelado
                Case "S"
                    lbl_cerrado.Caption = "TRAMITE CERRADO !!"
                    FrmABMDet2.Visible = False
                    BtnAñadir.Visible = False   'Cerrar Tramite
                    BtnDesAprobar.Visible = False     'Provisional
                    FrmABMDet.Visible = False
                    BtnVer.Visible = False
                Case "P"
                    lbl_cerrado.Caption = "TRAMITE PROVISIONAL !!"
                    If glusuario = "ADMIN" Or glusuario = "CARIZACA" Or glusuario = "FDELGADILLO" Or glusuario = "TCASTILLO" Or glusuario = "VBELLIDO" Or glusuario = "SQUISPE" Or glusuario = "JAVIER" Or glusuario = "KBETANCOURTH" Or glusuario = "LNAVA" Or glusuario = "FFLORES" Or glusuario = "MARTEAGA" Or glusuario = "RGIL" Or glusuario = "LMORALES" Or glusuario = "GMORA" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Or glusuario = "ARODRIGUEZ" Then
                        BtnModificar.Visible = True
                        BtnVer2.Visible = True
                        FrmABMDet.Visible = True
                        BtnDesAprobar.Visible = True     'Provisional
                        'SSTab1.TabEnabled(1) = True
                        If VAR_UORIGEN = "DNREP" Or VAR_UORIGEN = "DNINS" Then

                        Else
'                            BtnModDetalle.Visible = False
                        End If
                    Else
                        BtnModificar.Visible = False
                        BtnVer2.Visible = False
                        FrmABMDet.Visible = False
                        BtnDesAprobar.Visible = False 'Provisional
                    End If
                    FrmABMDet2.Visible = True
                    BtnAñadir.Visible = False   'Cerrar Tramite
                    BtnVer.Visible = True
                Case Else
                    BtnAñadir.Visible = True   'Cerrar Tramite
                    If glusuario = "MARTEAGA" Or glusuario = "ADMIN" Or glusuario = "CARIZACA" Or glusuario = "FDELGADILLO" Or glusuario = "VBELLIDO" Or glusuario = "TCASTILLO" Or glusuario = "RVALDIVIEZO" Or glusuario = "SQUISPE" Or glusuario = "KBETANCOURTH" Or glusuario = "LNAVA" Or glusuario = "FFLORES" Or glusuario = "RGIL" Or glusuario = "LMORALES" Or glusuario = "GMORA" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Or glusuario = "ARODRIGUEZ" Then
                        BtnDesAprobar.Visible = True     'Provisional
                    Else
                        BtnDesAprobar.Visible = False     'Provisional
                    End If
                    lbl_cerrado.Caption = ""
                    FrmABMDet2.Visible = True
                    BtnVer.Visible = True
            End Select
            FrmCobranza.Visible = True
            BtnImprimir2.Visible = True
            If (Ado_datos.Recordset!estado_codigo = "ANL" Or Ado_datos.Recordset!estado_codigo = "ERR") Then
                lbl_cerrado.Caption = "TRAMITE ANULADO !!"
                BtnAñadir.Visible = False
                BtnDesAprobar.Visible = False
            End If
        End If
        
        If Dtc_deudor2.Text = "SI" Then
            Dtc_deudor2.backColor = &HFF&
        Else
            Dtc_deudor2.backColor = &H80000010
        End If
        'If Ado_datos.Recordset("beneficiario_codigo") <> "" And Ado_datos.Recordset("beneficiario_codigo") <> "VD" Then
        If Ado_datos.Recordset("beneficiario_codigo") <> "" Then
            Set RS_BENEF = New ADODB.Recordset
            If RS_BENEF.State = 1 Then RS_BENEF.Close
            RS_BENEF.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
            'RS_BENEF.Recordset.Requery
            If RS_BENEF.RecordCount > 0 Then
                If RS_BENEF!beneficiario_deudor = "SI" Then
                    Dtc_deudor2.backColor = &HFF&
                Else
                    Dtc_deudor2.backColor = &H80000010
                End If
            End If
        End If
        Call ABRIR_DETALLE
        If VAR_UORIGEN = "DNMAN" Then
            FrmDetalle.Caption = "BIENES DEL CONTRATO NRO. " + (Ado_datos.Recordset!unidad_codigo_ant)
            FrmCobranza.Caption = "PLAN DE CUOTAS - CONTRATO NRO. " + (Ado_datos.Recordset!unidad_codigo_ant)
            lbl_cite.Caption = "Contrato(Cod.Adm)"
        Else
            FrmDetalle.Caption = "BIENES DE ORDEN DE SERVICIO NRO. " + (Ado_datos.Recordset!unidad_codigo_ant)
            FrmCobranza.Caption = "PLAN DE CUOTAS - ORDEN DE SERVICIO NRO. " + (Ado_datos.Recordset!unidad_codigo_ant)
            lbl_cite.Caption = "Orden Servicio"
        End If
        
        'FrmDetalle.Caption = "BIENES DEL CONTRATO NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
        'FrmCobranza.Caption = "PLAN DE CUOTAS - CONTRATO NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
        
        
        End If
        FrmDetalle.Visible = True
        FrmCobranza.Visible = True
    Else
        FrmABMDet.Visible = False
        FrmABMDet2.Visible = False
        'FrmCabecera.Enabled = True
        FrmDetalle.Visible = False
        FrmCobranza.Visible = False
        If buscados = 0 Then
           OptFilGral1.Visible = True
           OptFilGral2.Visible = True
        Else
           OptFilGral1.Visible = False
           OptFilGral2.Visible = False
        End If
    End If
End Sub

Private Sub ABRIR_DETALLE()
        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "' order by  par_codigo, bien_codigo ", db, adOpenKeyset, adLockOptimistic
        Set Ado_datos14.Recordset = rs_datos14
        Ado_datos14.Recordset.Requery
        If Ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
            'Call AbreAlmacen
        Else
            deta2 = 0
            FrmABMDet2.Visible = False
            FrmCobranza.Visible = False
        End If

        Set rs_datos16 = New ADODB.Recordset
        If rs_datos16.State = 1 Then rs_datos16.Close
        rs_datos16.Open "select * from ao_ventas_cobranza_prog where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        Set Ado_datos16.Recordset = rs_datos16
        Ado_datos16.Recordset.Requery
        If Ado_datos16.Recordset.RecordCount > 0 Then
            FrmCobranza.Visible = True
            
            'BtnImprimir2.Visible = True
            'BtnImprimir3.Visible = True
        Else
            FrmCobranza.Visible = False
            'BtnImprimir2.Visible = False
            'BtnImprimir3.Visible = False
        End If
        'ALCANCE
        Set rs_datos6 = New ADODB.Recordset
        If rs_datos6.State = 1 Then rs_datos6.Close
        rs_datos6.Open "select * from ao_ventas_alcance where venta_codigo= " & Ado_datos.Recordset!venta_codigo & "  order by ORDEN ", db, adOpenKeyset, adLockOptimistic, adCmdText
        'rs_datos6.Open "select * from ao_ventas_alcance where venta_codigo= " & nro_licitacion & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText       'order by ORDEN
        Set Ado_datos6.Recordset = rs_datos6
        Set DtgAlcance.DataSource = Ado_datos6.Recordset
        If Ado_datos6.Recordset.RecordCount > 0 Then
            DtgAlcance.Visible = True
        Else
            DtgAlcance.Visible = False
        End If
End Sub

Private Sub AbreAlmacen()
'    Set rs_datos13 = New ADODB.Recordset
'    If rs_datos13.State = 1 Then rs_datos13.Close
'    rs_datos13.Open "select * from Av_almacen_detalle where bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_datos13.Recordset = rs_datos13
'    Ado_datos13.Refresh

End Sub

Private Sub Ado_datos16_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 If (Not Ado_datos16.Recordset.BOF) And (Not Ado_datos16.Recordset.EOF) Then
    If Not IsNull(Ado_datos16.Recordset("venta_codigo")) Then
        'BtnModDetalle2.Visible = False
        If (Ado_datos16.Recordset("estado_codigo") = "REG") Then
'            If (Ado_datos.Recordset("estado_codigo") = "APR") Then
'                BtnAprobar2.Visible = False
'            Else
'                BtnAprobar2.Visible = True
'            End If
            BtnImprimir2.Visible = True
            BtnAprobar2.Visible = True
'            BtnAnlDetalle2.Visible = True
            BtnModDetalle2.Visible = True
        End If
        If (Ado_datos16.Recordset("estado_codigo") = "APR") Then
            BtnImprimir2.Visible = True
            BtnAprobar2.Visible = False
'            BtnAnlDetalle2.Visible = False
            BtnModDetalle2.Visible = False
'            Command4.Visible = True
        End If
        If (Ado_datos16.Recordset("estado_codigo") = "ANL") Then
            'BtnImprimir2.Visible = False
'            BtnAnlDetalle2.Visible = False
            BtnModDetalle2.Visible = False
            BtnAprobar2.Visible = False
'            Command4.Visible = False
        End If
    Else
        BtnAprobar2.Visible = False
        BtnImprimir2.Visible = False
'        BtnAnlDetalle2.Visible = False
        BtnModDetalle2.Visible = False
'        Command4.Visible = False
    End If
 Else
    BtnAprobar2.Visible = False
    BtnImprimir2.Visible = False
'    BtnAnlDetalle2.Visible = False
    BtnModDetalle2.Visible = False
 End If
End Sub



Private Sub BtnAddDetalle1_Click()
    NumComp = Ado_datos.Recordset!venta_codigo
    gestion0 = Ado_datos.Recordset!ges_gestion
    sino = MsgBox("Elije SI, para borrar los Registros del Alcance y Volver a Cargarlos ? . Elije NO, para Cancelar... ", vbYesNo, "Confirmando")
    If sino = vbYes Then
        db.Execute "DELETE ao_ventas_alcance WHERE venta_codigo = " & NumComp & " "
    Else
        Exit Sub
    End If
'    sino = MsgBox("El Contrato incluye... SERVICIO DE DESMONTAJE ? ...", vbYesNo, "Confirmando")
    Set rs_aux19 = New ADODB.Recordset
    If rs_aux19.State = 1 Then rs_aux19.Close
    If VAR_UORIGEN = "DNREP" Or VAR_UORIGEN = "DREPS" Or VAR_UORIGEN = "DREPB" Or VAR_UORIGEN = "DREPC" Then
       rs_aux19.Open "Select * from gc_tipo_solicitud where solicitud_tipo = '15' OR solicitud_tipo = '16' OR solicitud_tipo = '17' OR solicitud_tipo = '18' OR solicitud_tipo = '7'  order by ORDEN ", db, adOpenStatic
    Else
       'rs_aux19.Open "Select * from gc_tipo_solicitud WHERE (solicitud_num = '90') AND (solicitud_tipo <> '20') AND (solicitud_tipo <> '6') AND (solicitud_tipo <> '5') AND (solicitud_tipo <> '2') AND (solicitud_tipo <> '3') OR (solicitud_tipo = '15') order by ORDEN ", db, adOpenStatic
       rs_aux19.Open "Select * from gc_tipo_solicitud where solicitud_tipo = '15' OR solicitud_tipo = '16' OR solicitud_tipo = '17' OR solicitud_tipo = '18' OR solicitud_tipo = '4'  order by ORDEN ", db, adOpenStatic
    End If
    'Set Ado_datos1.Recordset = rs_aux19
    If rs_aux19.RecordCount > 0 Then
        'ao_ventas_alcance
        rs_aux19.MoveFirst
        While Not rs_aux19.EOF
            db.Execute "INSERT INTO ao_ventas_alcance (ges_gestion, venta_codigo, solicitud_tipo, venta_codigo_new, solicitud_tipo_descripcion, unidad_codigo_tec, venta_tiempo_dias, fecha_inicio_alcance, fecha_fin_alcance, fecha_inicio_real, fecha_fin_real, " & _
            " doc_codigo , correl_doc, estado_codigo, usr_codigo, fecha_registro, hora_registro, estado_acta, estado_mantenimiento, orden) " & _
            " VALUES ('" & gestion0 & "', " & NumComp & ", " & rs_aux19!solicitud_tipo & ", '0', '" & rs_aux19!solicitud_tipo_descripcion & "' , '" & rs_aux19!unidad_codigo & "','0', '01/01/1900' , '01/01/1900', '01/01/1900' , '01/01/1900', " & _
            " 'R-321', '0', 'REG', '" & glusuario & "', '" & Date & "', '0', 'REG', 'REG', " & rs_aux19!Orden & ") "
            
            rs_aux19.MoveNext
        Wend
    Else
        
    End If
    'Call ABRIR_TABLA_DET
    Call ABRIR_DETALLE
End Sub

Private Sub BtnAñadir_Click()
'CERRAR UN TRAMITE = S
If glusuario = "CARIZACA" Or glusuario = "VBELLIDO" Or glusuario = "JAVIER" Or glusuario = "JSAAVEDRA" Or glusuario = "ADMIN" Or glusuario = "KBETANCOURTH" Or glusuario = "LNAVA" Or glusuario = "FFLORES" Or glusuario = "MARTEAGA" Or glusuario = "ULEDEZMA" Or glusuario = "CSALINAS" Or glusuario = "ARODRIGUEZ" Or glusuario = "RGIL" Or glusuario = "LMORALES" Or glusuario = "GMORA" Or glusuario = "ASANTIVAÑEZ" Then
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_cancelado = "N" And Ado_datos.Recordset!estado_codigo = "APR" Then
      sino = MsgBox("Esta seguro de CERRAR EL TRAMITE, ya no podrá realizar modificaciones... ", vbYesNo, "Confirmando")
      If sino = vbYes Then
          db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_cancelado = 'S' Where ao_ventas_cabecera.venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  "
          db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'ANL' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " and estado_codigo = 'REG' "
          marca1 = Ado_datos.Recordset.Bookmark
'          'Ado_datos.Recordset.Requery
'          'Ado_datos.Refresh
          'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
            Dim iResult As Variant, i%, Y%
            Dim co As New ADODB.Command
            CryR02.ReportFileName = App.Path & "\reportes\Tecnico\tr_certificado_cumplim_contrato.rpt"
            'CryR02.WindowShowRefreshBtn = True
            CryR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
            CryR02.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo

            iResult = CryR02.PrintReport
            If iResult <> 0 Then MsgBox CryR02.LastErrorNumber & " : " & CryR02.LastErrorString, vbCritical, "Error de impresión"
          'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
          If Ado_datos.Recordset!estado_codigo = "REG" Then
            Call OptFilGral1_Click
          Else
            Call OptFilGral2_Click
          End If
          Ado_datos.Recordset.Move marca1 - 1
      End If
    Else
      MsgBox "NO se puede procesar el TRAMITE ya fue CERRADO...", , "Atencion"
    End If
  Else
    MsgBox "NO se puede procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
Else
MsgBox "No tiene acceso a esta opcion", vbInformation, "Atención!!!"
End If
End Sub

Private Sub BtnAñadir2_Click()
'' If Ado_datos.Recordset!estado_codigo = "REG" Then
''  VAR_VALD = "OK"
''  Call valida_campos
''  If VAR_VALD = "ERR" Then
''      Exit Sub
''  Else
'        'WWWWW GENERA CRONOGRAMA DIARIO WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
''        FrmABMDet2.Enabled = False
''        FrmABMDet.Enabled = False
''        fraOpciones.Enabled = False
'        'Screen.MousePointer = vbHourglass
'
'                ' INI VARIABLES DE LA VENTA CABECERA WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
''                VAR_COD1 = Ado_datos.Recordset!unidad_codigo
''                VAR_TIPOV = Ado_datos.Recordset!venta_tipo
''                NumComp = Ado_datos.Recordset!venta_codigo
''                VAR_ZONA = Ado_datos.Recordset!zpiloto_codigo
''                VAR_COD4 = Ado_datos.Recordset!unidad_codigo
''               VAR_GLOSA = Ado_datos.Recordset!venta_descripcion
''               VAR_DOL2 = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
''               VAR_BS2 = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
''               VAR_UNIMED = Ado_datos.Recordset!unimed_codigo
''               '
''               VAR_SOL = Ado_datos.Recordset!solicitud_codigo
''               VAR_MED = Ado_datos.Recordset!unimed_codigo_tec
''               VAR_MED2 = Ado_datos.Recordset!unimed_codigo_cobr
''               FInicio = Ado_datos.Recordset!venta_fecha_inicio
''               FFin = Ado_datos.Recordset!venta_fecha_fin
''               TimeD = Ado_datos.Recordset!venta_plazo_dias_calendario
''               CANTOT = Ado_datos.Recordset!venta_cantidad_total
''               VAR_GLOSA2 = Ado_datos.Recordset!venta_descripcion
''               VAR_PROY2 = Ado_datos.Recordset!edif_codigo
''               VAR_CITE = Ado_datos.Recordset!unidad_codigo_ant         'OS - 36AO - 36NB - 36NO...
''               VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
''               VAR_BEND = dtc_desc2.Text
''               VAR_EDIFD = dtc_desc3.Text
''               VAR_UNID = dtc_desc1.Text
''               'VAR_DPTO = Left(VAR_PROY2, 1)
''               VAR_DPTO = Ado_datos.Recordset!depto_codigo
''               VARG_ORGD = ""
''               VAR_CTAD = ""
'            ' FIN VARIABLES DE LA VENTA CABECERA WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'
'        NumComp = Ado_datos.Recordset!venta_codigo
'        gestion0 = Ado_datos.Recordset!ges_gestion
'        VAR_PROY2 = Ado_datos.Recordset!edif_codigo
'        VAR_EMPRESA = Ado_datos.Recordset!codigo_empresa
'        '---------- de to_cronograma ANTES
'        FInicio = Format(Ado_datos.Recordset!fecha_inicio_tec, "dd/mm/yyyy")
'        FFin = Format(Ado_datos.Recordset!fecha_fin_tec, "dd/mm/yyyy")
'        CANTOT = IIf(IsNull(Ado_datos.Recordset!cantidad_periodos_tec), 12, Ado_datos.Recordset!cantidad_periodos_tec)
'        VAR_MED = IIf(IsNull(Ado_datos.Recordset!unimed_codigo_tec), "MES", Ado_datos.Recordset!unimed_codigo_tec)
'        VAR_ZONA = Ado_datos.Recordset!zpiloto_codigo
'        VAR_UNITEC = Ado_datos.Recordset!unidad_codigo                  'unidad_codigo_tec
'        VAR_TECCOD = Ado_datos.Recordset!solicitud_codigo               'tec_plan_codigo
'
'        Set rs_aux0 = New ADODB.Recordset
'        If rs_aux0.State = 1 Then rs_aux0.Close
'        rs_aux0.Open "Select * from gc_edificaciones WHERE edif_codigo = '" & VAR_PROY2 & "'   ", db, adOpenStatic
'        If rs_aux0.RecordCount > 0 Then
'            VAR_EDIF = Ado_datos.Recordset!edif_descripcion                      'RTrim(dtc_desc3.Text)          'edif_descripcion
'        End If
'        VAR_LUN = "SI"                                                  'Ado_datos.Recordset!lunes_cambia
'        VAR_PRIM = "SI"                                                 'Ado_datos.Recordset!primero_mes
'
'        VAR_EMES = "Error: No se encontró el Mes de Inicio del Cronograma, verifique y vuelva a intentar..."
'        'dtc_codigo5.Text = "0"
'        'to_cronograma_vs_ventas
''      Set rs_datos6 = New ADODB.Recordset
''      If rs_datos6.State = 1 Then rs_datos6.Close
''      'rs_datos6.Open "Select * from to_cronograma WHERE estado_detalle = 'APR' AND unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   ", db, adOpenStatic
''      rs_datos6.Open "Select * from to_cronograma_vs_ventas WHERE fmes_plan = 'APR' AND correl_prog = '" & VAR_UNITEC & "' and bien_codigo = " & VAR_TECCOD & "   ", db, adOpenStatic
''      If rs_datos6.RecordCount > 0 Then
''           MsgBox "El Cronograma ya existe, verifique y vuelva a intentar ...", vbExclamation, "Validación de Registro"
''           Frame2.Visible = False
''           ProgressBar1.Visible = False
''           Exit Sub
''      Else
'        ' estado_activo = 'ANL'
'        ' jalar ORDEN de tc_zona_piloto_edif
'        Set rs_datos6 = New ADODB.Recordset
'        If rs_datos6.State = 1 Then rs_datos6.Close
'        rs_datos6.Open "Select * from tc_zona_piloto_edif WHERE edif_codigo = '" & VAR_PROY2 & "'    ", db, adOpenStatic
'        If rs_datos6.RecordCount > 0 Then
'            DIA_ORDEN = rs_datos6!zona_edif_orden
'        Else
'            Set rs_aux18 = New ADODB.Recordset
'            If rs_aux18.State = 1 Then rs_aux18.Close
'            rs_aux18.Open "Select ISNULL(max(zona_edif_orden),0) as Orden from tc_zona_piloto_edif where zpiloto_codigo = " & VAR_ZONA & " ", db, adOpenKeyset, adLockOptimistic
'            If rs_aux18.RecordCount > 0 Then
'                VAR_ORDEN = IIf(IsNull(rs_aux18!Orden), 1, rs_aux18!Orden + 1)
'            Else
'                VAR_ORDEN = 1
'            End If
'
'           db.Execute "INSERT INTO tc_zona_piloto_edif (zpiloto_codigo, edif_codigo, ges_gestion, zona_edif_orden, zona_codigo, beneficiario_codigo, beneficiario_codigo_rep, beneficiario_codigo_cobr, zorden_cambio, mes_par_impar, observaciones, " & _
'                      " estado_codigo , estado_activo, fecha_registro, usr_codigo, unimed_codigo, codigo_empresa, solicitud_tipo) " & _
'                      " VALUES (" & VAR_ZONA & ", '" & VAR_PROY2 & "', '" & gestion0 & "',      " & VAR_ORDEN & ",       '0',            '0',                    '0',                    '0',                    '0',            '1',        '',  " & _
'                      " 'REG',              'APR', '" & Date & "', '" & glusuario & "', '" & VAR_MED & "', " & VAR_EMPRESA & ", " & VAR_TIPO & ")"
'            DIA_ORDEN = "1"
'        End If
'        'DIA_ORDEN = Ado_datos.Recordset!zona_edif_orden
'        MControl = Ado_datos.Recordset!mes_inicio_crono_tec                     'mes_inicio_crono
'        'CALL CRONO_MANT
'        If IsNull(Ado_datos.Recordset!mes_inicio_crono_nro) Then
'            VAR_MESINI2 = Month(FInicio)
'            db.Execute "update ao_ventas_cabecera set mes_inicio_crono_nro = " & VAR_MESINI2 & " WHERE venta_codigo = " & NumComp & "  "
'        Else
'            VAR_MESINI2 = Ado_datos.Recordset!mes_inicio_crono_nro
'        End If
'        Select Case VAR_MED
'           Case "MES"
'               UMED_NRO = 1
'           Case "BMES"
'               UMED_NRO = 2
'           Case "TMES"
'               UMED_NRO = 3
'           Case "CMES"
'               UMED_NRO = 4
'           Case "5MES"
'               UMED_NRO = 5
'           Case "SMES"
'               UMED_NRO = 6
'           Case "7MES"
'               UMED_NRO = 7
'           Case "8MES"
'               UMED_NRO = 8
'           Case "9MES"
'               UMED_NRO = 9
'           Case "10MES"
'               UMED_NRO = 10
'           Case "11MES"
'               UMED_NRO = 11
'           Case "ANUAL"
'               UMED_NRO = 12
'       End Select
'       'UMED_NRO = Ado_datos.Recordset!unimed_codigo_nro                        ' Fijo MES=1, BMES=2, TMES=3
'        FControl = FInicio
'        CONT4 = 0
'        VAR_CONT = 1
'        VAR_MES = Month(FControl)
'        UMED_NRO2 = VAR_MESINI2      'UMED_NRO
'        Frame2.Visible = True
'        'ProgressBar1.Visible = True
''        With ProgressBar1
''            .Max = CANTOT     'rs_datos6.RecordCount
''            .Min = 0
''            .Value = 0
''        End With
'        gestion0 = Year(FControl)
'        While CANTOT >= VAR_CONT And FFin >= FControl   'UNIMED veces (12, 24, etc.)
'            If UMED_NRO2 > 12 And gestion0 <> Year(FControl) Then
'                UMED_NRO2 = 1 * UMED_NRO
'                'gestion0 = Year(FControl)
'             End If
'          gestion0 = Year(FControl)
'
'          CONT3 = 0
'          If VAR_MES = UMED_NRO2 Then
'             Set rs_aux1 = New ADODB.Recordset
'             'rs_aux1.Open "select * from to_cronograma_detalle where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   ", db, adOpenKeyset, adLockBatchOptimistic
'             rs_aux1.Open "select * from ao_ventas_detalle where venta_codigo = " & NumComp & " and par_codigo = '43340'   ", db, adOpenKeyset, adLockBatchOptimistic
'             If rs_aux1.RecordCount > 0 Then
'                 ' De acuerdo a la cantidad de equipos
'                 'var_cod5 = IIf(IsNull(rs_aux1!bien_cantidad_por_empaque), 2, rs_aux1!bien_cantidad_por_empaque) / 2
'                 'var_cod5 = IIf(IsNull(rs_aux1!bien_cantidad_por_empaque), 2, rs_aux1!bien_cantidad_por_empaque)
'                 var_cod5 = rs_aux1.RecordCount
'                 rs_aux1.MoveFirst
'                 While Not rs_aux1.EOF
'                     Set rs_aux2 = New ADODB.Recordset
'                     If rs_aux2.State = 1 Then rs_aux2.Close
'                     'rs_aux2.Open "select * from to_cronograma_mensual where ges_gestion = '" & gestion0 & "' and fmes_correl = " & VAR_MES & " and zpiloto_codigo = " & VAR_ZONA & "  and unidad_codigo_tec = '" & VAR_UNITEC & "'   ", db, adOpenKeyset, adLockOptimistic
'                     rs_aux2.Open "select * from to_cronograma_mensual where ges_gestion = '" & gestion0 & "' and fmes_correl = " & VAR_MES & " and zpiloto_codigo = " & VAR_ZONA & "    ", db, adOpenKeyset, adLockOptimistic
'                     If rs_aux2.RecordCount > 0 Then
'                         VAR_AUX2 = rs_aux2!fmes_plan
'                         VAR_COD0 = 0
'                         'UMED_NRO2 = 0
'                         Set rs_aux3 = New ADODB.Recordset
'                         If rs_aux3.State = 1 Then rs_aux3.Close
'                         rs_aux3.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & " AND bien_codigo = ''  ", db, adOpenKeyset, adLockBatchOptimistic
'                         If rs_aux3.RecordCount > 0 Then
'                             rs_aux3.MoveFirst
'                             'While Not rs_aux3.EOF
'                             While VAR_COD0 < var_cod5
'                                'If cmb_dia.Text = "AUTOMATICO" And dtc_codigo5.Text = "0" Then
'                                If rs_aux3!dia_nombre = "AUTOMATICO" And rs_aux3!horario_codigo = "0" Then
'                                   Set rs_aux4 = New ADODB.Recordset
'                                   If rs_aux4.State = 1 Then rs_aux4.Close
'                                   rs_aux4.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & " AND bien_codigo <> '' AND estado_activo = 'REG'  ", db, adOpenKeyset, adLockBatchOptimistic
'                                   If rs_aux4.RecordCount > 0 Then
'                                    If VAR_COD0 < var_cod5 And rs_aux3!estado_activo = "REG" Then
'                                       db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux1!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                       db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
'                                       db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                       db.Execute "update to_cronograma_diario set nro_total_horas = " & var_cod5 & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                       'VAR_COD0 = VAR_COD0 + 1
'                                       VAR_COD0 = VAR_COD0 + var_cod5
'                                       CONT3 = 1
'                                       'If VAR_MES Then
'                                       VAR_EMES = "NADA"
'                                       'End If
''                                       If VAR_LUN = "SI" Or VAR_PRIM = "SI" Then
''                                          'TODOS LOS LUNES O EL 1RO. DE CADA MES
''                                          If (rs_aux3!dia_nombre = "LUNES" Or rs_aux3!dia_correl = "1") And rs_aux3!hora_ingreso = "08:00" Then
''                                             rs_aux3.MoveNext
''                                             db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux1!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                                             db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                                             db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                                             VAR_COD0 = VAR_COD0 + 1
''                                             CONT3 = 1
''                                          End If
''                                       End If
'                                       'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
'                                       db.Execute "Update ao_ventas_cabecera Set estado_crono = 'APR' Where venta_codigo = " & NumComp & "  "
'                                    End If
'                                   Else
'                                        MsgBox "Ya no existen horarios laborales LIBRES, para la gestion: " & gestion0 & ", el Mes: " & VAR_MES & " y la Zona: " & VAR_ZONA, vbInformation, "Información"
'                                        rs_aux3.MoveLast
'                                   End If
'                                Else
''                                   If cmb_dia.Text = rs_aux3!dia_nombre And dtc_codigo5.Text = "0" Then
'    '                                         If rs_aux3!dia_nombre = "SÁBADO" Or rs_aux3!dia_nombre = "DOMINGO" Or rs_aux3!estado_activo = "ANL" Then
'    '                                            db.Execute "update to_cronograma_diario set observaciones = 'DIA NO LABORABLE' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'    '                                            db.Execute "update to_cronograma_diario set estado_activo = 'ANL' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'    '                                         Else
'                                    'var_cod5 = rs_aux1.RecordCount
'                                     If VAR_COD0 < var_cod5 Then     'And rs_aux3!estado_activo = "REG"
'                                         db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux1!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "'   WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                         db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
'                                         db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                         VAR_COD0 = VAR_COD0 + 1
'                                         CONT3 = 1
'                                         'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
'                                         db.Execute "Update ao_ventas_cabecera Set estado_crono = 'APR' Where venta_codigo = " & NumComp & "  "
'                                         VAR_EMES = "NADA"
'                                     End If
''                                   End If
''                                   If dtc_codigo5.Text = rs_aux3!horario_codigo Then
''                                     If VAR_COD0 < var_cod5 Then     'And rs_aux3!estado_activo = "REG"
''                                         db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux1!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "'   WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                                         db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
''                                         db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                                         VAR_COD0 = VAR_COD0 + 1
''                                         CONT3 = 1
''                                         'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
''                                         db.Execute "Update ao_venta_cabecera Set estado_crono = 'APR' Where venta_codigo = " & NumComp & "  "
''                                     End If
''                                   End If
'                                End If
'                                rs_aux3.MoveNext
'                                'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
'                             Wend
'                         End If
'                     End If
'                     rs_aux1.MoveNext
'                 Wend
'             End If
'             VAR_CONT = VAR_CONT + 1
'             UMED_NRO2 = UMED_NRO2 + UMED_NRO
'             'ProgressBar1.Value = ProgressBar1.Value + 1
'          Else
'            'VAR_EMES = "Error: No se encontró el Mes de Inicio del Cronograma, verifique u vuelva a intentar..."
'          End If
'             Select Case VAR_MES
'                 Case 2
'                     If gestion0 = "2016" Or gestion0 = "2020" Or gestion0 = "2024" Or gestion0 = "2028" Then
'                         Dias_Mes = 29
'                     Else
'                         Dias_Mes = 28
'                     End If
'                 Case 1, 3, 5, 7, 8, 10, 12
'                     Dias_Mes = 31
'                 Case 4, 6, 9, 11
'                     Dias_Mes = 30
'             End Select
'             'rs_aux2!cobranza_fecha_prog = FControl
'             'rs_aux2!cobranza_fecha_cobro = FControl + 10
'             FControl = CDate(FControl) + Dias_Mes
'             VAR_MES = Month(FControl)
'             Select Case VAR_MED
'                Case "MES"    'MENSUAL
''                    UMED_NRO2 = VAR_CONT
''                    If (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) Then
''                        'UMED_NRO2 = (VAR_MES * UMED_NRO) - 1
''                        UMED_NRO2 = VAR_CONT
''                    Else
''                        UMED_NRO2 = VAR_MES * UMED_NRO
''                    End If
'                Case "BMES"    'BIMESTRAL
''                    If (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) Then
''                        UMED_NRO2 = (VAR_CONT * UMED_NRO) - 1
''                    Else
''                        UMED_NRO2 = VAR_CONT * UMED_NRO
''                    End If
'                Case "TMES"    'TRIMESTRAL
'                    'UMED_NRO2 = (UMED_NRO2 + UMED_NRO)
''                    If (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) Then
''                        UMED_NRO2 = (VAR_CONT * UMED_NRO) '- 2
''                        'UMED_NRO2 = (UMED_NRO2 + VAR_MESINI2)
''                    Else
''                        UMED_NRO2 = (VAR_CONT * UMED_NRO) - 1
''                    End If
''                    'UMED_NRO2 = 3
''                    If VAR_MES = UMED_NRO2 Then
''                        UMED_NRO2 = UMED_NRO2 + VAR_MESINI2
''                    End If
''                    UMED_NRO2 = VAR_CONT * UMED_NRO
'                Case "CMES"    'CUATRIMESTRAL
'                Case "QMES"    'CADA 5 MESES
'                Case "SMES"    'SEMESTRAL
'                    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'                Case "ANUAL"    'ANUAL
'             End Select
''             If VAR_MED = "TMES" And CONT3 = 1 Then
''                UMED_NRO2 = (VAR_CONT * UMED_NRO) - 2
''                VAR_CONT = VAR_CONT + 1
''             End If
''                If CONT3 = 1 And VAR_MED = "MES" And (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) And (UMED_NRO = 2) Then
''                    UMED_NRO2 = (VAR_MES * UMED_NRO) - 1
''                Else
''                    UMED_NRO2 = VAR_MES * UMED_NRO
''                End If
''                'If CONT3 = 1 And VAR_MED = "BMES" Then
''                If VAR_MED = "BMES" Then
''                    If (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) And (UMED_NRO = 2) Then
''                        UMED_NRO2 = (VAR_CONT * UMED_NRO) - 1
''                        VAR_CONT = VAR_CONT + 1
''                    Else
''                        UMED_NRO2 = VAR_CONT * UMED_NRO
''                        VAR_CONT = VAR_CONT + 1
''                    End If
''                End If
''             'End If
'        Wend
'
'        FrmABMDet2.Enabled = True
'        FrmABMDet.Enabled = True
'        fraOpciones.Enabled = True
''        Screen.MousePointer = vbDefault
'        If VAR_EMES = "NADA" Then
'            MsgBox "El Cronograma fue creado Satisfactoriamente ...", vbInformation, "Información"
''            ProgressBar1.Visible = False
''            Frame2.Visible = False
'        Else
'            MsgBox VAR_EMES, vbInformation, "Información"
'        End If
'        'Call ABRIR_DETALLE
''      End If
''      ProgressBar1.Visible = False
''      Frame2.Visible = False
'      'WWWWW GENERA CRONOGRAMA DIARIO (FIN)
''  End If
'' Else
''        MsgBox "NO se puede generar un NUEVO CRONOGRAMA, en un Registro APROBADO o ANULADO !! ", vbExclamation, "Atención!"
'' End If
'
End Sub

Private Sub CRONO_MTTO()
'    Set rs_aux99 = New ADODB.Recordset
'    If rs_aux99.State = 1 Then rs_aux99.Close
'    rs_aux99.Open "Select * from AV_VENTAS_MTTO_2023", db, adOpenStatic
'    If rs_aux99.RecordCount > 0 Then
'        rs_aux99.MoveFirst
'        While Not rs_aux99.EOF
'               VAR_COD4 = rs_aux99!unidad_codigo
'               VAR_GLOSA = rs_aux99!venta_descripcion
'               VAR_DOL2 = Round(rs_aux99("venta_monto_total_dol"), 2)
'               VAR_BS2 = Round(rs_aux99("venta_monto_total_bs"), 2)
'               VAR_UNIMED = rs_aux99!unimed_codigo_tec
'               'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'
'               VAR_SOL = rs_aux99!solicitud_codigo
'               VAR_MED = rs_aux99!unimed_codigo_tec
'               VAR_MED2 = rs_aux99!unimed_codigo_cobr
'               FInicio = rs_aux99!venta_fecha_inicio
'               FFin = rs_aux99!venta_fecha_fin
'               TimeD = rs_aux99!venta_plazo_dias_calendario
'               CANTOT = rs_aux99!venta_cantidad_total
'               VAR_GLOSA2 = rs_aux99!venta_descripcion
'               VAR_PROY2 = rs_aux99!edif_codigo
'               VAR_CITE = rs_aux99!unidad_codigo_ant         'OS - 36AO - 36NB - 36NO...
'               VAR_BENEF = rs_aux99!beneficiario_codigo
'               VAR_BEND = "0" 'dtc_desc2.Text
'               VAR_EDIFD = "0" 'dtc_desc3.Text
'               VAR_UNID = "0" 'dtc_desc1.Text
'               NumComp = rs_aux99!venta_codigo

    Set rs_aux0 = New ADODB.Recordset
    If rs_aux0.State = 1 Then rs_aux0.Close
    rs_aux0.Open "Select * from gc_edificaciones WHERE edif_codigo = '" & VAR_PROY2 & "'   ", db, adOpenStatic
    If rs_aux0.RecordCount > 0 Then
        VAR_EDIF = rs_aux0!edif_descripcion                      'RTrim(dtc_desc3.Text)          'edif_descripcion
    End If
    VAR_LUN = "SI"                                                  'Ado_datos.Recordset!lunes_cambia
    VAR_PRIM = "SI"                                                 'Ado_datos.Recordset!primero_mes
    
    'VAR_EMES = "Error: No se encontró el Mes de Inicio del Cronograma, verifique y vuelva a intentar..."
    ' jalar ORDEN de tc_zona_piloto_edif
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from tc_zona_piloto_edif WHERE edif_codigo = '" & VAR_PROY2 & "'    ", db, adOpenStatic
    If rs_datos6.RecordCount > 0 Then
        DIA_ORDEN = rs_datos6!zona_edif_orden
    Else
        Set rs_aux18 = New ADODB.Recordset
        If rs_aux18.State = 1 Then rs_aux18.Close
        rs_aux18.Open "Select ISNULL(max(zona_edif_orden),0) as Orden from tc_zona_piloto_edif where zpiloto_codigo = " & VAR_ZONA & " ", db, adOpenKeyset, adLockOptimistic
        If rs_aux18.RecordCount > 0 Then
            VAR_ORDEN = IIf(IsNull(rs_aux18!Orden), 1, rs_aux18!Orden + 1)
        Else
            VAR_ORDEN = 1
        End If
    
       db.Execute "INSERT INTO tc_zona_piloto_edif (zpiloto_codigo, edif_codigo, ges_gestion, zona_edif_orden, zona_codigo, beneficiario_codigo, beneficiario_codigo_rep, beneficiario_codigo_cobr, zorden_cambio, mes_par_impar, observaciones, " & _
                  " estado_codigo , estado_activo, fecha_registro, usr_codigo, unimed_codigo, codigo_empresa, solicitud_tipo) " & _
                  " VALUES (" & VAR_ZONA & ", '" & VAR_PROY2 & "', '" & gestion0 & "',      " & VAR_ORDEN & ",       '0',            '0',                    '0',                    '0',                    '0',            '1',        '',  " & _
                  " 'REG',              'APR', '" & Date & "', '" & glusuario & "', '" & VAR_MED & "', " & VAR_EMPRESA & ", " & VAR_TIPO & ")"
        DIA_ORDEN = "1"
    End If
    'DIA_ORDEN = Ado_datos.Recordset!zona_edif_orden
    MControl = Ado_datos.Recordset!mes_inicio_crono_tec                     'mes_inicio_crono

    Set rs_aux1 = New ADODB.Recordset
    'rs_aux1.Open "select * from ao_ventas_detalle where venta_codigo = " & NumComp & " and par_codigo = '43340'   ", db, adOpenKeyset, adLockBatchOptimistic
    rs_aux1.Open "select * from ao_ventas_cobranza_prog where venta_codigo = " & NumComp & "   ", db, adOpenKeyset, adLockBatchOptimistic
    If rs_aux1.RecordCount > 0 Then
        var_cod5 = rs_aux1.RecordCount
        rs_aux1.MoveFirst
        While Not rs_aux1.EOF
            VAR_AUX2 = rs_aux1!fmes_plan
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            'rs_aux2.Open "select * from to_cronograma_mensual where ges_gestion = '" & gestion0 & "' and fmes_correl = " & VAR_MES & " and zpiloto_codigo = " & VAR_ZONA & "    ", db, adOpenKeyset, adLockOptimistic
            rs_aux2.Open "select * from ao_ventas_detalle where venta_codigo = " & NumComp & " and par_codigo = '43340'   ", db, adOpenKeyset, adLockBatchOptimistic
            If rs_aux2.RecordCount > 0 Then
                rs_aux2.MoveFirst
                While Not rs_aux2.EOF
                    'VERIFICA SI EXITE EQUIPO EN ESTE MES
                    Set rs_aux21 = New ADODB.Recordset
                    If rs_aux21.State = 1 Then rs_aux21.Close
                    rs_aux21.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  ", db, adOpenKeyset, adLockBatchOptimistic
                    If rs_aux21.RecordCount > 0 Then
                        db.Execute "update to_cronograma_diario set unidad_codigo_tec = '" & VAR_COD4 & "',  tec_plan_codigo = " & VAR_SOL & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "   "
                        db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "  "
                    Else
                        Set rs_aux3 = New ADODB.Recordset
                        If rs_aux3.State = 1 Then rs_aux3.Close
                        rs_aux3.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & " AND bien_codigo = ''  ", db, adOpenKeyset, adLockBatchOptimistic
                        If rs_aux3.RecordCount > 0 Then
                            rs_aux3.MoveFirst
                            'If VAR_COD0 < var_cod5 Then     'And rs_aux3!estado_activo = "REG"
                                'db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "'   WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & VAR_COD4 & "',  tec_plan_codigo = " & VAR_SOL & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
                                db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                'VAR_COD0 = VAR_COD0 + 1
                                'CONT3 = 1
                                db.Execute "Update ao_ventas_cabecera Set estado_crono = 'APR' Where venta_codigo = " & NumComp & "  "
                                'VAR_EMES = "NADA"
                            'End If
                        Else
                            'POR SI NO TIENE fmes_plan
                        End If
                    End If
                    rs_aux2.MoveNext
                Wend
            rs_aux1.MoveNext
            End If
        Wend
    End If
'    rs_aux99.MoveNext
'    Wend
'    End If
End Sub

Private Sub BtnAprobar_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  If Ado_datos.Recordset!estado_codigo_verif = "REG" Then
        MsgBox "No se puede APROBAR debe registrar el PLAN DE CUOTAS, verifique los datos y vuelva a intentar ...", , "Atención"
        Exit Sub
  End If
  'Plan de Cuotas
    NumComp = Ado_datos.Recordset!venta_codigo
    VAR_DOL2 = Round(Ado_datos.Recordset!venta_monto_total_dol, 2)
    VAR_BS2 = Round(Ado_datos.Recordset!venta_monto_total_bs, 2)
    Set rs_aux21 = New ADODB.Recordset     'Plan de Cuotas
    If rs_aux21.State = 1 Then rs_aux21.Close
    rs_aux21.Open "Select SUM(cobranza_programada_bs) AS SumaCuota from ao_ventas_cobranza_prog WHERE venta_codigo = " & NumComp & " AND es_liquidacion = 'NO' ", db, adOpenStatic
    If rs_aux21.RecordCount > 0 Then
        If Round(VAR_BS2, 0) = Round(rs_aux21!SumaCuota, 0) Then        'Las Cuotas no igualan con el Total del contrato
        Else
            MsgBox "No se puede APROBAR, la SUMA de las Cuotas Bs, NO iguala con el TOTAL del Contrato Bs, verifique y vuelva a intentar ...", , "Atención"
            Exit Sub
        End If
    End If

  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  VAR_VALD = "OK"
  Call valida_campos2
  If VAR_VAL = "OK" And VAR_VALD = "OK" Then
     If Ado_datos.Recordset.RecordCount > 0 Then
        'ACTUALIZA Unidad_Medida
        db.Execute "UPDATE ac_bienes SET unimed_codigo_empaque = unimed_codigo where (unimed_codigo_empaque Is Null) "
        db.Execute "UPDATE ao_solicitud_bienes SET unimed_codigo_empaque = unimed_codigo where (unimed_codigo_empaque Is Null) "
        'ACTUALIZA almacen_tipo ='Q' (Equipos)
        db.Execute "UPDATE ac_bienes SET almacen_tipo ='Q' WHERE (par_codigo ='43340' AND almacen_tipo IS NULL) "
        db.Execute "UPDATE ao_solicitud_bienes SET almacen_tipo ='Q' WHERE (par_codigo ='43340' AND almacen_tipo IS NULL) "
        db.Execute "UPDATE ao_ventas_detalle SET almacen_tipo ='Q' WHERE (par_codigo ='43340' AND almacen_tipo IS NULL) "
        'ACTUALIZA
        db.Execute "UPDATE ao_ventas_cabecera SET solicitud_tipo = 7 WHERE (unidad_codigo LIKE '%REP%' AND solicitud_tipo <> 7) "
        db.Execute "UPDATE ao_ventas_cabecera SET solicitud_tipo = 10 WHERE (unidad_codigo LIKE '%MAN%' AND solicitud_tipo <> 10) "
        db.Execute "UPDATE ao_ventas_cabecera SET solicitud_tipo = 4 WHERE (unidad_codigo LIKE '%INS%' AND solicitud_tipo <> 4) "
        'INI VALIDACIONES
        If IsNull(Ado_datos.Recordset("venta_tipo")) Then        'Or (Ado_datos.Recordset("venta_monto_total_bs") = 0)       ' JQA ENE-2016
            MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
            Exit Sub
        End If
        'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'        Call CRONO_MTTO
        'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
        VAR_COD1 = Ado_datos.Recordset!unidad_codigo
        VAR_TIPOV = Ado_datos.Recordset!venta_tipo
        NumComp = Ado_datos.Recordset!venta_codigo
        VAR_ZONA = Ado_datos.Recordset!zpiloto_codigo
        VAR_TIPO = Ado_datos.Recordset!solicitud_tipo
        VAR_EXPOR = "NN"
        db.Execute "update ao_ventas_cabecera set estado_cancelado = 'N' Where venta_codigo = " & NumComp & "  "
        db.Execute "UPDATE ao_ventas_cabecera SET ao_ventas_cabecera.trans_codigo  = ao_solicitud.trans_codigo FROM ao_ventas_cabecera INNER JOIN ao_solicitud ON ao_ventas_cabecera.unidad_codigo = ao_solicitud.unidad_codigo AND ao_ventas_cabecera.solicitud_codigo  = ao_solicitud.solicitud_codigo WHERE (ao_ventas_cabecera.venta_codigo = " & NumComp & ") AND (ao_ventas_cabecera.trans_codigo IS NULL) "
        If VAR_COD1 = "DNREP" Or VAR_COD1 = "DREPS" Or VAR_COD1 = "DREPB" Or VAR_COD1 = "DREPC" Or VAR_COD1 = "DNINS" Or VAR_COD1 = "DINSB" Or VAR_COD1 = "DINSS" Or VAR_COD1 = "DINSC" Then
            'DETALLE DE BB.SS.
            Set rs_aux19 = New ADODB.Recordset
            If rs_aux19.State = 1 Then rs_aux19.Close
            rs_aux19.Open "select * from ao_ventas_detalle where venta_codigo= " & NumComp & " AND par_codigo <> '43340' ", db, adOpenKeyset, adLockBatchOptimistic
            If rs_aux19.RecordCount = 0 Then
                MsgBox "No existen Bienes o Servicios en el detalle, verifique y vuelva a intentar ... ", vbInformation, "Información!"
                Exit Sub
            End If
            If VAR_TIPOV = "R" Then     'IMPORTACION=R
                'ALCANCE - VALIDACION
                Set rs_aux9 = New ADODB.Recordset
                If rs_aux9.State = 1 Then rs_aux9.Close
                rs_aux9.Open "Select * from ao_ventas_alcance where venta_codigo = " & NumComp & "   ", db, adOpenStatic
                If rs_aux9.RecordCount <= 1 Then
                    MsgBox "No se puede APROBAR debe registrar el ALCANCE DEL CONTRATO, verifique los datos y vuelva a intentar ...", , "Atención"
                    Exit Sub
                End If
                'REPUESTOS IMPORTADOS - VALIDACION
                Set rs_aux19 = New ADODB.Recordset
                If rs_aux19.State = 1 Then rs_aux19.Close
                rs_aux19.Open "select * from ao_ventas_detalle where venta_codigo= " & NumComp & " AND par_codigo = '39810' ", db, adOpenKeyset, adLockBatchOptimistic
                If rs_aux19.RecordCount = 0 Then
                    MsgBox "No existen Repuestos IMPORTADOS, verifique y vuelva a intentar ... ", vbInformation, "Información!"
                    Call OptFilGral1_Click
                    Exit Sub
                End If
                sino = MsgBox("Esta seguro de Aprobar y Enviar el registro a COMEX ? ", vbYesNo, "Confirmando")
                VAR_EXPOR = "SI"
            Else
                sino = MsgBox("Esta seguro de Aprobar y Enviar el registro a Almacenes ? ", vbYesNo, "Confirmando")
                VAR_EXPOR = "NO"
            End If
        Else
            sino = MsgBox("Esta seguro de Aprobar el registro? ", vbYesNo, "Confirmando")
            VAR_EXPOR = "NO"
        End If
        'FIN VALIDACIONES
         If Ado_datos.Recordset("estado_codigo") = "REG" Then
           If sino = vbYes Then
               'ASIGNA A VARIABLES CAMPOS CLAVES
               'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
               VAR_COD4 = Ado_datos.Recordset!unidad_codigo
               VAR_GLOSA = Ado_datos.Recordset!venta_descripcion
               VAR_DOL2 = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
               VAR_BS2 = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
               VAR_UNIMED = Ado_datos.Recordset!unimed_codigo_tec
               'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW

               VAR_SOL = Ado_datos.Recordset!solicitud_codigo
               VAR_MED = Ado_datos.Recordset!unimed_codigo_tec
               VAR_MED2 = Ado_datos.Recordset!unimed_codigo_cobr
               FInicio = Ado_datos.Recordset!venta_fecha_inicio
               FFin = Ado_datos.Recordset!venta_fecha_fin
               TimeD = Ado_datos.Recordset!venta_plazo_dias_calendario
               CANTOT = Ado_datos.Recordset!venta_cantidad_total
               VAR_GLOSA2 = Ado_datos.Recordset!venta_descripcion
               VAR_PROY2 = Ado_datos.Recordset!edif_codigo
               VAR_CITE = Ado_datos.Recordset!unidad_codigo_ant         'OS - 36AO - 36NB - 36NO...
               VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
               VAR_BEND = dtc_desc2.Text
               VAR_EDIFD = dtc_desc3.Text
               VAR_UNID = dtc_desc1.Text
               'VAR_DPTO = Left(VAR_PROY2, 1)
               'VAR_DPTO = Ado_datos.Recordset!depto_codigo
               If Ado_datos.Recordset!depto_codigo = Left(Ado_datos.Recordset!edif_codigo, 1) Then
                     VAR_DPTO = Ado_datos.Recordset!depto_codigo
                Else
                     VAR_DPTO = Left(Ado_datos.Recordset!edif_codigo, 1)
                End If
               VARG_ORGD = ""
               VAR_CTAD = ""
               'Dim VARG_ORGD, VAR_CTAD, ,  As String
               'If Ado_datos.Recordset("venta_tipo") = "C" Or Ado_datos.Recordset("venta_tipo") = "V" Then
               If Ado_datos.Recordset("venta_tipo") <> "D" Then
                    db.Execute "update gc_beneficiario set beneficiario_deudor = 'SI' where beneficiario_codigo = '" & dtc_codigo2 & "' "
               End If
               'Actualiza venta_precio_total_bs y venta_precio_total_dol
               db.Execute "update ao_ventas_detalle set venta_precio_total_bs = round(venta_det_cantidad * venta_precio_unitario_bs,2)  "
               db.Execute "update ao_ventas_detalle set venta_precio_total_dol = venta_det_cantidad * venta_precio_unitario_dol  "

               'INI Correl OS por Depto
                Set rs_aux17 = New ADODB.Recordset
                If rs_aux17.State = 1 Then rs_aux17.Close
                rs_aux17.Open "Select * from gc_departamento where depto_codigo = " & VAR_DPTO & "  ", db, adOpenStatic
                If rs_aux17.RecordCount > 0 Then
                    VAR_DPTOD = rs_aux17!depto_descripcion
                    'Actualiza correaltivo OS ...
                    If VAR_COD1 = "DNREP" Or VAR_COD1 = "DREPS" Or VAR_COD1 = "DREPB" Or VAR_COD1 = "DREPC" Or VAR_COD1 = "DNINS" Or VAR_COD1 = "DINSB" Or VAR_COD1 = "DINSS" Or VAR_COD1 = "DINSC" Then
                        If VAR_TIPOV = "R" Then
                           'REPUESTOS IMPORTADOS - VALIDACION
                           VAR_EXPOR = "SI"
                           If Left(VAR_CITE, 3) = "36A" Then
                                VAR_CORREL = rs_aux17!correl_AO
                            Else
                                VAR_CORREL = rs_aux17!correl_AO + 1
                            End If
                            db.Execute "update ao_ventas_cabecera set doc_numero = " & VAR_CORREL & " Where ao_ventas_cabecera.venta_codigo = " & NumComp & " "
                            db.Execute "Update gc_departamento Set correl_AO = " & VAR_CORREL & " Where depto_codigo = " & VAR_DPTO & "   "
                        Else
                            VAR_EXPOR = "NO"
                            VAR_CORREL = rs_aux17!correl_OS + 1
                            db.Execute "update ao_ventas_cabecera set doc_numero = " & VAR_CORREL & " Where ao_ventas_cabecera.venta_codigo = " & NumComp & " "
                            db.Execute "Update gc_departamento Set correl_OS = " & VAR_CORREL & " Where depto_codigo = " & VAR_DPTO & "   "
                        End If
                    End If
                Else
                    VAR_DPTOD = "LA PAZ"
                    If VAR_COD1 = "DNREP" Or VAR_COD1 = "DREPS" Or VAR_COD1 = "DREPB" Or VAR_COD1 = "DREPC" Or VAR_COD1 = "DNINS" Or VAR_COD1 = "DINSB" Or VAR_COD1 = "DINSS" Or VAR_COD1 = "DINSC" Then
                        sino = MsgBox("Elija SI: para enviar a COMEX para Importacion del Repuesto ..." & vbCr & _
                        "Elija NO: para enviar a Almacen de Repuestos ...", vbYesNo + vbQuestion, "Atención")
                        If sino = vbYes Then
                            VAR_EXPOR = "SI"
                            VAR_CORREL = 1001
                            db.Execute "update ao_ventas_cabecera set doc_numero = " & VAR_CORREL & " Where ao_ventas_cabecera.venta_codigo = " & NumComp & " "
                            db.Execute "Update gc_departamento Set correl_AO = " & VAR_CORREL & " Where depto_codigo = '0'   "
                        Else
                            VAR_EXPOR = "NO"
                            VAR_CORREL = 200001
                            db.Execute "update ao_ventas_cabecera set doc_numero = " & VAR_CORREL & " Where ao_ventas_cabecera.venta_codigo = " & NumComp & " "
                            db.Execute "Update gc_departamento Set correl_OS = " & VAR_CORREL & " Where depto_codigo = '0'   "
                        End If
                    End If
                End If
               'VAR_CORREL = 19939
               'FIN Correl OS por Depto

               'ACTUALIZA CORRELATIVO DE DOC. RESPALDO
               If VAR_COD1 = "DNMAN" Or VAR_COD1 = "DMANS" Or VAR_COD1 = "DMANB" Or VAR_COD1 = "DMANC" Or VAR_COD1 = "DNEME" Or VAR_COD1 = "DEMEB" Or VAR_COD1 = "DEMES" Or VAR_COD1 = "DEMEC" Then
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos.Recordset!doc_codigo & "'  "
                    rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                    If rs_aux2.RecordCount > 0 Then
                        rs_aux2!correl_doc = rs_aux2!correl_doc + 1
                        VAR_CORRELM = rs_aux2!correl_doc + 1
                        db.Execute "update ao_ventas_cabecera set doc_numero = " & VAR_CORRELM & " Where ao_ventas_cabecera.venta_codigo = " & NumComp & " "
    '                    Ado_datos.Recordset!doc_numero = rs_aux2!correl_doc
                        'Txt_campo1.Caption = rs_aux2!correl_doc
                        'rs_aux2.Update
                    Else
                        rs_aux2!correl_doc = 1
                        VAR_CORRELM = 1
                        db.Execute "update ao_ventas_cabecera set doc_numero = " & VAR_CORRELM & " Where ao_ventas_cabecera.venta_codigo = " & NumComp & " "
                        'rs_aux2.Update
                    End If
                    'Actualiza correaltivo Doc ...
                    db.Execute "Update gc_documentos_respaldo Set correl_doc = " & VAR_CORRELM & " Where doc_codigo = '" & Ado_datos.Recordset!doc_codigo & "'  "
                End If
               'INI GRABA ao_ventas_alcance
               Select Case VAR_COD1
                    Case "DNINS", "DINSB", "DINSC", "DINSS"
                        VAR_TIPO = 4
                        'ODS
                        If VAR_EXPOR = "SI" Then
                            If Left(VAR_CITE, 4) = "36AO" Then
                                VAR_CITE = VAR_CITE
                            Else
                                VAR_CITE = "36AO" + Trim(Str(VAR_CORREL))
                            End If
                        Else
                            VAR_CITE = "OS-" + Trim(Str(VAR_CORREL))
                        End If
                        TxtConcepto.Text = "SERVICIO DE INSTALACIONES. Segun: " + VAR_CITE + ". Edificio: " + RTrim(dtc_desc3.Text) + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
                        VAR_ARCH = "COM_" + RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(IIf(IsNull(Ado_datos.Recordset!doc_numero), 1, Ado_datos.Recordset!doc_numero)))
                    Case "DNAJS", "DAJSB", "DAJSC", "DAJSS"
                        VAR_TIPO = 5
                        If VAR_EXPOR = "SI" Then
                            VAR_CITE = "36AO" + Trim(Str(VAR_CORREL))
                        Else
                            VAR_CITE = "OS-" + Trim(Str(VAR_CORREL))
                        End If
                        TxtConcepto.Text = "Servicio de AJUSTE. Segun: " + VAR_CITE + ". Edificio: " + RTrim(dtc_desc3.Text) + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
                        VAR_ARCH = "COM_" + RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(IIf(IsNull(Ado_datos.Recordset!doc_numero), 1, Ado_datos.Recordset!doc_numero)))
                    Case "DNMAN", "DMANB", "DMANC", "DMANS"
                        VAR_TIPO = 10
                        'CONTRATO
                        TxtConcepto.Text = "ServiSERVICIO DE MANTENIMIENTO INTEGRAL. Edificio: " + RTrim(dtc_desc3.Text) + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
                        VAR_ARCH = "TEC_" + RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(IIf(IsNull(Ado_datos.Recordset!doc_numero), 1, Ado_datos.Recordset!doc_numero)))
                        VAR_CITE = Ado_datos.Recordset!unidad_codigo_ant
                        'ACTUALIZA ALCANCE
                        If Ado_datos.Recordset("estado_alcance") = "N" Then
                          Set rs_aux9 = New ADODB.Recordset
                          If rs_aux9.State = 1 Then rs_aux9.Close
                          rs_aux9.Open "Select * from ao_ventas_alcance where venta_codigo = " & NumComp & "  And solicitud_tipo = " & VAR_TIPO & "  ", db, adOpenStatic
                          If rs_aux9.RecordCount = 0 Then
                              db.Execute "INSERT INTO ao_ventas_alcance (ges_gestion, venta_codigo, solicitud_tipo, solicitud_tipo_descripcion, unidad_codigo_tec, venta_tiempo_dias, fecha_inicio_alcance, fecha_fin_alcance , estado_codigo, usr_codigo, fecha_registro) VALUES ('" & glGestion & "', " & NumComp & ", " & VAR_TIPO & ", 'MANTENIMIENTO PREVENTIVO DE EQUIPOS', '" & VAR_COD1 & "', '" & TimeD & "', '" & FInicio & "' , '" & FFin & "', 'APR', '" & glusuario & "', '" & Date & "' )"
                              db.Execute "update ao_ventas_cabecera set estado_alcance = 'S' Where venta_codigo = " & NumComp & " "
                          Else
                              db.Execute "update ao_ventas_cabecera set estado_alcance = 'S' Where venta_codigo = " & NumComp & " "
                          End If
                        End If
                    Case "DNREP", "DREPB", "DREPC", "DREPS"
                        VAR_TIPO = 7
                        'ODS
                        If VAR_EXPOR = "SI" Then
                            If Left(VAR_CITE, 3) = "36A" Then
                                VAR_CITE = VAR_CITE
                            Else
                                VAR_CITE = "36AO" + Trim(Str(VAR_CORREL))
                            End If
                        Else
                            VAR_CITE = "OS-" + Trim(Str(VAR_CORREL))
                        End If
                        TxtConcepto.Text = "SERVICIO DE REPARACIONES. Segun: " + VAR_CITE + ". Edificio: " + RTrim(dtc_desc3.Text) + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
                        VAR_ARCH = "TEC_" + RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(IIf(IsNull(Ado_datos.Recordset!doc_numero), 1, Ado_datos.Recordset!doc_numero)))
                    Case "DNEME", "DEMEB", "DEMEC", "DEMES"
                        VAR_TIPO = 8
                        'CONTRATO
                        VAR_CITE = "OS-" + Trim(Str(VAR_CORREL))
                        TxtConcepto.Text = "Servicio de EMERGENCIAS. Segun: " + VAR_CITE + ". Edificio: " + RTrim(dtc_desc3.Text) + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
                        VAR_ARCH = "TEC_" + RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(IIf(IsNull(Ado_datos.Recordset!doc_numero), 1, Ado_datos.Recordset!doc_numero)))
                    Case "DNMOD", "DMODB", "DMODC", "DMODS"
                        VAR_TIPO = 9
                        'CONTRATO
                        VAR_ARCH = "MOD_" + RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(IIf(IsNull(Ado_datos.Recordset!doc_numero), 1, Ado_datos.Recordset!doc_numero)))
                    Case Else
                        MsgBox "No se ha definido el tipo " & vbCrLf & " de registro que está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
                        Exit Sub
               End Select
               'ACTUALIZA TOTALES EN CABECERA
               Call acumulaMont(Ado_datos.Recordset!ges_gestion, NumComp)

                ' GRABA Nombre de Archivo en ao_ventas_cabecera. VERIFICAR JQA 2014-07-08
                'rs_datos!doc_numero = Txt_campo1.Caption
                'VAR_ARCH = RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(Ado_datos.Recordset!doc_numero))
                db.Execute "update ao_ventas_cabecera set archivo_respaldo = '" & VAR_ARCH & "' + '.PDF' Where venta_codigo = " & NumComp & " "
                db.Execute "update ao_ventas_cabecera set archivo_respaldo_cargado = 'N' Where venta_codigo = " & NumComp & " "
                db.Execute "update ao_solicitud set unidad_codigo_ant = '" & VAR_CITE & "' Where unidad_codigo= '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " "
                db.Execute "update ao_ventas_cabecera set unidad_codigo_ant = '" & VAR_CITE & "' Where venta_codigo = " & NumComp & " "

                db.Execute "update ao_ventas_cabecera set venta_descripcion = '" & TxtConcepto.Text & "'  Where venta_codigo = " & NumComp & " "

               'INI GENERA INFORMACION COMEX, INSTALACION, AJUSTE Y/O MANTENIMIENTO
               'CABECERA CRONOGRAMA
               'If VAR_TIPOV = "C" Then
                 Set rs_aux1 = New ADODB.Recordset
                 If rs_aux1.State = 1 Then rs_aux1.Close
                 rs_aux1.Open "select * from ao_ventas_alcance where venta_codigo= " & NumComp & "  ", db, adOpenKeyset, adLockBatchOptimistic
                 If rs_aux1.RecordCount > 0 Then
                    VAR_COD1 = rs_aux1!unidad_codigo_tec
                    Select Case rs_aux1!solicitud_tipo
                      Case 7
                        ' PREGUNTAR SI ES EXPORTACION   ------------------------- WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW 2022-MAR-18
                        If VAR_EXPOR = "SI" Then
                            VAR_SOLTIPO = rs_aux1!solicitud_tipo
                            Call PARA_COMEX
                        End If

                      Case 6

                   'rs_aux1.MoveFirst
                   'While Not rs_aux1.EOF
                   '  VAR_COD1 = rs_aux1!unidad_codigo_tec
                   '  If (VAR_COD1 = "COMEX" Or VAR_COD1 = "DNREP" Or VAR_COD1 = "DREPS" Or VAR_COD1 = "DREPB" Or VAR_COD1 = "DREPC" Or VAR_COD1 = "DNINS" Or VAR_COD1 = "DINSC" Or VAR_COD1 = "DINSS" Or VAR_COD1 = "DINSB") Then             'INI GRABA CRONOGRAMA COMEX O REPARACIONES
                   '     ' PREGUNTAR SI ES EXPORTACION   ------------------------- WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW 2022-MAR-18
                   '     If VAR_EXPOR = "SI" Then
                   '         VAR_SOLTIPO = rs_aux1!solicitud_tipo
                   '         Call PARA_COMEX
                   '     End If
                   '  Else
                        If VAR_PROY2 = "" Then
                            VAR_PROY2 = Ado_datos.Recordset!edif_codigo
                        End If
                            Set rs_aux4 = New ADODB.Recordset
                            If rs_aux4.State = 1 Then rs_aux4.Close
                            'rs_aux4.Open "select * from ao_ventas_detalle where venta_codigo= " & NumComp & " AND (par_codigo ='43340') ", db, adOpenKeyset, adLockBatchOptimistic
                            rs_aux4.Open "select * from ao_ventas_cobranza_prog where venta_codigo= " & NumComp & "  ", db, adOpenKeyset, adLockBatchOptimistic
                            If rs_aux4.RecordCount > 0 Then
                               rs_aux4.MoveFirst
                               While Not rs_aux4.EOF
                                 Set rs_aux3 = New ADODB.Recordset
                                 If rs_aux3.State = 1 Then rs_aux3.Close
                                 rs_aux3.Open "select * from to_cronograma_mensual where ges_gestion = '" & rs_aux4!gestion & "' AND fmes_correl = " & rs_aux4!cobranza_mes & " AND zpiloto_codigo = " & VAR_ZONA & "  ", db, adOpenKeyset, adLockOptimistic
                                 If rs_aux3.RecordCount > 0 Then
                                    db.Execute "UPDATE ao_ventas_cobranza_prog SET fmes_plan = " & rs_aux3!fmes_plan & " WHERE correl_prog = " & rs_aux4!correl_prog & " "
                                    VAR_TECPLAN = rs_aux3!fmes_plan
                                 End If
                                 Set rstdestino = New ADODB.Recordset
                                 If rstdestino.State = 1 Then rstdestino.Close
                                 'rstdestino.Open "select * from to_cronograma_detalle where tec_plan_codigo = " & VAR_TECPLAN & " AND bien_codigo = '" & rs_aux4!bien_codigo & "' ", db, adOpenKeyset, adLockBatchOptimistic
                                 rstdestino.Open "select * from ao_ventas_detalle where venta_codigo= " & NumComp & " AND (par_codigo ='43340') ", db, adOpenKeyset, adLockBatchOptimistic
                                 If rstdestino.RecordCount > 0 Then
                                    'If IsNull(rs_aux4!bien_cantidad_por_empaque) Or (rs_aux4!bien_cantidad_por_empaque = 0) Then
                                    '    rs_aux4!bien_cantidad_por_empaque = 2
                                    'End If
                                    'db.Execute "UPDATE to_cronograma_detalle SET fecha_inicio='" & Format(rs_aux1!fecha_inicio_alcance, "dd/mm/yyyy") & "', fecha_fin='" & Format(rs_aux1!fecha_fin_alcance, "dd/mm/yyyy") & "', bien_tiempo_dias=" & rs_aux1!venta_tiempo_dias & ", usr_codigo='" & glusuario & "', fecha_registro='" & Date & "'  WHERE  bien_codigo = '" & rstdestino!bien_codigo & "' and tec_plan_codigo = " & VAR_TECPLAN & "    "
                                    db.Execute "UPDATE to_cronograma_vs_ventas SET fecha_inicio='" & Format(rs_aux1!fecha_inicio_alcance, "dd/mm/yyyy") & "', fecha_fin='" & Format(rs_aux1!fecha_fin_alcance, "dd/mm/yyyy") & "', bien_tiempo_dias=" & rs_aux1!venta_tiempo_dias & ", usr_codigo='" & glusuario & "', fecha_registro='" & Date & "'  WHERE  bien_codigo = '" & rstdestino!bien_codigo & "' and tec_plan_codigo = " & VAR_TECPLAN & "    "
                                 Else
                                    'db.Execute "INSERT INTO to_cronograma_detalle (ges_gestion, unidad_codigo_tec, tec_plan_codigo, bien_codigo, beneficiario_codigo, grupo_codigo, subgrupo_codigo, par_codigo, munic_codigo, fecha_inicio, fecha_fin, bien_tiempo_dias, hora_inicio, hora_fin, estado_codigo, usr_codigo, fecha_registro, bien_cantidad_por_empaque, precio_unitario_bs) " & _
                                    '"VALUES ('" & glGestion & "', '" & VAR_COD1 & "', " & correldetalle & ", '" & rs_aux4!bien_codigo & "', '0', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '" & Left(VAR_PROY2, 5) & "', '" & Format(rs_aux1!fecha_inicio_alcance, "dd/mm/yyyy") & "', '" & Format(rs_aux1!fecha_fin_alcance, "dd/mm/yyyy") & "', " & rs_aux1!venta_tiempo_dias & ", '8:00', '18:30', 'REG', '" & glusuario & "', '" & Date & "', " & rs_aux4!bien_cantidad_por_empaque & ", " & Round(rs_aux4!venta_precio_unitario_bs, 2) & "  )"

                                    'fecha_conformidad, fecha_equipo_hdm, bien_codigo1, bien_codigo2, bien_codigo3, bien_codigo4, bien_codigo5,
                                    db.Execute "INSERT INTO to_cronograma_vs_ventas (fmes_plan, correl_prog, bien_codigo,  doc_numero, doc_numero_equipo, carta, doc_numero_carta, cantidad1, cantidad2, cantidad3, cantidad4, cantidad5, estado_prog, estado_crono, estado_codigo, usr_codigo,       fecha_registro) " & _
                                     " VALUES (" & VAR_TECPLAN & ", " & rs_aux4!correl_prog & ", '" & rstdestino!bien_codigo & "', '0', '0',           'NO',   '0',                '0',        '0',    '0',        '0',        '0',        'REG',      'REG',      'REG',    '" & glusuario & "', '" & Date & "'  )  "

                                 End If
                                 rs_aux4.MoveNext
                               Wend
                            End If
                        End Select
                        'Call BtnAñadir2_Click               'GENERA CRONOGRAMA DE MANTENIMIENTO
                        If VAR_COD1 = "DNMAN" Or VAR_COD1 = "DMANS" Or VAR_COD1 = "DMANB" Or VAR_COD1 = "DMANC" Or VAR_COD1 = "DNEME" Or VAR_COD1 = "DEMEB" Or VAR_COD1 = "DEMES" Or VAR_COD1 = "DEMEC" Then
                            Call CRONO_MTTO
                        End If
                 End If                 ' FIN ALCANCE
                db.Execute "UPDATE dbo.ao_ventas_cabecera SET usr_codigo_aprueba = '" & glusuario & "' WHERE venta_codigo = " & NumComp & " "
               db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo = 'APR' Where ao_ventas_cabecera.venta_codigo = " & NumComp & " "
               MsgBox "La Venta fue Enviada y Aprobada Exitosamente... ", vbInformation, "Información!"
               'FIN GENERA INFORMACION COMEX, INSTALACION, AJUSTE Y/O MANTENIMIENTO
               Call OptFilGral1_Click
           Else
                MsgBox "La Aprobación fue cancelada ... ", vbInformation, "Información!"
           End If
         End If
         '           'INI CONTABILIZACION NUEVA
'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
            
            'Call Contabiliza_Contratos(correlv)

'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'            'FIN CONTABILIZACION NUEVA
     Else
        MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
     End If
  End If
  
  Exit Sub
UpdateErr:
  MsgBox Err.Description
    
End Sub

Private Sub PARA_COMEX()

  If Ado_datos.Recordset.RecordCount > 0 Then
     If Ado_datos.Recordset("estado_codigo") = "REG" Then
        Dim VAR_SOLTIPO2 As String
        VAR_TIPOV = "R"       'Ado_datos.Recordset!venta_tipo
        'INI GENERA INFORMACION COMEX, INSTALACION, AJUSTE
        If VAR_SOLTIPO = "15" Then
            VAR_SOLTIPO = "3"
        End If
        If VAR_SOLTIPO = "7" Then
            VAR_SOLTIPO = "3"
        End If
        VAR_SOLTIPO2 = Str(VAR_SOLTIPO)
        Set rs_aux11 = New ADODB.Recordset
        If rs_aux11.State = 1 Then rs_aux11.Close
        'rs_aux1.Open "select * from ao_ventas_alcance where venta_codigo= " & NumComp & "  ", db, adOpenKeyset, adLockBatchOptimistic
        rs_aux11.Open "select * from ac_bienes where kit = '90' AND observaciones = '" & Trim(VAR_SOLTIPO2) & "' ", db, adOpenKeyset, adLockBatchOptimistic
        If rs_aux11.RecordCount > 0 Then
          rs_aux11.MoveFirst
          While Not rs_aux11.EOF
            VAR_COD1 = Ado_datos.Recordset!unidad_codigo            'rs_aux1!unidad_codigo_tec
            'VAR_CANT0 = "1"         'Round((rs_aux1!venta_tiempo_dias / 30), 0)
            'rs_aux1.MoveNext
            If VAR_COD1 = "DNREP" Or VAR_COD1 = "DREPS" Or VAR_COD1 = "DREPB" Or VAR_COD1 = "DREPC" Or VAR_COD1 = "DNINS" Or VAR_COD1 = "DINSS" Or VAR_COD1 = "DINSB" Or VAR_COD1 = "DINSC" Then         'INI GRABA COMEX
'                    'WWWWWWWWWWWWWWW
               'NumComp = Ado_datos.Recordset!venta_codigo
               'VAR_TIPOV = Ado_datos.Recordset!venta_tipo
               Set rs_aux3 = New ADODB.Recordset
               If rs_aux3.State = 1 Then rs_aux3.Close
               rs_aux3.Open "select * from ao_compra_cabecera where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo = " & VAR_SOL & " ", db, adOpenKeyset, adLockOptimistic
               If rs_aux3.RecordCount = 0 Then
                   'INI CORREL DETALLE COMPRA
                   Set rs_aux2 = New ADODB.Recordset
                   If rs_aux2.State = 1 Then rs_aux2.Close
                   rs_aux2.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & VAR_COD1 & "'  ", db, adOpenKeyset, adLockOptimistic
                   If rs_aux2.RecordCount > 0 Then
                      rs_aux2!correl_area = rs_aux2!correl_area + 1
                      correldetalle = rs_aux2!correl_area
                      rs_aux2.Update
                   End If
                   'FIN CORREL DETALLE COMPRA
                    'FALTANTES
                    ' beneficiario_codigo_resp, doc_numero, nro_nota_remision, estado_codigo_tra, estado_codigo_nac, estado_codigo_des,
                    ' hora_registro, usr_codigo_aprueba, fecha_registro_aprueba, archivo_respaldo, archivo_respaldo_cargado, estado_codigo_tec, adjudica_codigo
                   rs_aux3.AddNew
                   rs_aux3!ges_gestion = glGestion     'Year(Date)
                   'rs_aux3!compra_codigo = 0      'Autonumerico
                   If VAR_EXPOR = "SI" Then
                        rs_aux3!unidad_codigo_adm = "COMEX"
                   Else
                        rs_aux3!unidad_codigo_adm = VAR_COD1
                   End If
                   rs_aux3!solicitud_codigo_adm = correldetalle
                   rs_aux3!unidad_codigo = VAR_COD4
                   rs_aux3!solicitud_codigo = VAR_SOL
                   rs_aux3!edif_codigo = VAR_PROY2
                   rs_aux3!beneficiario_codigo = VAR_BENEF
                   rs_aux3!beneficiario_codigo_alm = IIf(IsNull(Ado_datos.Recordset!beneficiario_codigo_resp), "0", Ado_datos.Recordset!beneficiario_codigo_resp)
                   rs_aux3!solicitud_tipo = rs_aux11!observaciones     '"15"
                   rs_aux3!venta_tipo = VAR_TIPOV
                   rs_aux3!unidad_codigo_ant = VAR_CITE
                   rs_aux3!compra_fecha = Date
                   rs_aux3!compra_DESCRIPCION = "COMPRA POR: " + VAR_GLOSA
                   rs_aux3!compra_observaciones = "PROVISION E IMPORTACION DE EQUIPOS Y/O REPUESTOS"
                   rs_aux3!compra_cantidad_total = Ado_datos.Recordset!venta_cantidad_total
                   rs_aux3!compra_monto_bs = VAR_BS2
                   rs_aux3!tipo_moneda = "USD"
                   rs_aux3!compra_monto_DOL = VAR_DOL2
                   rs_aux3!proceso_codigo = "CMX"
                   rs_aux3!subproceso_codigo = "CMX-01"
                   rs_aux3!etapa_codigo = "CMX-01-01"
                   rs_aux3!clasif_codigo = "CMX"
                   rs_aux3!doc_codigo = "R-207"
                   rs_aux3!poa_codigo = "4.1.1"
                   rs_aux3!doc_codigo_alm = "R-207"
                   rs_aux3!beneficiario_codigo_resp = "4828818"                ' OJO ---- (PARAMETRIZAR)
                   'doc_numero_alm
                   'GENERAR CORRELATIVO
                   rs_aux3!estado_codigo_eqp = "REG"
                   rs_aux3!estado_codigo = "REG"
                   rs_aux3!usr_codigo = glusuario
                   rs_aux3!fecha_registro = Date
                   rs_aux3.Update
                   
                   'db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " &
                       '"VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", '" & rs_aux4!bien_codigo & "', '1', " & rs_aux4!venta_precio_unitario_bs & ", '0', " & rs_aux4!venta_precio_total_bs & ", " & rs_aux4!venta_precio_unitario_dol & ", '0', " & rs_aux4!venta_precio_total_dol & ", '" & concepto_venta & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
                      
                   'DETALLE Carga ao_ventas_detalle
                   'Set rstdestino = New ADODB.Recordset
                   'If rstdestino.State = 1 Then rstdestino.Close
                   'rstdestino.Open "select * from ao_compra_detalle  ", db, adOpenKeyset, adLockBatchOptimistic
                   'INI DISTRIBUYE TRAMITES EN ao_compra_detalle
                   'Select Case rs_aux1!solicitud_tipo
                   
                   'Select Case rs_aux11!observaciones
                   '    Case 3, 7, 4
                           'VAR_TRAMITE = "BANCO"
'                                'EQUIPOS
'                                Set rs_aux4 = New ADODB.Recordset
'                                If rs_aux4.State = 1 Then rs_aux4.Close
'                                rs_aux4.Open "select * from ao_ventas_detalle where venta_codigo= " & NumComp & " AND PAR_CODIGO = '43340' ", db, adOpenKeyset, adLockBatchOptimistic
'                                If rs_aux4.RecordCount > 0 Then
'                                   rs_aux4.MoveFirst
'                                   While Not rs_aux4.EOF
'                                        db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " & _
'                                        "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", '" & rs_aux4!bien_codigo & "', '1', " & rs_aux4!venta_precio_unitario_bs & ", '0', " & rs_aux4!venta_precio_total_bs & ", " & rs_aux4!venta_precio_unitario_dol & ", '0', " & rs_aux4!venta_precio_total_dol & ", '" & rs_aux4!concepto_venta & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '1', '" & glusuario & "', '" & Date & "')"
'                                        rs_aux4.MoveNext
'                                   Wend
'                                Else
'                                    MsgBox "No existe Equipos, verifique el registro y vuelva a intentar ... ", vbInformation, "Información!"
'                                End If
                           'REPUESTOS
                           Set rs_aux19 = New ADODB.Recordset
                           If rs_aux19.State = 1 Then rs_aux19.Close
                           rs_aux19.Open "select * from ao_ventas_detalle where venta_codigo= " & NumComp & " AND almacen_tipo = 'R' ", db, adOpenKeyset, adLockBatchOptimistic
                           If rs_aux19.RecordCount > 0 Then
                              rs_aux19.MoveFirst
                              While Not rs_aux19.EOF
                                   db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " & _
                                   "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", '" & rs_aux19!bien_codigo & "', " & rs_aux19!venta_det_cantidad & ", " & rs_aux19!venta_precio_unitario_bs & ", '0', " & rs_aux19!venta_precio_total_bs & ", " & rs_aux19!venta_precio_unitario_dol & ", '0', " & rs_aux19!venta_precio_total_dol & ", '" & rs_aux19!concepto_venta & "', '" & rs_aux19!grupo_codigo & "', '" & rs_aux19!subgrupo_codigo & "', '" & rs_aux19!par_codigo & "', '1', " & VAR_ALMACEN & ", '" & glusuario & "', '" & Date & "')"
                                   rs_aux19.MoveNext
                              Wend
                           Else
                               MsgBox "No existen Repuestos IMPORTADOS, verifique el registro y vuelva a intentar ... ", vbInformation, "Información!"
                               Exit Sub
                           End If
                   ' VERIFICAR RRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR
'                        Set rs_aux4 = New ADODB.Recordset
'                        If rs_aux4.State = 1 Then rs_aux4.Close
'                        rs_aux4.Open "select * from ac_bienes where bien_codigo_anterior= '" & VAR_TRAMITE & "' AND KIT = '90'  ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rs_aux4.RecordCount > 0 Then
'                           rs_aux4.MoveFirst
'                           While Not rs_aux4.EOF
'                                db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto,            grupo_codigo,                   subgrupo_codigo,                    par_codigo,                 tipo_descuento, almacen_codigo , usr_usuario,       fecha_registro) " & _
'                                "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", '" & rs_aux4!bien_codigo & "', '1',        '0',                    '0',                    '0',                    '0',                        '0',                '0',                    '" & rs_aux4!bien_descripcion & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1',           '1',            '" & glusuario & "', '" & Date & "')"
'                                rs_aux4.MoveNext
'                           Wend
'                        End If
                   'cargar ADJUDICA_COMPRA Y CRONOGRAMA
               Else
                   Select Case rs_aux11!observaciones
                       Case 3, 15
                           VAR_TRAMITE = "BANCO"
                       Case 16
                           VAR_TRAMITE = "TRANS"
                       Case 17
                           VAR_TRAMITE = "ADUAN"
                       Case 18
                           VAR_TRAMITE = "DESCA"

                       Case Else
                           VAR_TRAMITE = "BANCO"
                           VAR_TRAMITE = "CONTR"
                   End Select
                   
                   db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto,            grupo_codigo,                   subgrupo_codigo,                    par_codigo,                 tipo_descuento, almacen_codigo , usr_usuario,       fecha_registro, solicitud_tipo) " & _
                           "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", '" & rs_aux11!bien_codigo & "', '1',        '0',                    '0',                    '0',                    '0',                        '0',                '0',                    '" & rs_aux11!bien_descripcion & "', '" & rs_aux11!grupo_codigo & "', '" & rs_aux11!subgrupo_codigo & "', '" & rs_aux11!par_codigo & "', '1',           '1',            '" & glusuario & "', '" & Date & "', " & Val(Trim(rs_aux11!observaciones)) & ")"
               End If
               'WWWWWWWWWW
'                 Else
            End If
            rs_aux11.MoveNext
          Wend
        End If
           
        'FIN GENERA INFORMACION COMEX, INSTALACION, AJUSTE Y/O MANTENIMIENTO
           'Call OptFilGral1_Click
        'MsgBox "La Venta fue Enviada Exitosamente... ", vbInformation, "Información!"
     End If
     'MsgBox "Verifique si el Registro ya fue APROBADO o ANULADO previamente ...", , "Atención"

 Else
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
 End If


End Sub

Private Sub Contabiliza_Contratos()
    ' Contabilizacion al momento de aprobacion
    'Base de datos
    Dim db2 As New ADODB.Connection
    ' Recordset
    Dim rs_aux100 As New ADODB.Recordset
    Dim rs_aux101 As New ADODB.Recordset
    'Declaracion de variables
    Dim VAR_CODTIPO As String
    Dim VAR_EMPRESA As Integer
    Dim VAR_TIPOCOMPID As Integer
    Dim VAR_FECHA As Date
    Dim VAR_MONEDAID As Integer
    Dim VAR_TIPOCAMBIO As Double
    Dim EntregadoA As String
    Dim VAR_DEBEORG As Double
    Dim VAR_HABERORG As Double
    'Impuestos
    Dim VAR_PorIVA As Double
    Dim VAR_PorIT As Double
    Dim VAR_PorITF As Double
    'Otros valores
    Dim VAR_ConFac As Integer
    Dim VAR_SinFac As Integer
    Dim VAR_Automatico As Integer
    Dim VAR_TipoNotaId As Integer
    Dim VAR_NotaNro As Integer
    Dim VAR_EstadoId As Integer
    Dim VAR_iConcurrency_id As Integer
    Dim VAR_TipoAsientoId As Integer
    Dim VAR_CentroCostoId As Integer
    Dim VAR_TipoRetencionId As Integer
    Dim VAR_TipoId As Integer
    Dim VAR_CompDetIdOrg As Integer
    ' Variables intermedias
    Dim VAR_transDescripcion As String
    ' Asignacion de valores del procedimiento Call graba_ingreso
    VAR_BS2 = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
    VAR_DOL2 = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
    ' Codigo Tipo
    VAR_CODTIPO = "DEI"
    ' Rubro codigo, descripcion, centro de costo id
    Set rs_aux100 = New ADODB.Recordset
    If rs_aux100.State = 1 Then rs_aux100.Close
    rs_aux100.Open "SELECT trans_descripcion, rubro_codigo, CentroCostoId FROM gc_tipo_transaccion WHERE trans_codigo = '" & Ado_datos.Recordset!trans_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
    If rs_aux100.RecordCount > 0 Then
        VAR_transDescripcion = rs_aux100!trans_descripcion
        VAR_PARTIDA = rs_aux100!rubro_codigo
        VAR_CentroCostoId = rs_aux100!CentroCostoId
        rs_aux100.Close
    Else
        VAR_transDescripcion = "None"
        VAR_PARTIDA = "None"
        VAR_CentroCostoId = "None"
    End If
    ' Empresa
    If VAR_TIPOV = "G" Then
        VAR_EMPRESA = 2
    Else
        VAR_EMPRESA = 1
    End If
    ' Fecha de venta
    VAR_FECHA = CDate(Ado_datos.Recordset!venta_fecha)
    ' Tipo de cambio -> BOB - USD
    If IsNull(Ado_datos.Recordset!venta_tipo_cambio) Or (Ado_datos.Recordset!venta_tipo_cambio = 0) Or (Ado_datos.Recordset!venta_tipo_cambio = 1) Then
        VAR_TIPOCAMBIO = GlTipoCambioOficial
    Else
        VAR_TIPOCAMBIO = Ado_datos.Recordset!venta_tipo_cambio
    End If
    'VAR_TIPOCAMBIO = Ado_datos.Recordset!venta_tipo_cambio
    ' Tipo moneda/Debe/Haber
    VAR_MONEDAID = 1
    VAR_DEBEORG = VAR_BS2 'Boliviano
    VAR_HABERORG = VAR_BS2 'Boliviano
    ' If Ado_datos.Recordset!tipo_moneda = "USD" Then
    '     VAR_MONEDAID = 2
    '     VAR_DEBEORG = VAR_DOL2 'Dolar
    '     VAR_HABERORG = VAR_DOL2 'Dolar
    ' Else
    '     VAR_MONEDAID = 1
    '     VAR_DEBEORG = VAR_BS2 'Boliviano
    '     VAR_HABERORG = VAR_BS2 'Boliviano
    ' End If
    ' Entregado A
    EntregadoA = "Responsable: " & Ado_datos.Recordset!beneficiario_codigo + " - " + Ado_datos.Recordset!beneficiario_denominacion
    ' Por Concepto
    VAR_CONCEPTO = "Devengamiento de contrato: " & Ado_datos.Recordset!unidad_codigo_ant & " - Edificio " & Ado_datos.Recordset!edif_codigo_corto
    Set rs_aux101 = New ADODB.Recordset
    If rs_aux101.State = 1 Then rs_aux101.Close
    rs_aux101.Open "select edif_descripcion from gc_edificaciones where edif_codigo = '" & VAR_PROY2 & "'  ", db, adOpenKeyset, adLockOptimistic
    If rs_aux101.RecordCount > 0 Then
        VAR_CONCEPTO = VAR_CONCEPTO & " " & rs_aux101!edif_descripcion
        rs_aux101.Close
    End If
    If VAR_transDescripcion <> "None" Then
        VAR_CONCEPTO = VAR_CONCEPTO & " - " & VAR_transDescripcion
    End If
    ' TipoCompId (Tipo comprobante id) Traspaso
    VAR_TIPOCOMPID = 3
    ' Impuestos
    VAR_PorIVA = 0.13
    VAR_PorIT = 0.03
    VAR_PorITF = 0.0015
    ' Otros valores
    VAR_ConFac = 0
    VAR_SinFac = 1
    VAR_Automatico = 1 '0 Permite edicion, 1 no permite editar
    VAR_TipoNotaId = Ado_datos.Recordset!solicitud_tipo
    VAR_NotaNro = Ado_datos.Recordset!venta_codigo
    ' Glosa general
    VAR_GLOSA = "INGRESO POR: " & Ado_datos.Recordset!venta_descripcion & " - Nro. Venta: " & VAR_NotaNro
    VAR_EstadoId = 11 'Libro Mayor requiere que sean de EstadoId = 10 Cerrado OR EstadoId = 11 Abierto
    VAR_TipoAsientoId = 0 ' Operativo
    VAR_TipoRetencionId = 0
    VAR_TipoId = 0
    VAR_CompDetIdOrg = 0
    ' Creamos conexion unica para CONDOBO
    db2.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CONDOBO;Data Source=SSOFIA"
    ' Procedimiento almacenado
    db2.Execute ("EXEC fp_contabiliza_ingresos '" & VAR_CODTIPO & "', '" & VAR_PARTIDA & "', " & VAR_EMPRESA & ", " & VAR_DPTO & ", " & VAR_TIPOCOMPID & ", '" & VAR_FECHA & "', " & VAR_MONEDAID & ", '" & VAR_TIPOCAMBIO & "', '" & VAR_DEBEORG & "', '" & VAR_HABERORG & "', '" & EntregadoA & "', '" & VAR_CONCEPTO & "', '" & VAR_PorIVA & "', '" & VAR_PorIT & "', '" & VAR_PorITF & "', " & VAR_ConFac & ", " & VAR_SinFac & ", " & VAR_Automatico & ", '" & VAR_GLOSA & "', " & VAR_TipoNotaId & ", " & VAR_NotaNro & ", " & VAR_EstadoId & ", '" & glusuario & "', " & VAR_TipoAsientoId & ", " & VAR_CentroCostoId & ", " & VAR_TipoRetencionId & ", " & VAR_TipoId & ", " & VAR_CompDetIdOrg & ", '" & VAR_PROY2 & "'")
    db2.Close
End Sub

Private Sub GENERA_COMPRA()
'    If rs_datos!estado_cotiza = "REG" Then
'      VAR_COD4 = Ado_datos.Recordset!unidad_codigo
'      VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'      VAR_PROY2 = Ado_datos.Recordset!edif_codigo
'      VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'        ' MANTENIMIENTO PREVENTIVO - INSUMOS y/o COMPRAS BB y SS
'                'EQUIPO
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    rs_aux2.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & parametro & "'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux2.RecordCount > 0 Then
'                       rs_aux2!correl_negocia = rs_aux2!correl_negocia + 1
'                       correldetalle = rs_aux2!correl_negocia
'                       rs_aux2.Update
'                    End If
'                    'WWWWWWWWWWWWWWW
'                    'NumComp = Ado_datos.Recordset!venta_codigo
'                    'VAR_TIPOV = Ado_datos.Recordset!venta_tipo
'
'                    Set rs_aux3 = New ADODB.Recordset
'                    If rs_aux3.State = 1 Then rs_aux3.Close
'                    rs_aux3.Open "select * from ao_compra_cabecera where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo = " & VAR_SOL & " ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux3.RecordCount = 0 Then
'                    'beneficiario_codigo_resp,'doc_numero,estado_codigo_tra, estado_codigo_nac, estado_codigo_des, hora_registro, usr_codigo_aprueba,'                      fecha_registro_aprueba
'                        rs_aux3.AddNew
'                        rs_aux3!ges_gestion = glGestion     'Year(Date)
'                        'rs_aux3!compra_codigo = 0      'Autonumerico
'                        rs_aux3!unidad_codigo_adm = parametro
'                        rs_aux3!solicitud_codigo_adm = correldetalle
'                        rs_aux3!unidad_codigo = VAR_COD4
'                        rs_aux3!solicitud_codigo = VAR_SOL
'                        rs_aux3!edif_codigo = VAR_PROY2
'                        rs_aux3!beneficiario_codigo = VAR_BENEF
'                        rs_aux3!solicitud_tipo = Ado_datos.Recordset!solicitud_tipo       '"10"
'                        rs_aux3!venta_tipo = "E"
'                        rs_aux3!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant   'VAR_CITE
'                        rs_aux3!compra_fecha = Date
'                        rs_aux3!compra_descripcion = "COMPRA POR: " + lbl_titulo.Caption
'                        rs_aux3!compra_observaciones = "Edificio: " + Trim(dtc_desc3.Text)
'                        rs_aux3!compra_cantidad_total = 1   'Ado_datos.Recordset!venta_cantidad_total
'                        rs_aux3!compra_monto_bs = 0     'VAR_BS2
'                        rs_aux3!tipo_moneda = "BOB"
'                        rs_aux3!compra_monto_dol = 0        'VAR_DOL2
'                        rs_aux3!proceso_codigo = "TEC"
'                        rs_aux3!subproceso_codigo = "TEC-06"
'                        rs_aux3!etapa_codigo = "TEC-06-01"
'                        rs_aux3!clasif_codigo = "ADM"
'                        rs_aux3!doc_codigo = "R-114"
'                        rs_aux3!poa_codigo = "3.2.8"
'                        rs_aux3!estado_codigo_eqp = "REG"
'                        rs_aux3!estado_codigo = "REG"
'                        rs_aux3!usr_codigo = glusuario
'                        rs_aux3!fecha_registro = Date
'                        rs_aux3.Update
'
'                        'DETALLE Carga ao_ventas_detalle
'                        Set rstdestino = New ADODB.Recordset
'                        If rstdestino.State = 1 Then rstdestino.Close
'                        rstdestino.Open "select * from ao_compra_detalle  ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rstdestino.RecordCount > 0 Then
'                        End If
'                        Set rs_aux4 = New ADODB.Recordset
'                        If rs_aux4.State = 1 Then rs_aux4.Close
'                        'rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo= " & rs_aux3!compra_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
'                        rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo= " & VAR_SOL & "  and grupo_codigo = '30000' ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rs_aux4.RecordCount > 0 Then
'                            VAR_REG = 1
'                           rs_aux4.MoveFirst
'                           While Not rs_aux4.EOF
'                              If rs_aux4!grupo_codigo = "30000" Then
'                                db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, compra_codigo_det, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " & _
'                                "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", " & VAR_REG & ", '" & rs_aux4!bien_codigo & "', " & rs_aux4!bien_cantidad & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", '" & rs_aux3!compra_descripcion & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
'
'                                db.Execute "Update ao_compra_detalle SET ao_compra_detalle.compra_concepto  = ac_bienes.bien_descripcion From ao_compra_detalle INNER JOIN ac_bienes ON ao_compra_detalle.bien_codigo = ac_bienes.bien_codigo where ao_compra_detalle.compra_codigo = " & rs_aux3!compra_codigo & " and ao_compra_detalle.bien_codigo = '" & rs_aux4!bien_codigo & "' "
'                                VAR_REG = VAR_REG + 1
'                              End If
'                               rs_aux4.MoveNext
'                           Wend
'                        End If
'                        If rstdestino.State = 1 Then rstdestino.Close
'                    End If
'                    'WWWWWWWWWW
'        Set rs_aux2 = New ADODB.Recordset
'        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9 & "'  "
'        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'            Txt_campo1.Caption = rs_aux2!correl_doc
'            rs_aux2.Update
'        End If
'        rs_datos!doc_numero = Txt_campo1.Caption
'        'REVISAR !!! JQA 2014_07_08
'        'VAR_ARCH = RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
'        VAR_ARCH = "COM_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(Txt_campo1.Caption)))
'        rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
'        rs_datos!archivo_respaldo_cargado = "N"
'        rs_datos!estado_cotiza = "APR"
'        rs_datos!fecha_aprueba = Date
'        rs_datos!usr_codigo_aprueba = glusuario
'        rs_datos.UpdateBatch adAffectAll
'      End If
'
'  Else
'      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
End Sub

Private Sub BtnAprobar2_Click()
 NumComp = Ado_datos.Recordset!venta_codigo
 VAR_COBRANZA = Ado_datos16.Recordset!cobranza_prog_codigo
 VAR_PROY2 = Ado_datos.Recordset!edif_codigo
 VAR_EDIFC = Ado_datos.Recordset!edif_codigo_corto
 VAR_EMPRESA = Ado_datos.Recordset!codigo_empresa
 If VAR_COBRANZA > 1 Then
    'VERIFICA SI HAY CUOTAS ANTERIORES
    Set rs_aux22 = New ADODB.Recordset
    If rs_aux22.State = 1 Then rs_aux22.Close
    rs_aux22.Open "Select * from ao_ventas_cobranza_prog where venta_codigo= " & NumComp & " AND estado_codigo = 'REG' AND cobranza_prog_codigo < " & VAR_COBRANZA & " ", db, adOpenKeyset, adLockOptimistic
    If rs_aux22.RecordCount > 0 Then
        MsgBox "No se puede APROBAR, existen Cuotas anteriores NO Aprobadas, verifique los datos y vuelva a intentar ...", , "Atención"
        Exit Sub
    End If
 End If
 
 If IsNull(Ado_datos16.Recordset("cobranza_observaciones")) Or (Ado_datos16.Recordset("cobranza_programada_bs") = 0) Or Ado_datos16.Recordset!beneficiario_codigo_resp = "" Or IsNull(Ado_datos16.Recordset!beneficiario_codigo_resp) Then
    'If Ado_datos16.Recordset!beneficiario_codigo_resp = "" Or IsNull(Ado_datos16.Recordset!beneficiario_codigo_resp) Then
    MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
    Exit Sub
 Else
    If Ado_datos.Recordset("estado_codigo") = "REG" Then
        MsgBox "No se puede APROBAR el registro (Cronograma), previamente debe APROBAR la Venta (Cabecera) y vuelva a intentar ...", , "Atención"
        Exit Sub
    End If
    If Ado_datos.Recordset("estado_codigo") = "ANL" Then
        MsgBox "No se puede APROBAR el registro (Cronograma), porque este fue ANULADO(Cabecera) ...", , "Atención"
        Exit Sub
    End If
    'Plan de Cuotas
    VAR_DOL2 = Round(Ado_datos.Recordset!venta_monto_total_dol, 2)
    VAR_BS2 = Round(Ado_datos.Recordset!venta_monto_total_bs, 2)
    Set rs_aux21 = New ADODB.Recordset     'Plan de Cuotas
    If rs_aux21.State = 1 Then rs_aux21.Close
    rs_aux21.Open "Select SUM(cobranza_programada_bs) AS SumaCuota from ao_ventas_cobranza_prog WHERE venta_codigo = " & NumComp & " AND es_liquidacion = 'NO' ", db, adOpenStatic
    If rs_aux21.RecordCount > 0 Then
        If Round(VAR_BS2, 0) = Round(rs_aux21!SumaCuota, 0) Then        'Las Cuotas no igualan con el Total del contrato
        Else
            MsgBox "No se puede APROBAR, la SUMA de las Cuotas Bs, NO iguala con el TOTAL del Contrato Bs, verifique y vuelva a intentar ...", , "Atención"
            Exit Sub
        End If
    End If
    
    If Ado_datos16.Recordset("estado_codigo") = "REG" Then
       sino = MsgBox("Realizarás la solicitud de VARIAS cuotas en UNA sola FACTURA ? ", vbYesNo, "Confirmando")
       If sino = vbYes Then             'VARIAS CUOTAS PARA UNA SOLA FACTURA
            tw_ventas_cuotas_vs_fac.Show vbModal
       Else                             'UNA CUOTA PARA UNA FACTURA
            'SI ES CGI o CGE (Falta)
            ', edif_codigo_corto
            ', " & Ado_datos.Recordset!edif_codigo_corto & "
            nroventa = Ado_datos16.Recordset!venta_codigo
            db.Execute "update gc_documentos_respaldo set gc_documentos_respaldo.correl_doc = " & nroventa & " Where gc_documentos_respaldo.doc_codigo = '" & Ado_datos16.Recordset!doc_codigo & "' "
            'GRABA CABECERA DE FACTURACION NUEVA (ao_ventas_cobranza_fac)   'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
            'SE GENERAN CON LA FACTURA (dosifica_autorizacion, nro_factura, fecha_fac, codigo_control, archivo_foto, depto_codigo, Gestion, mes, edif_codigo_corto)
            db.Execute "INSERT INTO ao_ventas_cobranza_fac (ges_gestion, venta_codigo, doc_codigo_fac,              beneficiario_codigo_fac,                                beneficiario_nit,           glosa_Descripcion,                                  beneficiario_RazonSocial, nro_dui,      total_bs,                                       total_dol,                                      cambio_oficial, " & _
                        " Importe_ICE, Exportaciones_Exentas, Ventas_tasa_0, Subtotal_ICE, Descuentos_Bonos, Importe_Base_Debito_Fiscal,                    factura_87_bs,                                                      factura_87_dol,                                                 debito_fiscal_13_bs,                                                debito_fiscal_13_dol,                                               literal, " & _
                        " clasif_codigo, doc_codigo, doc_numero, factura_impresa, tipo_moneda, cta_codigo, cta_codigo2, correl_contab, estado_fac, estado_codigo_fac, estado_codigo,  " & _
                        " usr_codigo, fecha_registro, edif_codigo_corto, edif_codigo, codigo_empresa ) " & _
                " VALUES ('" & glGestion & "',  " & nroventa & ", '" & Ado_datos16.Recordset!doc_codigo_fac & "', '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & dtc_codigo2A.Text & "', '" & Ado_datos16.Recordset!cobranza_concepto_plazo & "', '" & dtc_desc2A.Text & "',  '0', " & Ado_datos16.Recordset!cobranza_total_bs & ",  " & Ado_datos16.Recordset!cobranza_total_dol & ",  " & GlTipoCambioOficial & ",  " & _
                        " '0',          '0',                    '0',            '0',            '0',    " & Ado_datos16.Recordset!cobranza_total_bs & ", " & Round(Ado_datos16.Recordset!cobranza_total_bs * 0.87, 2) & ", " & Round(Ado_datos16.Recordset!cobranza_total_dol * 0.87, 2) & ", " & Round(Ado_datos16.Recordset!cobranza_total_bs * 0.13, 2) & ", " & Round(Ado_datos16.Recordset!cobranza_total_dol * 0.13, 2) & ", '" & Ado_datos16.Recordset!Literal & "',  " & _
                        " 'ADM',        'R-103',        '0',        'N',            'BOB',      'NN',           'NN',        '0',            'REG',      'REG',          'REG',  " & _
                        " '" & glusuario & "', '" & CDate(Date) & "', " & VAR_EDIFC & ", '" & VAR_PROY2 & "', " & VAR_EMPRESA & "  ) "
                        
            'Actualiza CORREO ELECTRONICO
            db.Execute "UPDATE ao_ventas_cobranza_fac SET ao_ventas_cobranza_fac.beneficiario_email  = gc_beneficiario.beneficiario_email FROM ao_ventas_cobranza_fac INNER JOIN gc_beneficiario ON ao_ventas_cobranza_fac.beneficiario_codigo_fac = gc_beneficiario.beneficiario_codigo where ao_ventas_cobranza_fac.beneficiario_email Is Null "

            Set rs_aux20 = New ADODB.Recordset
            If rs_aux20.State = 1 Then rs_aux20.Close
            rs_aux20.Open "Select max(IdFactura) as Codigo3 from ao_ventas_cobranza_fac  ", db, adOpenKeyset, adLockOptimistic
            If IsNull(rs_aux20!codigo3) Then
               VAR_IDFAC = 1
            Else
               VAR_IDFAC = rs_aux20!codigo3
            End If
            'GRABA CABECERA DE LA FACTURA (QR)
            db.Execute "INSERT INTO ao_ventas_cobranza_fac_QR (IdFactura, archivo_foto_cargado, estado_codigo, usr_codigo, fecha_registro ) " & _
                " VALUES ('" & VAR_IDFAC & "',  'N',            'REG',   '" & glusuario & "', '" & CDate(Date) & "' ) "

            'GRABA DETALLE DE FACTURACION NUEVA (ao_ventas_cobranza)
            db.Execute "INSERT INTO ao_ventas_cobranza (ges_gestion, cobranza_prog_codigo, venta_codigo,                                    beneficiario_codigo,                                    beneficiario_codigo_fac,                            beneficiario_codigo_resp,                               cobranza_programada_bs,                                 cobranza_programada_dol,                                cobranza_solicitado_bs,                                  cobranza_solicitado_dol,                 cobranza_descuento_bs, cobranza_descuento_dol, cobranza_total_bs,         cobranza_total_dol,                                     Literal,    cobranza_fecha_prog,                              cobranza_fecha_cobro, cobranza_observaciones, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, doc_codigo_fac, cobranza_nro_factura, cobranza_nro_autorizacion, poa_codigo,  " & _
            " estado_codigo, usr_codigo, fecha_registro, cobranza_fecha_sol, estado_codigo_sol, estado_codigo_fac, venta_codigo_new) " & _
            " VALUES ('" & glGestion & "', " & Ado_datos16.Recordset!cobranza_prog_codigo & ", " & nroventa & ", '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & Ado_datos16.Recordset!beneficiario_codigo_resp & "', " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", '0', '0', " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", '" & Ado_datos16.Recordset!Literal & "', '" & Ado_datos16.Recordset!cobranza_fecha_prog & "', '" & Ado_datos16.Recordset!cobranza_fecha_cobro & "', '" & Ado_datos16.Recordset!cobranza_concepto_plazo & "', 'FIN', 'FIN-02', 'FIN-02-02', 'ADM', 'R-105', '0', 'R-101', '0', '0', '3.1.2',  " & _
            " 'REG', '" & glusuario & "', '" & Date & "', '" & Date & "', 'APR', 'REG', " & VAR_IDFAC & " )"

            ' APRUEBA ao_ventas_cobranza_prog
            'db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'APR' Where venta_codigo = " & nroventa & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "
            db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'APR', fecha_registro= '" & Ado_datos16.Recordset!fecha_registro & "' Where venta_codigo = " & nroventa & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "
            ' Actualiza CODIGO_COBRNAZA en el cronogrma
            db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.cobranza_codigo = ao_ventas_cobranza.cobranza_codigo from ao_ventas_cobranza_prog INNER JOIN ao_ventas_cobranza " & _
            " ON ao_ventas_cobranza_prog.venta_codigo = ao_ventas_cobranza.venta_codigo and ao_ventas_cobranza_prog.cobranza_prog_codigo = ao_ventas_cobranza.cobranza_prog_codigo WHERE (ao_ventas_cobranza_prog.venta_codigo = " & nroventa & " and ao_ventas_cobranza_prog.cobranza_prog_codigo=" & Ado_datos16.Recordset!cobranza_prog_codigo & " )"

            db.Execute "update ao_ventas_cobranza_prog SET Gestion = YEAR(cobranza_fecha_prog) Where venta_codigo = " & nroventa & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "

            db.Execute "update ao_ventas_cobranza_prog SET cobranza_mes = MONTH(cobranza_fecha_prog) Where venta_codigo = " & nroventa & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "
            'Call ABRIR_TABLAS_AUX
            sino = MsgBox("Deseas IMPRIMIR la Solicitud de FACTURA de la CUOTA elegida ? ", vbYesNo, "Confirmando")
            If sino = vbYes Then             'IMPRIME FACTURA
                  Dim iResult As Variant  ', i%, y%
                  CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
                  CryR01.WindowShowRefreshBtn = True
                  CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
                  CryR01.StoredProcParam(2) = Me.Ado_datos16.Recordset!cobranza_prog_codigo
                  'Literal por el Total
                  var_literal = Literal(CStr(Ado_datos.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
                  CryR01.Formulas(1) = "literalcobro = '" & var_literal & "' "
                  CryR01.Formulas(2) = "correlcobro = '" & Ado_datos16.Recordset!cobranza_prog_codigo & "' "
                  iResult = CryR01.PrintReport
                  If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
            'Else                             'NO IMPRIME FACTURA
            End If
            db.Execute "tp_actualiza_datos_venta " & NumComp
            MsgBox "Se APROBOBO la Cuota y se Envió satisfactoriamente la Solicitud de FACTURA ...", , "Atención"
            Call ABRIR_DETALLE
            If (DtgCobro.SelBookmarks.Count <> 0) Then
                DtgCobro.SelBookmarks.Remove 0
            End If
            If Ado_datos16.Recordset.RecordCount > 0 Then
             'VAR_SW = ""
                rs_datos16.Find "cobranza_prog_codigo = " & VAR_COBRANZA & "   ", , , 1
                DtgCobro.SelBookmarks.Add (rs_datos16.Bookmark)
        '        Set Ado_datos.Recordset = rs_datos.DataSource
        '        Set dg_datos.DataSource = Ado_datos.Recordset
            Else
            'VAR_SW = ""
               rs_datos16.MoveLast
            End If
            'Ado_datos16.Refresh
       End If
    End If
 End If
End Sub

Private Sub BtnAprobar3_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
    If Ado_datos.Recordset.RecordCount > 0 And (glusuario = "RCUELA" Or glusuario = "ADMIN" Or glusuario = "SQUISPE" Or glusuario = "CSALINAS") Then
        If CDate(Format(DTPfechasol.Value, "dd/mm/yyyy")) = "01/01/1900" Or DTPfechasol.Value = "" Then
            MsgBox "Debe registrar la Fecha de Venta !! , Verifique y vuelva a Intentar ...", vbExclamation, "Atención"
            VAR_VAL = "ERR"
            Exit Sub
        End If

        correlv = Ado_datos.Recordset!venta_codigo
        GlEdificio = Ado_datos.Recordset!edif_codigo
        'VALIDA EDIFICIO Y EQUIPOS
        Set rs_aux10 = New ADODB.Recordset     'Proyecto de Edificación
        If rs_aux10.State = 1 Then rs_aux10.Close
        rs_aux10.Open "Select * from gc_edificaciones WHERE edif_codigo = '" & dtc_codigo3.Text & "' and estado_codigo = 'APR' ", db, adOpenStatic
        If rs_aux10.RecordCount = 0 Then
            'Si Faltarian Aprobar
            MsgBox "No se puede CONTABILIZAR, verifique los datos del Edificio si estan correctos y si está Aprobado, luego vuelva a intentar ...", , "Atención"
            Exit Sub
        End If
        
        Set rs_aux11 = New ADODB.Recordset     'Equipos de Venta_Detalle
        If rs_aux11.State = 1 Then rs_aux11.Close
        rs_aux11.Open "Select * from mv_bienes_vs_venta_det WHERE venta_codigo = " & correlv & "  ", db, adOpenStatic
        If rs_aux11.RecordCount > 0 Then
            'Si Faltarian Aprobar
            MsgBox "No se puede CONTABILIZAR, verifique los datos de los EQUIPOS y si estos están Aprobados, luego vuelva a intentar ...", , "Atención"
            Exit Sub
        End If
        
        Set rs_aux12 = New ADODB.Recordset     'Partidas de Venta_Detalle
        If rs_aux12.State = 1 Then rs_aux12.Close
        rs_aux12.Open "Select * from ao_ventas_detalle WHERE venta_codigo = " & correlv & " and par_codigo=''  ", db, adOpenStatic
        If rs_aux12.RecordCount > 0 Then
            'Si Faltarian Partida
            MsgBox "No se puede CONTABILIZAR, verifique los datos de Detalle de Bienes , luego vuelva a intentar ...", , "Atención"
            Exit Sub
        End If
        'rs_aux18
        Set rs_aux18 = New ADODB.Recordset     'Alcance del Contrato
        If rs_aux18.State = 1 Then rs_aux18.Close
        rs_aux18.Open "Select * from ao_ventas_alcance WHERE venta_codigo = " & correlv & "  ", db, adOpenStatic
        If rs_aux18.RecordCount < 6 Then
            'Si Faltarian Partida
            MsgBox "No se puede CONTABILIZAR, verifique los datos del Alcance del Contrato , luego vuelva a intentar ...", , "Atención"
            Exit Sub
        End If
        If Ado_datos.Recordset!estado_contab = "REG" Then
           sino = MsgBox("Esta seguro de CONTABILIZAR el registro?", vbYesNo, "Confirmando")
           If sino = vbYes Then
               ' CONTABILIZA ao_ventas_cabecera
               ' AQUIIIIIIIIIIIIIIIIIIIIIIIIIII
               Ado_datos.Recordset!estado_contab = "APR"
               Ado_datos.Recordset.Update
            Else
            End If
        End If
    End If
End Sub

Private Sub BtnBuscar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
'    'JQA
'    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
'    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
'      buscados = 1
'      PosibleApliqueFiltro = False
'      Dim rsNada As ADODB.Recordset
'      Dim GrSqlAux As String
'      Set ClBuscaGrid = New ClBuscaEnGridExterno
'      Set ClBuscaGrid.Conexión = db
'      ClBuscaGrid.EsTdbGrid = False
'      Set ClBuscaGrid.GridTrabajo = dg_datos
'      ClBuscaGrid.QueryUtilizado = queryinicial
'      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
'      ClBuscaGrid.CamposVisibles = "110"
'      ClBuscaGrid.Ejecutar
'      PosibleApliqueFiltro = True
    buscados = 1
    OptFilGral2.Visible = False
    OptFilGral1.Visible = False
    Call OptFilGral2_Click
    Call ABRIR_DETALLE
    PosibleApliqueFiltro = False
    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexión = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = dg_datos
    ClBuscaGrid.QueryUtilizado = queryinicial
    Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
    ClBuscaGrid.CamposVisibles = "110"
    ClBuscaGrid.Ejecutar
    PosibleApliqueFiltro = True
  Else
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
    OptFilGral1.Visible = True
    OptFilGral2.Visible = True
  End If
End Sub

Private Sub BtnCancelar_Click()
  'Ado_datos.Refresh
  fraOpciones.Visible = True
  FraGrabarCancelar.Visible = False
  marca1 = Ado_datos.Recordset.Bookmark
'  If Ado_datos.Recordset("estado_codigo") = "REG" Then
'    Call OptFilGral2_Click
'  Else
'    Call OptFilGral1_Click
'  End If
  If OptFilGral1.Value = True Then
    Call OptFilGral1_Click
  Else
    Call OptFilGral2_Click
  End If
  FraNavega.Enabled = True
  FrmCabecera.Enabled = False
  Fra_datos.Enabled = True
  FrmDetalle.Visible = True
  FrmCobranza.Visible = True
  Fra_Total.Visible = True
  dg_datos.Visible = True
  FrmABMDet.Visible = True
  FrmABMDet2.Visible = True
  BtnImprimir2.Visible = True
  SSTab1.Tab = 0
  SSTab1.TabEnabled(0) = True
  SSTab1.TabEnabled(1) = True
  SSTab1.TabEnabled(2) = True
  'Ado_datos.Recordset.Move marca1 - 1
  dtc_desc2.backColor = &HC0C0C0
  Text11.Visible = True
End Sub

Private Sub BtnCancelar2_Click()
    FraAnula.Visible = False
End Sub

Private Sub BtnCancelarBen_Click()
    frm_benef.Visible = False
End Sub

Private Sub btnEliminar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo <> "ANL" Then      'And Ado_datos.Recordset!estado_almacen <> "APR"
       NumComp = Ado_datos.Recordset!venta_codigo
       If ExisteReg(Ado_datos.Recordset!venta_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado Consulte en Cobranzas y/o Facturación ...", vbInformation + vbOKOnly, "Atención": Exit Sub
       sino = MsgBox("Esta seguro de ANULAR la venta registrada ?", vbYesNo, "Confirmando")
       If sino = vbYes Then
          db.Execute "update ao_ventas_cabecera set estado_codigo = 'ANL', estado_almacen = 'ANL' Where venta_codigo = " & NumComp & "  "
          db.Execute "update ao_ventas_cabecera set estado_cancelado = 'A' Where venta_codigo = " & NumComp & "  "
          db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'ANL' Where venta_codigo = " & NumComp & "  "
          db.Execute "update ao_solicitud set estado_codigo = 'ANL' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' AND solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
          marca1 = Ado_datos.Recordset.Bookmark
          'Ado_datos.Recordset.Requery
          'Ado_datos.Refresh
          Call OptFilGral1_Click
          Ado_datos.Recordset.Move marca1 - 1
       End If
    Else
      MsgBox "NO se puede ANULAR el registro que ya fue Aprobado o previamente ANULADO...", , "Atencion"
    End If
  Else
    MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
End Sub

Private Function ExisteReg(Codigo As Integer) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_ventas_cobranza_prog WHERE  venta_codigo= " & Codigo & " and estado_codigo = 'APR'   "
    '    <> 'ANL'
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub valida_campos()
  If (dtc_codigo8.Text = "" Or dtc_codigo8.Text = "0") Then
    MsgBox "Debe Elejir la Empresa... Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo11 = "" Then
    MsgBox "Debe Elejir el Tipo de Venta!! (Credito, pago ne Efectivo, etc.), Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo2 = "" Then
    MsgBox "Debe Registrar el Cliente para la Venta. Consulte con el Administrador del Sistema ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If CDate(Format(DTPfechaFin.Value, "dd/mm/yyyy")) = "01/01/1900" Or DTPfechaFin.Value = "" Then
      MsgBox "Debe registrar la Fecha de Inicio !! , Verifique y vuelva a Intentar ...", vbExclamation, "Atención"
      VAR_VAL = "ERR"
      Exit Sub
  End If
  If cmb_mes_ini = "" Then
    MsgBox "Debe Elejir el " + lbl_mes_ini + ", Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If parametro = "DNMAN" Then
    If CDate(Format(DTPfechaFin.Value, "dd/mm/yyyy")) <= CDate(Format(DTPfechaIni.Value, "dd/mm/yyyy")) Then
      MsgBox "La Fecha de Inicio debe ser MENOR a la Fecha de Fin del Contrato!! , Vuelva a Intentar ...", vbExclamation, "Atención"
      VAR_VAL = "ERR"
      Exit Sub
    End If
  Else
    If CDate(Format(DTPfechaFin.Value, "dd/mm/yyyy")) < CDate(Format(DTPfechaIni.Value, "dd/mm/yyyy")) Then
      MsgBox "La Fecha de Inicio debe ser MENOR o IGUAL a la Fecha de Fin del Contrato!! , Vuelva a Intentar ...", vbExclamation, "Atención"
      VAR_VAL = "ERR"
      Exit Sub
    End If
  End If
  Select Case RTrim(cmb_mes_ini)
        Case "ENERO"
            VAR_MES2 = 1
        Case "FEBRERO"
            VAR_MES2 = 2
        Case "MARZO"
            VAR_MES2 = 3
        Case "ABRIL"
            VAR_MES2 = 4
        Case "MAYO"
            VAR_MES2 = 5
        Case "JUNIO"
            VAR_MES2 = 6
        Case "JULIO"
            VAR_MES2 = 7
        Case "AGOSTO"
            VAR_MES2 = 8
        Case "SEPTIEMBRE"
            VAR_MES2 = 9
        Case "OCTUBRE"
            VAR_MES2 = 10
        Case "NOVIEMBRE"
            VAR_MES2 = 11
        Case "DICIEMBRE"
            VAR_MES2 = 12
  End Select
'  If Month(CDate(Format(DTPFechaIni.Value, "dd/mm/yyyy"))) <> 12 And VAR_MES2 <> 1 Then
'    If Val(VAR_MES2) < Month(CDate(Format(DTPFechaIni.Value, "dd/mm/yyyy"))) Then
'        MsgBox "El MES de Inicio de Cobranza NO puede ser MENOR al de la Fecha de Inicio del Contrato!! , Vuelva a Intentar ...", vbExclamation, "Atención"
'        VAR_VAL = "ERR"
'        Exit Sub
'    End If
'  End If
  If Month(CDate(Format(DTPfechaIni.Value, "dd/mm/yyyy"))) <> VAR_MES2 Then
    'If Val(VAR_MES2) < Month(CDate(Format(DTPfechaIni.Value, "dd/mm/yyyy"))) Then
        MsgBox "El 'MES Inicio del Plan de Cuotas' NO puede ser DIFERENTE al MES de la Fecha de Inicio del Contrato!! , Vuelva a Intentar ...", vbExclamation, "Atención"
        VAR_VAL = "ERR"
        Exit Sub
    'End If
  End If
    'DTPFechaIni
    'DTPFechaFin
    'meses = DateDiff("m", Text1.Text, Text2.Text)
    'txtCantCobr
    CONT4 = DateDiff("m", DTPfechaIni.Value, DTPfechaFin.Value)
  If (txtCantCobr.Text <> CONT4 + 1) And (cmd_unimed2.Text = "MES") Then
     sino = MsgBox("El 'Número de Cuotas' es DIFERENTE al número de meses de la Fecha de INICIO y FIN, aún así desea continuar ??...", vbYesNo + vbQuestion, "Atención ...")
     If sino = vbYes Then
     Else
        VAR_VAL = "ERR"
        Exit Sub
     End If
  End If
  'FALTA VERIFICAR SI EXISTE EN ORGANIZACION DE ZONAS...
  If Ado_datos.Recordset!unidad_codigo = "DNMAN" Or Ado_datos.Recordset!unidad_codigo = "DMANS" Or Ado_datos.Recordset!unidad_codigo = "DMANB" Or Ado_datos.Recordset!unidad_codigo = "DMANC" Then
    If dtc_codigo7.Text = "" Or dtc_codigo7.Text = "0" Then
      MsgBox "En pestaña DATOS.CRONOGRAMA.MTTO.; Debe Elejir Zona Piloto !! , Vuelva a Intentar por favor ...", vbExclamation, "Atención"
      VAR_VALD = "ERR"
      Exit Sub
    End If

    If dtc_codigo4 = "" Then
      MsgBox "Debe Elejir Responsable del Servicio Técnico:, Vuelva a Intentar ...", vbExclamation, "Atención"
      VAR_VAL = "ERR"
      Exit Sub
    End If

  End If
'  If dtc_codigo11.Text = "C" And dtc_codigo2 = "VD" Then
'        MsgBox "NO se puede realizar la Venta a Credito, Debe cambiar de Cliente ..."
'  Else

End Sub

Private Sub BtnEliminar2_Click()
 If Ado_datos.Recordset!estado_codigo = "REG" Then
  Set rs_datos6 = New ADODB.Recordset
  If rs_datos6.State = 1 Then rs_datos6.Close
  rs_datos6.Open "Select * from to_cronograma WHERE estado_detalle = 'APR' AND unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "' and tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & "   ", db, adOpenStatic
  If rs_datos6.RecordCount > 0 Then
     ProgressBar1.Visible = True
     With ProgressBar1
        .Max = rs_datos6.RecordCount
        .Min = 0
        .Value = 0
     End With
     ProgressBar1.Value = ProgressBar1.Value + 1
       db.Execute "Update to_cronograma Set estado_detalle = 'REG' Where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "' and tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & "   "
       db.Execute "update to_cronograma_diario set bien_codigo = '', unidad_codigo_tec = '',  tec_plan_codigo = '', observaciones = 'HORARIO LABORABLE', edif_descripcion = '', estado_activo = 'REG', estado_codigo = 'REG' Where unidad_codigo_tec = '" & Ado_datos.Recordset!unidad_codigo_tec & "' and tec_plan_codigo = " & Ado_datos.Recordset!tec_plan_codigo & "   "     'WHERE fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & "  "
       MsgBox "El Cronograma fue deshabilitado exitosamente ...", vbExclamation, "Validación de Registro"
       Call ABRIR_DETALLE
       ProgressBar1.Visible = False
  Else
        ProgressBar1.Visible = False
        MsgBox "El Cronograma ya fue deshabilitado, verifique y vuelva a generarlo (Nuevo) ...", vbExclamation, "Validación de Registro"
  End If
 Else
        MsgBox "NO se puede ANULAR EL CRONOGRAMA, en un Registro APROBADO o ANULADO !! ", vbExclamation, "Atención!"
 End If

End Sub

Private Sub BtnGrabar_Click()
  VAR_VAL = "OK"
  Call valida_campos
  VAR_VALD = "OK"
  Call valida_campos2
  If VAR_VAL = "OK" And VAR_VALD = "OK" Then
    NumComp = Ado_datos.Recordset!venta_codigo
    FInicio = IIf(DTPfechaIni.Value = "", Format(Date, "dd,mm,yyyy"), DTPfechaIni.Value)            'Ado_datos.Recordset!venta_fecha_inicio
    FFin = IIf(DTPfechaFin.Value = "", Format(Date, "dd,mm,yyyy"), DTPfechaFin.Value)
    CANTOT = Ado_datos.Recordset!venta_cantidad_total
    gestion0 = glGestion        'Ado_datos.Recordset("ges_gestion")
    VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
    corrprog = Ado_datos.Recordset!correl_cobro_prog
    VAR_MED = Ado_datos.Recordset!unimed_codigo
    FrmCabecera.Enabled = False
    Call grabar
    CONT1 = Ado_datos.Recordset!venta_cantidad_cobr
'    'CREA VENTA CABECERA
'    Set rs_aux3 = New ADODB.Recordset
'    If rs_aux3.State = 1 Then rs_aux3.Close
'    rs_aux3.Open "Select max(cobranza_prog_codigo) as Codigo3 from ao_ventas_cobranza_prog where venta_codigo= " & NumComp & " ", db, adOpenStatic
'    'If rs_aux3.RecordCount > 0 Then
'    If IsNull(rs_aux3!codigo3) Then
'        db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & NumComp & " "
'        corrprog = 0
'        Call CRONO2
'    Else
'        sino = MsgBox("El Cronograma ya existe, desea volver a Generarlo ? (los items Aprobados no serán modificados)...", vbYesNo + vbQuestion, "Atención ...")
'        If sino = vbYes Then
'            'OJO BORRAR ao_ventas_cobranza_prog
'            'db.Execute "DELETE ao_ventas_cobranza_prog where venta_codigo= " & NumComp & " and estado_codigo = 'REG' AND estado_ac <> 'APR' "
'            db.Execute "DELETE ao_ventas_cobranza_prog where venta_codigo= " & NumComp & " and estado_codigo = 'REG' "
'            db.Execute "update ao_ventas_cobranza_prog set venta_codigo_new = cobranza_prog_codigo where venta_codigo= " & NumComp & " "
'            db.Execute "update ao_ventas_cobranza_prog set cobranza_prog_codigo = venta_codigo_new + 100 where venta_codigo= " & NumComp & " "
'            'db.Execute "update ao_ventas_cobranza set ao_ventas_cobranza.cobranza_prog_codigo = ao_ventas_cobranza_prog.cobranza_prog_codigo FROM ao_ventas_cobranza INNER JOIN ao_ventas_cobranza_prog ON ao_ventas_cobranza.venta_codigo= ao_ventas_cobranza_prog.venta_codigo AND ao_ventas_cobranza.cobranza_prog_codigo= ao_ventas_cobranza_prog.venta_codigo_new "
'            db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & NumComp & " "
'            db.Execute "update ao_ventas_cabecera set tipo_moneda = 'BOB' where venta_codigo= " & NumComp & " "
'            corrprog = 0
'            Call CRONO2
'        Else
'        'If rs_aux3!codigo3 > corrprog Then
'            'ACTUALIZAR CORRELATIVO CRONO
'            db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '" & rs_aux3!codigo3 & "' where venta_codigo= " & NumComp & " "
'            corrprog = rs_aux3!codigo3
'        'End If
'        End If
'    End If
    db.Execute "update ao_ventas_cobranza set ao_ventas_cobranza.cobranza_prog_codigo = ao_ventas_cobranza_prog.cobranza_prog_codigo FROM ao_ventas_cobranza INNER JOIN ao_ventas_cobranza_prog ON ao_ventas_cobranza.venta_codigo= ao_ventas_cobranza_prog.venta_codigo AND ao_ventas_cobranza.cobranza_codigo = ao_ventas_cobranza_prog.cobranza_codigo WHERE ao_ventas_cobranza.venta_codigo = " & NumComp & " "
'    db.Execute "update ao_ventas_detalle set venta_det_cantidad = " & txtCantCobr.Text & " where venta_codigo = " & NumComp & " AND par_codigo = 43340"
'    db.Execute "update ao_ventas_detalle set venta_precio_total_bs = (venta_precio_unitario_bs * venta_det_cantidad), venta_precio_total_dol = (venta_precio_unitario_dol * venta_det_cantidad) where venta_codigo = " & NumComp & " AND par_codigo = 43340"
'    db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.unimed_codigo = ao_ventas_cabecera.unimed_codigo_cobr,ao_ventas_cabecera.venta_monto_total_bs = aa_datos_venta_cabecera.tot_bs,ao_ventas_cabecera.venta_monto_total_dol = aa_datos_venta_cabecera.tot_dol,ao_ventas_cabecera.venta_cantidad_total = aa_datos_venta_cabecera.COBRANZA,ao_ventas_cabecera.venta_saldo_p_cobrar_bs = aa_datos_venta_cabecera.tot_bs - ao_ventas_cabecera.venta_monto_cobrado_bs From aa_datos_venta_cabecera WHERE aa_datos_venta_cabecera.venta_codigo = ao_ventas_cabecera.venta_codigo and ao_ventas_cabecera.venta_codigo = " & NumComp & ""
    db.Execute "tp_actualiza_datos_venta " & NumComp
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FraNavega.Enabled = True
    FrmCabecera.Enabled = False
    Fra_datos.Enabled = True
    dg_datos.Visible = True
    FrmDetalle.Visible = True
    FrmCobranza.Visible = True
    Fra_Total.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
    BtnImprimir2.Visible = True
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    
    dtc_desc2.backColor = &HC0C0C0
    Text11.Visible = True
  
    Set rs_aux13 = New ADODB.Recordset
    If rs_aux13.State = 1 Then rs_aux13.Close
    rs_aux13.Open "select sum(venta_precio_total_bs) as total from ao_ventas_detalle where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "' ", db, adOpenKeyset, adLockOptimistic
    If rs_aux13!total <> "NULL" Then
        'Ado_datos.Recordset!literal_a = Literal(rs_aux13!total)
        var_literal = Literal(rs_aux13!total)
        db.Execute "update ao_ventas_cabecera set literal_a = '" & var_literal & "' where venta_codigo= " & NumComp & " "
    End If
  
     'Ado_datos.Recordset.Update
     If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
     'VAR_SW = ""
        rs_datos.Find "venta_codigo = " & NumComp & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
     'VAR_SW = ""
        rs_datos.MoveLast
     End If
    
    'If OptFilGral1.Value = True Then
    '    Call OptFilGral1_Click
    'Else
    '    Call OptFilGral2_Click
    'End If
  End If
End Sub

Private Sub CRONO2()
    Set rs_aux5 = New ADODB.Recordset
    If rs_aux5.State = 1 Then rs_aux5.Close
    rs_aux5.Open "select * from ao_ventas_cabecera where venta_codigo= " & NumComp & "  ", db, adOpenKeyset, adLockBatchOptimistic
    'Set AdoAux.Recordset = rsAuxDetalle
    If rs_aux5.RecordCount > 0 Then
      CONT2 = 1
      FInicio = rs_aux5!venta_fecha_inicio
'      CANTOT = rs_aux5!venta_cantidad_total
      gestion0 = Ado_datos.Recordset!ges_gestion
      VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
      VAR_MED2 = Ado_datos.Recordset!unimed_codigo_cobr
      VAR_COBR2 = Ado_datos.Recordset!venta_cantidad_cobr
      MControl = Ado_datos.Recordset!mes_inicio_crono
      Select Case RTrim(MControl)
        Case "ENERO"
            VAR_MES2 = 1
        Case "FEBRERO"
            VAR_MES2 = 2
        Case "MARZO"
            VAR_MES2 = 3
        Case "ABRIL"
            VAR_MES2 = 4
        Case "MAYO"
            VAR_MES2 = 5
        Case "JUNIO"
            VAR_MES2 = 6
        Case "JULIO"
            VAR_MES2 = 7
        Case "AGOSTO"
            VAR_MES2 = 8
        Case "SEPTIEMBRE"
            VAR_MES2 = 9
        Case "OCTUBRE"
            VAR_MES2 = 10
        Case "NOVIEMBRE"
            VAR_MES2 = 11
        Case "DICIEMBRE"
            VAR_MES2 = 12
      End Select
      FControl = "01/" + Str(Month(FInicio)) + "/" + Str(Year(FInicio))
      CONT3 = 0
      CONT4 = 0
      Select Case VAR_MED2
        Case "MES"
            CONT_MED = 1
        Case "BMES"
            CONT_MED = 2
        Case "TMES"
            CONT_MED = 3
        Case "CMES"
            CONT_MED = 4
        Case "5MES"
            CONT_MED = 5
        Case "SMES"
            CONT_MED = 6
        Case "7MES"
            CONT_MED = 7
        Case "8MES"
            CONT_MED = 8
        Case "9MES"
            CONT_MED = 9
        Case "10MES"
            CONT_MED = 10
        Case "11MES"
            CONT_MED = 11
        Case "ANUAL"
            CONT_MED = 12
      End Select
      'fanio = Year(FControl)
      'fmes = Month(FControl)
      While (CONT2 <= VAR_COBR2)
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from ao_ventas_cobranza_prog where venta_codigo = '" & NumComp & "'  ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 And corrprog >= VAR_COBR2 Then
            MsgBox "El Cronograma ya fue generado... ", , "Atención"
            CONT2 = CONT2 + 1
        Else
           'wwwwwwwwwwwwwwwwwwwwww
          correldet2 = rs_aux5!correl_cobro_prog + 1
          rs_aux5!correl_cobro_prog = rs_aux5!correl_cobro_prog + 1
          corrprog = correldet2
          rs_aux5.Update
          Set rs_aux8 = New ADODB.Recordset
          If rs_aux8.State = 1 Then rs_aux8.Close
          'rs_aux8.Open "select * from ao_ventas_cobranza_prog where venta_codigo = '" & NumComp & "' and cobranza_prog_codigo =" & correldet2 & "  ", db, adOpenKeyset, adLockOptimistic
          
          rs_aux8.Open "select * from ao_ventas_cobranza_prog where venta_codigo = " & NumComp & " and YEAR(cobranza_fecha_prog) ='" & Year(FControl) & "'  AND MONTH(cobranza_fecha_prog) = '" & Month(FControl) & "'  ", db, adOpenKeyset, adLockOptimistic
          If rs_aux8.RecordCount = 0 Then
            rs_aux2.AddNew
            'gestion0 = Ado_datos.Recordset!ges_gestion
            'VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
            rs_aux2!ges_gestion = gestion0
            rs_aux2!venta_codigo = NumComp 'Ado_datos.Recordset("venta_codigo")
            rs_aux2!cobranza_prog_codigo = correldet2
            rs_aux2!beneficiario_codigo = VAR_BENEF                   'Codigo Beneficiario/Cliente
            'OJO MODIFICAR COBRADOR - JQA 03-ENE-2015
            rs_aux2!beneficiario_codigo_resp = IIf(dtc_codigo5.Text = "", "4333735", dtc_codigo5.Text) '                                                     'Codigo Cobrador
            'rs_aux2!nombre_cobrador = dtc_desc4A.Text   '+ " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
            Set rs_aux6 = New ADODB.Recordset
            If rs_aux6.State = 1 Then rs_aux6.Close
            If VAR_UORIGEN = "DNMAN" Then
                rs_aux6.Open "select sum(venta_precio_unitario_bs) as acumBs from ao_ventas_detalle where venta_codigo = " & NumComp & " AND (par_codigo = '99990' or par_codigo = '43340') ", db, adOpenKeyset, adLockReadOnly
                If rs_aux6.RecordCount > 0 Then
                    rs_aux2!cobranza_programada_bs = Round(rs_aux6!acumBs, 2)                    'Monto Programado Bs CONT1
                    rs_aux2!cobranza_total_bs = Round(rs_aux6!acumBs, 2)                         'Monto Total Bs
                Else
                    rs_aux2!cobranza_programada_bs = 0
                End If
            Else
                If CONT1 = "" Then
                    CONT1 = "1"
                End If
                If CONT1 = "0" Then CONT1 = "1"
                rs_aux6.Open "select sum(venta_precio_total_bs) as acumBs from ao_ventas_detalle where venta_codigo = " & NumComp & " AND (par_codigo <> '43340') ", db, adOpenKeyset, adLockReadOnly
                If rs_aux6.RecordCount > 0 Then
                    rs_aux2!cobranza_programada_bs = Round(rs_aux6!acumBs / CONT1, 2)                                   'Monto Programado Bs
                    rs_aux2!cobranza_programada_dol = Round(rs_aux2!cobranza_programada_bs / GlTipoCambioMercado, 2)    'Monto Programado en Dolares
                    rs_aux2!cobranza_total_bs = Round(rs_aux6!acumBs / CONT1, 2)                                        'Monto Total Bs
                    rs_aux2!cobranza_total_dol = Round(rs_aux2!cobranza_total_bs / GlTipoCambioMercado, 2)              'Monto Total Dol
                Else
                    rs_aux2!cobranza_programada_bs = 0
                    rs_aux2!cobranza_programada_dol = 0
                    rs_aux2!cobranza_total_bs = 0
                    rs_aux2!cobranza_total_dol = 0
                End If
            End If
            
            'rs_aux2!cobranza_programada_dol = rs_aux6!acumBs / GlTipoCambioMercado  'Monto Programado en Dolares
            rs_aux2!cobranza_descuento_bs = 0                                       'Descuento Bs
            rs_aux2!cobranza_descuento_dol = 0                                      'Descuento Dol
            
            'rs_aux2!cobranza_total_dol = rs_aux6!acumBs / GlTipoCambioMercado       'Monto Total Dol
            'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW=
            'rs_aux2!cobranza_fecha_prog = rs_aux5!venta_fecha_inicio + (30 * CONT2)
            'Dim fdia, fmes, fanio As Integer
            'fmes = Month(FInicio) + CONT2
            
            fdia = Day(FControl)
            fanio = Year(FControl)
            'CONT3 = CONT2 * CONT_MED
            CONT3 = 1
            While (CONT3 <= CONT_MED)
                fmes = Month(FControl)
                Select Case fmes
                    Case 2
                        If fanio = "2012" Or fanio = "2016" Or fanio = "2020" Or fanio = "2024" Then
                            Dias_Mes = 29
                        Else
                            Dias_Mes = 28
                            'Dias_Del_Mes = IIf(saltarYear(Fecha), 29, 28)
                        End If
                    Case 1, 3, 5, 7, 8, 10, 12
                        Dias_Mes = 31
                    Case 4, 6, 9, 11
                        Dias_Mes = 30
                End Select
                If Val(VAR_MES2) = Month(FControl) Then
                    rs_aux2!cobranza_fecha_prog = FControl
                    'rs_aux2!cobranza_fecha_conformidad = FControl + 10
                    rs_aux2!cobranza_fecha_cobro = FControl + 20
                    VAR_MES2 = VAR_MES2 + CONT_MED
                    If Val(VAR_MES2) > 12 Then
                        VAR_MES2 = Val(VAR_MES2) - 12
                    End If
                End If
                FControl = FControl + Dias_Mes
                CONT3 = CONT3 + 1
                CONT4 = CONT4 + Dias_Mes
            Wend
            'FControl = Str(fdia) + "/" + Str(fmes) + "/" + Str(fanio)
            'rs_aux2!cobranza_fecha_prog = FInicio + (30 * CONT2)
            'rs_aux2!cobranza_fecha_prog = FControl
            If rs_aux2!cobranza_fecha_prog = Null Then
                rs_aux2!cobranza_fecha_prog = Date
            End If
            rs_aux2!gestion = Year(rs_aux2!cobranza_fecha_prog)
            rs_aux2!cobranza_mes = Month(rs_aux2!cobranza_fecha_prog)
            
            'VAR_FEC2 = MonthName(Month(IIf(IsNull(rs_aux2!cobranza_fecha_prog), Date, rs_aux2!cobranza_fecha_prog)))
            
            VAR_FEC2 = MonthName(Month(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog)))
            'rs_aux2!cobranza_fecha_cobro = FControl + 10 ' rs_aux2!cobranza_fecha_prog + 10
            'If VAR_MED2 = "MES" Then
            '    FControl = FControl + Dias_Mes
            'End If
            'rs_aux2!cobranza_observaciones = "CUOTA Nro. " + Str(corrprog) + " - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(Date)) + " - " + lbl_titulo
            'rs_aux2!cobranza_observaciones = "CUOTA Nro. " + Str(corrprog) + " - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + lbl_titulo
            Select Case parametro
              Case "DNMAN", "DMANS", "DMANB", "DMANC"
                  'rs_aux2!cobranza_observaciones = lbl_titulo + " - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + "Trámite: " + VAR_CITE + "-C-" + Str(corrprog)
                  rs_aux2!cobranza_observaciones = "SERVICIO DE MANTENIMIENTO INTEGRAL - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + "Trámite: " + VAR_CITE + "-C-" + Str(corrprog)
                  rs_aux2!cobranza_concepto_plazo = "SERVICIO DE MANTENIMIENTO INTEGRAL - CUOTA Nº " + Str(corrprog)
              Case "DNREP", "DREPS", "DREPB", "DREPC"
                  rs_aux2!cobranza_observaciones = "SERVICIO DE REPARACIONES, SEGÚN " + Txt_campo2.Text + " - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + "Cite: " + VAR_CITE + "-C-" + Str(corrprog)
                  rs_aux2!cobranza_concepto_plazo = "SERVICIO DE REPARACION, SEGÚN " + VAR_CITE
              Case "DNINS", "DINSS", "DINSB", "DINSC"
                  rs_aux2!cobranza_observaciones = "SERVICIO DE INSTALACION - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + "Cite: " + VAR_CITE + "-C-" + Str(corrprog)
                  rs_aux2!cobranza_concepto_plazo = "SERVICIO DE INSTALACION, SEGÚN " + VAR_CITE
              Case Else
                  rs_aux2!cobranza_observaciones = lbl_titulo + " - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + "Cite: " + VAR_CITE + "-C-" + Str(corrprog)
                  rs_aux2!cobranza_concepto_plazo = "CONFORMIDAD DEL SERVICIO"
            End Select
            
            CONT2 = CONT2 + 1
            rs_aux2!cobranza_requisito_plazo = "S"
            'rs_aux2!cobranza_concepto_plazo = "CONFORMIDAD DEL SERVICIO"
            If rs_aux2!cobranza_programada_bs <> 0 Then
                rs_aux2!Literal = Literal(CStr(rs_aux2!cobranza_programada_bs)) + " BOLIVIANOS"
            End If
            rs_aux2!proceso_codigo = "TEC"
            rs_aux2!subproceso_codigo = "TEC-02"
            rs_aux2!etapa_codigo = "TEC-02-02"
            rs_aux2!clasif_codigo = "TEC"
            rs_aux2!doc_codigo = "R-105"    ' R-307 Certificado de Mantenimiento ' Colocar en la conformidad
            rs_aux2!doc_numero = "0"        'NumComp
            rs_aux2!poa_codigo = "3.2.3"
            rs_aux2!estado_codigo = "REG"
            rs_aux2!usr_codigo = glusuario
            rs_aux2!fecha_registro = Format(Date, "dd/mm/yyyy")
            rs_aux2!hora_registro = Format(Time, "hh:mm:ss")
            rs_aux2!correl_ac = 0
            rs_aux2!estado_ac = "REG"
            rs_aux2.Update
            'Asigna IdCrono (fmes_plan)
            'VAR_ZONA = Ado_datos.Recordset!zpiloto_codigo
            Set rs_aux18 = New ADODB.Recordset
            If rs_aux18.State = 1 Then rs_aux18.Close
            rs_aux18.Open "Select fmes_plan from to_cronograma_mensual where zpiloto_codigo = " & VAR_ZONA & " AND ges_gestion = '" & rs_aux2!gestion & "' AND fmes_correl = " & rs_aux2!cobranza_mes & "  ", db, adOpenKeyset, adLockOptimistic
            If rs_aux18.RecordCount > 0 Then
                db.Execute "update ao_ventas_cobranza_prog set fmes_plan = " & rs_aux18!fmes_plan & " where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
            Else
                db.Execute "update ao_ventas_cobranza_prog set fmes_plan = '0' where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
            End If
            '
          Else
            db.Execute "UPDATE ao_ventas_cobranza_prog SET gestion = '" & Year(rs_aux2!cobranza_fecha_prog) & "', cobranza_mes = '" & Month(rs_aux2!cobranza_fecha_prog) & "' where  venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & ""
            db.Execute "UPDATE ao_ventas_cobranza_prog SET cobranza_prog_codigo = " & correldet2 & " where venta_codigo = " & NumComp & " and YEAR(cobranza_fecha_prog) ='" & Year(FControl) & "'  AND MONTH(cobranza_fecha_prog) ='" & Month(FControl) & "'  "
            
            'Asigna IdCrono (fmes_plan) WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
            Set rs_aux18 = New ADODB.Recordset
            If rs_aux18.State = 1 Then rs_aux18.Close
            rs_aux18.Open "Select fmes_plan from to_cronograma_mensual where zpiloto_codigo = " & VAR_ZONA & " AND ges_gestion = '" & rs_aux2!gestion & "' AND fmes_correl = " & rs_aux2!cobranza_mes & "  ", db, adOpenKeyset, adLockOptimistic
            If rs_aux18.RecordCount > 0 Then
                db.Execute "update ao_ventas_cobranza_prog set fmes_plan = " & rs_aux18!fmes_plan & " where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
            Else
                db.Execute "update ao_ventas_cobranza_prog set fmes_plan = '0' where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
            End If
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
            fdia = Day(FControl)
            fanio = Year(FControl)
            'CONT3 = CONT2 * CONT_MED
            CONT3 = 1
            While (CONT3 <= CONT_MED)
                fmes = Month(FControl)
                Select Case fmes
                    Case 2
                        If fanio = "2012" Or fanio = "2016" Or fanio = "2020" Or fanio = "2024" Then
                            Dias_Mes = 29
                        Else
                            Dias_Mes = 28
                            'Dias_Del_Mes = IIf(saltarYear(Fecha), 29, 28)
                        End If
                    Case 1, 3, 5, 7, 8, 10, 12
                        Dias_Mes = 31
                    Case 4, 6, 9, 11
                        Dias_Mes = 30
                End Select
                If Val(VAR_MES2) = Month(FControl) Then
                    rs_aux2!cobranza_fecha_prog = FControl
                    'rs_aux2!cobranza_fecha_conformidad = FControl + 10
                    rs_aux2!cobranza_fecha_cobro = FControl + 20
                    VAR_MES2 = VAR_MES2 + CONT_MED
                    If Val(VAR_MES2) > 12 Then
                        VAR_MES2 = Val(VAR_MES2) - 12
                    End If
                End If
                FControl = FControl + Dias_Mes
                CONT3 = CONT3 + 1
                CONT4 = CONT4 + Dias_Mes
            Wend
            VAR_FEC2 = MonthName(Month(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog)))
            CONT2 = CONT2 + 1
          End If
        End If
      Wend
      MsgBox "El Cronograma fue Generado Exitosamente... ", , "Atención"
      db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo_verif = 'APR' Where ao_ventas_cabecera.venta_codigo = " & NumComp & " "
      If corrprog > 0 Then
        db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '" & corrprog & "' "
        db.Execute "update ao_ventas_cabecera set venta_plazo_dias_calendario = " & CONT4 & " "
      End If

    Else
       MsgBox "Error Verifique la Venta de Productos..."
    End If
End Sub

Private Sub BtnGrabar1_Click()
    NumComp = Ado_datos.Recordset!venta_codigo
    'VALIDACION FECHAS
    Set rs_aux19 = New ADODB.Recordset
    If rs_aux19.State = 1 Then rs_aux19.Close
    rs_aux19.Open "Select * from gc_tipo_solicitud where solicitud_num = '90' order by ORDEN ", db, adOpenStatic
    If rs_aux19.RecordCount > 0 Then
        rs_aux19.MoveFirst
        While Not rs_aux19.EOF
            'UPDATE  ao_ventas_alcance SET venta_tiempo_dias = (SELECT DATEDIFF(day, fecha_inicio_alcance, fecha_fin_alcance) FROM ao_ventas_alcance WHERE venta_codigo = '8204' and solicitud_tipo='3')
            ' WHERE venta_codigo = '8204' and solicitud_tipo='3'
            
            db.Execute "UPDATE  ao_ventas_alcance SET venta_tiempo_dias = (SELECT DATEDIFF(day, fecha_inicio_alcance, fecha_fin_alcance) FROM ao_ventas_alcance WHERE venta_codigo =  " & NumComp & " and solicitud_tipo= " & rs_aux19!solicitud_tipo & ") WHERE venta_codigo = " & NumComp & " and solicitud_tipo = " & rs_aux19!solicitud_tipo & " "
            
            rs_aux19.MoveNext
        Wend
    End If
    'rs_aux19.Requery
    'Call ABRIR_TABLA_DET
    Call ABRIR_DETALLE
    'Txt_descripcion = DateDiff("y", DTPfechaIni, DTPfechaFin)
    'If Val(Txt_descripcion) < 0 Then
    '    MsgBox "La Fecha de Inicio NO puede ser MAYOR a la Fecha de Finalización, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
    '    DTPfechaFin.SetFocus
    'End If
    DtgAlcance.AllowUpdate = False
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = True
    FraNavega.Enabled = True
    FrmABMDet2.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet1.Visible = True
    FraGrabarCancelar1.Visible = False
End Sub

Private Sub BtnGrabar2_Click()
    Set rs_aux10 = New ADODB.Recordset
    If rs_aux10.State = 1 Then rs_aux10.Close
    rs_aux10.Open "Select * from ao_ventas_cobranza WHERE cobranza_codigo = '" & Ado_datos16.Recordset!cobranza_codigo & "' ", db, adOpenStatic
    If rs_aux10.RecordCount > 0 Then
        rs_aux10.MoveFirst
        db.Execute "UPDATE ao_ventas_cobranza_fac set JustificaAnulacionFac = '" & TxtAnula.Text & "' where IdFactura = '" & rs_aux10!venta_codigo_new & "' "
    Else
        MsgBox "No se puede Solicitar Anulación de Factura, Verifique si se emitió Factura para la cuota ... " & FrmDetalle.Caption, , "Atención"
        'Exit Sub
    End If
    Set rs_aviso_cob = New ADODB.Recordset
    If rs_aviso_cob.State = 1 Then rs_aviso_cob.Close
    rs_aviso_cob.Open "Select * from fc_correl where tipo_tramite = 'aviso_cob'", db, adOpenStatic
    If rs_aviso_cob.RecordCount > 0 Then
        aviso_cob = rs_aviso_cob!numero_correlativo + 1
        db.Execute "update fc_correl set numero_correlativo = " & aviso_cob & " where tipo_tramite = 'aviso_cob'"
        db.Execute "update ao_ventas_cobranza_prog set correl_ac = " & aviso_cob & ", estado_ac = 'APR' where correl_prog = " & Ado_datos16.Recordset!correl_prog & " "
        'Ado_datos16.Recordset!correl_ac = aviso_cob
        'Ado_datos16.Recordset!estado_ac = "APR"
        'Ado_datos16.Recordset.Update
        'cry_ac.ReportFileName = App.Path & "\reportes\ventas\ar_aviso_cobranza.rpt"
        Dim iResult As Variant  ', i%, y%
        cry_ac.WindowShowRefreshBtn = True
        cry_ac.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
        cry_ac.StoredProcParam(1) = Me.Ado_datos16.Recordset!cobranza_prog_codigo
        cry_ac.ReportFileName = App.Path & "\reportes\ventas\ar_solicita_ANL_factura.rpt"
        'cry_ac.ReportFileName = App.Path & "\reportes\ventas\ar_solicita_ANL_factura_PRUEBA.rpt"
        '
        cry_ac.Formulas(1) = "correl = '" & aviso_cob & "' "
        iResult = cry_ac.PrintReport
        If iResult <> 0 Then MsgBox cry_ac.LastErrorNumber & " : " & cry_ac.LastErrorString, vbCritical, "Error de impresión"
    End If
    FraAnula.Visible = False
End Sub

Private Sub BtnGrabarBen_Click()
    db.Execute "UPDATE gc_beneficiario set beneficiario_email = '" & LTrim(TxtEmail.Text) & "', beneficiario_telefono_Cel = '" & TxtCelular.Text & "' where beneficiario_codigo = '" & dtc_benef2A.Text & "' "
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_beneficiario WHERE beneficiario_codigo = '" & dtc_benef2A.Text & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_codigo2A.BoundText = dtc_benef2A.BoundText
    dtc_desc2A.BoundText = dtc_benef2A.BoundText
    dtc_email2A.BoundText = dtc_benef2A.BoundText
    frm_benef.Visible = False
End Sub

Private Sub BtnImprimir_Click()
    fra_reportes.Visible = True
    opt_salir.Value = True
End Sub

Private Sub BtnImprimir1_Click()
   If Ado_datos.Recordset.RecordCount > 0 Then
      If Ado_datos14.Recordset.RecordCount > 0 Then
        Dim iResult As Variant, i%, Y%
        Dim co As New ADODB.Command
        CryV01.ReportFileName = App.Path & "\reportes\Tecnico\tr_orden_servicio_new.rpt"
        'CryV01.WindowShowRefreshBtn = True
        CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
        'CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
     Else
        MsgBox "No se puede Imprimir. Debe registrar datos... " & FrmDetalle.Caption, , "Atención"
     End If
   Else
        MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
   End If

End Sub

Private Sub BtnModDetalle1_Click()
    If Ado_datos6.Recordset.RecordCount > 0 Then
        'nro_licitacion = Ado_datos.Recordset!venta_codigo
        'mw_ventas_alcance.Show vbModal
        NumComp = Ado_datos.Recordset!venta_codigo
        DtgAlcance.Enabled = True
        DtgAlcance.AllowUpdate = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = False
        FraNavega.Enabled = False
        FrmABMDet2.Visible = False
        FrmABMDet.Visible = False
        FrmABMDet1.Visible = False
        FraGrabarCancelar1.Visible = True
    '    NumComp = Ado_datos.Recordset!venta_codigo
    '    frm_ao_ventas_alcance.Show vbModal
    '    Call ABRIR_TABLA_DET
    Else
        MsgBox "No se puede Modificar, debe existir por lo menos un registro, vuelva a intentar...", , "Atención ..."
    End If
End Sub

'Private Sub BtnImprimir3_Click()
'  If (Ado_datos.Recordset.RecordCount > 0) Then
'        Dim iResult As Integer
'        'Dim co As New ADODB.Command
'        CryV01.ReportFileName = App.Path & "\Reportes\tecnico\tr_orden_adenda.rpt"
'        CryV01.WindowShowPrintSetupBtn = True
'        CryV01.WindowShowRefreshBtn = True
'        'MsgBox rs.RecordCount
'        '  cr01.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
'        '  cr01.Formulas(1) = "Subtitulo = '" & FraDet1.Caption & "' "
'
'        CryV01.StoredProcParam(0) = Ado_datos.Recordset!venta_codigo      'ges_gestion
'        'cr01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
'        'cr01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
'        iResult = CryV01.PrintReport
'        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
'        CryV01.WindowState = crptMaximized
'  Else
'    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
'  End If
'End Sub

Private Sub BtnModificar_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  If Ado_datos.Recordset.RecordCount > 0 And dtc_codigo3.Text <> "" Then
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        FrmCabecera.Enabled = True
        FrmDetalle.Visible = False
        FrmCobranza.Visible = False

        FraNavega.Enabled = False
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        Fra_datos.Enabled = True
        Fra_Total.Visible = False
        FrmABMDet.Visible = False
        FrmABMDet2.Visible = False
        BtnImprimir2.Visible = False
        BtnImprimir1.Visible = False
        BtnImprimir4.Visible = False
        
        FrmEdita.Enabled = True
        
        swgrabar = 0
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(2) = False
        If Ado_datos.Recordset!unidad_codigo = "DNMAN" Or Ado_datos.Recordset!unidad_codigo = "DMANS" Or Ado_datos.Recordset!unidad_codigo = "DMANB" Or Ado_datos.Recordset!unidad_codigo = "DMANC" Then
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(3) = False
        Else
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(3) = True
        End If
        dtc_desc2.backColor = &H80000018
        Text11.Visible = False
    Else
        If (Ado_datos.Recordset!estado_codigo = "APR" And Ado_datos.Recordset!estado_cancelado = "P") And (glusuario = "KBETANCOURTH" Or glusuario = "LNAVA" Or glusuario = "FFLORES" Or glusuario = "CARIZACA" Or glusuario = "ADMIN" Or glusuario = "VMEJIA" Or glusuario = "TCASTILLO" Or glusuario = "VBELLIDO" Or glusuario = "FDELGADILLO" Or glusuario = "KGARCIA" Or glusuario = "FCABRERA" Or glusuario = "MARTEAGA" Or glusuario = "RGIL" Or glusuario = "LMORALES" Or glusuario = "GMORA" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Or glusuario = "ARODRIGUEZ") Then
            FrmCabecera.Enabled = True
            FrmDetalle.Visible = False
            FrmCobranza.Visible = False
    
            FraNavega.Enabled = False
            fraOpciones.Visible = False
            FraGrabarCancelar.Visible = True
            Fra_datos.Enabled = True
            Fra_Total.Visible = False
            FrmABMDet.Visible = False
            FrmABMDet2.Visible = False
            BtnImprimir2.Visible = False
            FrmEdita.Enabled = True
            
            swgrabar = 0
            SSTab1.Tab = 0
            SSTab1.TabEnabled(0) = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = True
            
            dtc_desc2.backColor = &H80000018
            Text11.Visible = False
        Else
            MsgBox "NO se puede MODIFICAR, porque el registro ya fue Aprobado, Anulado o Cerrado.", , "Atencion"
        End If
    End If
  Else
        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnSalir_Click()
    sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
    If sino = vbYes Then
'        Ado_datos.Recordset.Close
'        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        If rs_Ventas.State = 1 Then rs_Ventas.Close
        Unload Me
    End If
End Sub

Private Sub BtnVer_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        NumComp = Ado_datos.Recordset!venta_codigo
        Cod_Comp = Ado_datos.Recordset!solicitud_tipo
        tw_ventas_adenda.Show vbModal
    End If
End Sub

Private Sub BtnVer2_Click()
  If Ado_datos.Recordset!estado_codigo = "REG" Or Ado_datos.Recordset!estado_cancelado = "P" Then
    If Ado_datos.Recordset!venta_monto_total_bs = "0" Then
        '
    End If
    NumComp = Ado_datos.Recordset!venta_codigo
    VAR_ZONA = Ado_datos.Recordset!zpiloto_codigo
    'VERIFICA SI EXISTE ITEMS PARA CRONOGRAMA
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select max(cobranza_prog_codigo) as Codigo3 from ao_ventas_cobranza_prog where venta_codigo= " & NumComp & " ", db, adOpenStatic
    If IsNull(rs_aux3!codigo3) Then
        db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & NumComp & " "
        corrprog = 0
        Call CRONO2
    Else
        sino = MsgBox("El Cronograma ya existe, desea volver a Generarlo ? (los items Aprobados no serán modificados)...", vbYesNo + vbQuestion, "Atención ...")
        If sino = vbYes Then
            'OJO BORRAR ao_ventas_cobranza_prog
            db.Execute "DELETE ao_ventas_cobranza_prog where venta_codigo= " & NumComp & " and estado_codigo = 'REG' "
            db.Execute "update ao_ventas_cobranza_prog set venta_codigo_new = cobranza_prog_codigo where venta_codigo= " & NumComp & " "
            db.Execute "update ao_ventas_cobranza_prog set cobranza_prog_codigo = venta_codigo_new + 100 where venta_codigo= " & NumComp & " "
            'db.Execute "update ao_ventas_cobranza set ao_ventas_cobranza.cobranza_prog_codigo = ao_ventas_cobranza_prog.cobranza_prog_codigo FROM ao_ventas_cobranza INNER JOIN ao_ventas_cobranza_prog ON ao_ventas_cobranza.venta_codigo= ao_ventas_cobranza_prog.venta_codigo AND ao_ventas_cobranza.cobranza_prog_codigo= ao_ventas_cobranza_prog.venta_codigo_new "
            db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & NumComp & " "
            db.Execute "update ao_ventas_cabecera set tipo_moneda = 'BOB' where venta_codigo= " & NumComp & " "
            corrprog = 0
            Call CRONO2
        Else
        'If rs_aux3!codigo3 > corrprog Then
            'ACTUALIZAR CORRELATIVO CRONO
            db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '" & rs_aux3!codigo3 & "' where venta_codigo= " & NumComp & " "
            'wwwwwwwwwwwwwwwwwww
            db.Execute "UPDATE ao_ventas_cobranza_prog SET ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' where venta_codigo = " & NumComp & " and ges_gestion <> '" & Ado_datos.Recordset!ges_gestion & "'  "
            db.Execute "update ao_ventas_cabecera set estado_codigo_verif = 'APR' Where venta_codigo = " & NumComp & " "
            'wwwwwwwwwwwwwwwwwww
            corrprog = rs_aux3!codigo3
        'End If
        End If
    End If
  Else
    MsgBox "NO se puede procesar, el trámite ya fue APROBADO o ANULADO ...", , "Atencion"
  End If
End Sub

Private Sub BtnDesAprobar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_cancelado = "S" And Ado_datos.Recordset!estado_codigo = "APR" Then
        MsgBox "NO se puede procesar, el TRAMITE ya fue CERRADO ...", , "Atencion"
    Else
        If Ado_datos.Recordset!estado_cancelado = "N" And Ado_datos.Recordset!estado_codigo = "APR" Then
          sino = MsgBox("Esta seguro marcar como PROVISIONAL, posteriormente debe modificarlo... ", vbYesNo, "Confirmando")
          If sino = vbYes Then
              db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_cancelado = 'P' Where ao_ventas_cabecera.venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  "
              marca1 = Ado_datos.Recordset.Bookmark
              If Ado_datos.Recordset!estado_codigo = "REG" Then
                Call OptFilGral1_Click
              Else
                Call OptFilGral2_Click
              End If
              Ado_datos.Recordset.Move marca1 - 1
          End If
        Else
          If Ado_datos.Recordset!estado_cancelado = "P" And Ado_datos.Recordset!estado_codigo = "APR" Then
            sino = MsgBox("Esta seguro de desmarcar como PROVISIONAL, se convertirá en trámite Aprobado y ya no podrá modificarlo... ", vbYesNo, "Confirmando")
            If sino = vbYes Then
                db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_cancelado = 'N' Where ao_ventas_cabecera.venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  "
                marca1 = Ado_datos.Recordset.Bookmark
                If Ado_datos.Recordset!estado_codigo = "REG" Then
                  Call OptFilGral1_Click
                Else
                  Call OptFilGral2_Click
                End If
                Ado_datos.Recordset.Move marca1 - 1
            End If
          End If
          
        End If
    End If
  Else
    MsgBox "NO se puede procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub Chk_plazo_Click()
    If Chk_plazo.Value = 1 Then
        lbl_plazo.Visible = True
        txt_plazo.Visible = True
    Else
        lbl_plazo.Visible = False
        txt_plazo.Visible = False
    End If
End Sub

'Private Sub Cmd_Cliente_Click()
'    glPersNew = "P"
'    frmBeneficiario.Show 'vbModal
'End Sub

Private Sub CmdCancelaCobro_Click()
  FrmCobros.Enabled = False
  'swgrabar = 0
  'Call cerea
  swnuevo = 0
  If Ado_datos.Recordset("estado_codigo") = "REG" Then
    Call OptFilGral1_Click
  Else
    Call OptFilGral2_Click
  End If
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    FraNavega.Enabled = True
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = True
    'BtnImprimir1.Visible = True
    'BtnImprimir4.Visible = True
    FrmDetalle.Visible = True
    FrmCobranza.Visible = True
    TxtCobrador.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
'        DTPFechaProg.Visible = True
'        DTPFechaConf.Visible = True
'        DTPFechaProg.Enabled = True
    TxtMonto.Enabled = True
    TxtDsctoTot.Enabled = True
    TxtObs.Enabled = True
End Sub

Private Sub BtnAnlDetalle2_Click()
 If Ado_datos.Recordset!estado_codigo = "REG" Then
   sino = MsgBox("Está seguro de ANULAR este registro", vbYesNo + vbQuestion, "Atención ...")
   If sino = vbYes Then
      db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.estado_codigo = 'ANL' Where ao_ventas_cobranza_prog.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza_prog.cobranza_codigo = " & Ado_datos16.Recordset("cobranza_codigo") & " "
      'db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.cobranza_deuda_bs = '0', ao_ventas_cobranza_prog.cobranza_deuda_dol = '0'  Where ao_ventas_cobranza_prog.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cobranza_prog.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza_prog.cobranza_codigo = " & ado_datos16.Recordset("cobranza_codigo") & " "

     'ado_ventas_COBRANZAS.Recordset.Delete
     'ado_ventas_COBRANZAS.Recordset.Update
     'ado_ventas_COBRANZAS.Requery
     'ado_ventas_COBRANZAS.Refresh
     ''cerea
     'ado_ventas_COBRANZAS.Refresh
   End If
  Else
    MsgBox "Los productos del registro sin Aprobar, NO pueden ser ANULADOS !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnModDetalle2_Click()
  If Ado_datos16.Recordset!estado_codigo = "REG" Then       'And (Ado_datos.Recordset!venta_tipo = "E" Or Ado_datos.Recordset!venta_tipo = "V" Or Ado_datos.Recordset!venta_tipo = "C")
    
    marca1 = Ado_datos16.Recordset.Bookmark
    Call ABRIR_TABLAS_AUX
    VAR_COBRANZA = Ado_datos16.Recordset!cobranza_prog_codigo
    FraNavega.Enabled = False
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = False
    BtnImprimir1.Visible = False
    BtnImprimir4.Visible = False
    FrmDetalle.Visible = False
    FrmCobranza.Visible = False
    VAR_COBR1 = Ado_datos16.Recordset!cobranza_prog_codigo
    'swgrabar = 0
    swnuevo = 2
    TxtCobrador.Visible = False
    If glusuario = "VBELLIDO" Or glusuario = "GSOLIZ" Or glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "KBETANCOURTH" Or glusuario = "LNAVA" Or glusuario = "FFLORES" Or glusuario = "CARIZACA" Or glusuario = "RGIL" Or glusuario = "LMORALES" Or glusuario = "GMORA" Or glusuario = "MARTEAGA" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CARIZACA" Or glusuario = "CSALINAS" Or glusuario = "ARODRIGUEZ" Or glusuario = "RLAVAYEN" Then
        TxtMonto.Enabled = True
        TxtMonto.Locked = False
        TxtDsctoTot.Enabled = False
        TxtObs.Enabled = True
    Else
        TxtMonto.Enabled = False
        TxtMonto.Locked = True
        TxtDsctoTot.Enabled = False
        TxtObs.Enabled = False
    End If
    'TxtMonto.SetFocus
    'TxtNroVenta.Enabled = False
    'marca1 = ado_datos14.Recordset.BookMark
    'txt_descripcion_venta.Enabled = True
    'TxtNroVenta.Text = txt_venta.Text
    'lbltipoVenta.Caption = dtc_desc11.Text
    'lblges_gestion.Caption = Ado_datos.Recordset!ges_gestion
    SSTab1.Tab = 2
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    FrmCobros.Visible = True
    FrmCobros.Enabled = True
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
    'If Ado_datos.Recordset!estado_codigo = "APR" Then
        'sino = MsgBox("Registrará la cobranza efectiva, ahora ? ", vbYesNo, "Confirmando")
        'If sino = vbYes Then
        '    DTPFechaProg.Visible = False
        '    DTPFechaCobro.Visible = True
        '    Lbl_nombre_fac.Caption = "Factura a Nombre de:"
        '    lbl_fechas.Caption = "Fecha Efectiva de Cobranza"
        '    Txt_parche.Visible = False      '&H80000013&
        '    'dtc_desc2A.BackColor = &H80000013
        'Else
        '    DTPFechaProg.Visible = True
        '    DTPFechaCobro.Visible = False
        '    Lbl_nombre_fac.Caption = "Cliente :"
        '    lbl_fechas.Caption = "Fecha Programada de Cobranza"
        '    Txt_parche.Visible = True       '&H80000005&
        '    'dtc_desc2A.BackColor = &H80000005
        'End If
    'Else
'        DTPFechaProg.Visible = True
''        DTPFechaCobro.Visible = False
'        DTPFechaConf.Visible = True
''        Lbl_nombre_fac.Caption = "Cliente :"
        lbl_fechas.Caption = "Fecha Programada de Cobranza"
        'FACTURA FISICA
'        Txt_parche.Visible = True       '&H80000005&
        'dtc_desc2A.BackColor = &H80000005
    'End If
    VAR_MBS2 = Ado_datos16.Recordset!cobranza_programada_bs
    'TxtMonto.SetFocus
'    Call ABRIR_TABLA_DET
'    Ado_datos16.Recordset.Move marca1 - 1
'    If Aux = "DNMAN" Then
'        If (glusuario = "VBELLIDO" Or glusuario = "GPALLY" Or glusuario = "ADMIN" Or glusuario = "VPAREDES") Then
'            TxtMonto.Enabled = True
'            TxtDsctoTot.Enabled = True
'        Else
'            TxtMonto.Enabled = False
'            TxtDsctoTot.Enabled = False
'        End If
'    End If

  Else
    MsgBox "El Registro ya fue Aprobado o Anulado !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnAddDetalle2_Click()
  marca1 = Ado_datos16.Recordset.Bookmark
  'If Ado_datos.Recordset!venta_tipo = "C" And Ado_datos.Recordset!estado_codigo = "APR" Then
  If Ado_datos.Recordset!venta_tipo = "C" Or Ado_datos.Recordset!venta_tipo = "V" Then
    If Ado_datos.Recordset!venta_saldo_p_cobrar_bs > 0 Then
    'If Ado_datos.Recordset!venta_monto_total_bs - Ado_datos.Recordset!venta_monto_cobrado_bs > 0 Then
        swnuevo = 1
        SSTab1.Tab = 2
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        FrmCobros.Visible = True
        FrmCobros.Enabled = True
        fraOpciones.Enabled = False
        FraNavega.Enabled = False
        FrmDetalle.Visible = False
        FrmCobranza.Visible = False
        FrmABMDet.Visible = False
        FrmABMDet2.Visible = False
        TxtCobrador.Visible = False
        Ado_datos16.Recordset.AddNew
        dtc_codigo2A.Text = dtc_codigo2.Text
        dtc_desc2A.Text = dtc_desc2.Text
        TxtMonto.SetFocus
        DTPFechaProg.Visible = True
'        DTPFechaCobro.Visible = False
'        Lbl_nombre_fac.Caption = "Cliente :"
        lbl_fechas.Caption = "Fecha Programada de la Cobranza"
        'Txt_parche.Visible = True
        'Ado_datos.Recordset.Move marca1 - 1
'        Dim thisDate As Date
'        Dim thisMonth As Integer
'        thisDate = #2/12/1969#
'        thisMonth = Month(thisDate)
'        ' thisMonth now contains 2.
'
'
'        Dim thisMonth As Integer
'        Dim name As String
'        thisMonth = 4
'        ' Set Abbreviate to True to return an abbreviated name.
'        name = MonthName(thisMonth, True)
'        ' name now contains "Apr".
    Else
        MsgBox "Ya se cobró el total de la deuda, Verifique por favor !! ", vbExclamation, "Atención!"
    End If
  Else
    MsgBox "La Venta (al Contado o Donación) NO tiene saldo para cobrar, Verifique por favor !! ", vbExclamation, "Atención!"
  End If
End Sub

'Private Sub BtnDesAprobar_Click()
''  sino = MsgBox("Esta seguro de Desaprobar el registro?", vbYesNo, "Confirmando")
''  If sino = vbYes Then
''    Dim rstdestino As New ADODB.Recordset
''    Set rstdestino = New ADODB.Recordset
''    If rstdestino.State = 1 Then rstdestino.Close
''    rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correl_venta = " & Ado_datos.Recordset("correl_venta") & " and venta_codigo = " & Ado_datos.Recordset("venta_codigo") & " ", db, adOpenDynamic, adLockOptimistic
''    If Not rstdestino.BOF Then rstdestino.MoveFirst
''    If Not rstdestino.BOF And Not rstdestino.EOF Then
''      rstdestino("estado_codigo") = "REG"
''      rstdestino.Update
''    End If
''    If rstdestino.State = 1 Then rstdestino.Close
''    marca1 = Ado_datos.Recordset.Bookmark
''    Call OptFilGral1_Click
''    Ado_datos.Recordset.Move marca1 - 1
''  End If
'End Sub

'Private Sub CmdDetallePoa_Click()
'  If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
'   MsgBox "No Existen Registros ", vbInformation, "Formulario 11"
'  Else
'    marca1 = Ado_datos.Recordset.BookMark
'    FrmPoasCapturaALB.Lblformulario = "F11"
'    FrmPoasCapturaALB.lblges_gestion = Ado_datos.Recordset!ges_gestion
'    FrmPoasCapturaALB.lblcodigo_unidad = Ado_datos.Recordset!codigo_unidad
'    FrmPoasCapturaALB.lblcodigo_solicitud = Ado_datos.Recordset!codigo_solicitud
'    FrmPoasCapturaALB.lbltipo_beneficiario = "N" 'Ado_datos.Recordset!tipoben_codigo
'    FrmPoasCapturaALB.Show vbModal
'  If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
'    '
'  Else
'    Ado_datos.Refresh
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
'  End If
'End Sub

'Private Sub cmdElige_Click()
'  With ALFrmMateriales
'        .ALPrincipal
'        If .QResp Then
'            TxtCodigo.Text = .QCodigo
'            txtDesc.Text = .QItem
'        End If
'    End With
'    Txtcant_alm = 0
'    Cant_Alm = 0
'    DE.dbo_albSacaDetalleMaterial Mid(TxtCodigo, 3, 12), descri_bien, Cant_Alm
'    Txtcant_alm = Cant_Alm
'    If Cant_Alm >= TxtCantPedi Then
'        optSi = True
'    Else
'        optNo = True
'    End If
'End Sub

Private Sub graba_proyecto()
'    Select Case Ado_datos.Recordset!unidad_codigo
'       Case "DNAJS", "DNEME", "DNINS", "DNMAN", "DNMOD", "DNREP", "DINSB", "DINSC", "DINSS", "DAJSB", "DAJSC", "DAJSS", "DMANB", "DMANC", "DMANS", "DREPB", "DREPC", "DREPS", "DEMEB", "DEMEC", "DEMES", "DMODB", "DMODC", "DMODS"
'            VAR_PROY = 12
'        Case "UCOM"
'            VAR_PROY = 17
'        Case "DVTA"
'            VAR_PROY = 18
'
'    End Select
'
'    Set rs_aux1 = New ADODB.Recordset
'    If rs_aux1.State = 1 Then rs_aux1.Close
'    SQL_FOR = "select * from fo_proyectos_ejecucion where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
'    rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'    If rs_aux1.RecordCount > 0 Then
'        db.Execute "update fo_proyectos_ejecucion set pro_codigo_det_descripcion = '" & dtc_desc3.Text & "' Where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
'    Else
'        db.Execute "INSERT INTO fo_proyectos_ejecucion (pro_codigo, pro_codigo_det, pro_codigo_det_descripcion, unidad_codigo, ges_gestion, estado_codigo, usr_codigo, fecha_registro) " & _
'           "VALUES (" & VAR_PROY & ", '" & Ado_datos.Recordset!edif_codigo & "', '" & dtc_desc3.Text & "', '" & Ado_datos.Recordset!unidad_codigo & "', " & glGestion & ", 'APR', '" & glusuario & "', '" & Date & "')"
'    End If
'    '
End Sub

Private Sub graba_ingreso()
'    '======= Ini grabado de datos
'   'swgraba = 0
'   'Call valida
'   VAR_COD4 = Ado_datos.Recordset!unidad_codigo
'   VAR_CODTIPO = "DEI"
'   Select Case VAR_COD4
'        Case "DVTA"              'INI COMERCIAL
'            VAR_ORG = "111"
'            VAR_PARTIDA = "11310"
'        Case "COMEX"            'INI COMEX
'            VAR_ORG = "111"
'            VAR_PARTIDA = "11310"
'        Case "DNINS", "DINSB", "DINSC", "DINSS"            'INI INSTALACIONES
'            VAR_ORG = "111"
'            VAR_PARTIDA = "11350"
'        Case "DNAJS", "DAJSB", "DAJSC", "DAJSS"           'INI AJUSTE
'            VAR_ORG = "113"
'            VAR_PARTIDA = "11350"
'        Case "DNMAN", "DMANB", "DMANC", "DMANS"            'INI MANTENIMIENTO
'            VAR_ORG = "112"
'            VAR_PARTIDA = "11320"
'        Case "DNREP", "DREPB", "DREPC", "DREPS"            'INI REPARACIONES
'            VAR_ORG = "113"
'            VAR_PARTIDA = "11330"
'        Case "DNMOD", "DMODB", "DMODC", "DMODS"            'INI MODERNIZACION
'            VAR_ORG = "114"
'            VAR_PARTIDA = "11340"
'        Case "DNEME", "DEMEB", "DEMEC", "DEMES"            'INI EMERGENCIAS
'            VAR_ORG = "113"
'            VAR_PARTIDA = "11330"
'        Case Else               'INI CREDITO
'            VAR_ORG = "311"
'            VAR_PARTIDA = "11350"
'   End Select
''   If swgraba = 1 Then
''      FraOpciones2.Visible = False
''      fraOpciones.Visible = True
''      FraIngresosNav.Enabled = True
''      FraIngresosDat.Enabled = False
'
'      'If v_añadir = 1 Then
'        'EFECTIVO o a CREDITO
'         'db.BeginTrans
'         Call add_correl
'         Set rstdestino = New ADODB.Recordset
'         rstdestino.Open "select * from fo_ingresos_cabecera order by org_codigo, ingreso_codigo   ", db, adOpenDynamic, adLockOptimistic
'         rstdestino.AddNew
'         rstdestino("Ges_Gestion") = glGestion      'Year(Date)     'Ado_datos.Recordset("ges_gestion")
'         rstdestino("ingreso_codigo") = correlativo1
'         VAR_CODANT = correlativo1
'         'CAMBIAR org_codigo
'         rstdestino("org_codigo") = VAR_ORG
'         'CAMBIAR org_codigo
'         'CAMBIAR COD ingreso_codigo_anterior
'         rstdestino("ingreso_codigo_anterior") = correlativo1
'         'CAMBIAR COD ingreso_codigo_anterior
'         'CAMBIAR DEI O REC
'         'VAR_CODTIPO = "DEI"
'         rstdestino("Codigo_tipo") = VAR_CODTIPO    '"DEI"
'         'VAR_CODTIPO = "DEI"
'         'CAMBIAR DEI O REC
'         rstdestino("proceso_codigo") = "FIN"
'         rstdestino("subproceso_codigo") = "FIN-01"
'         rstdestino("etapa_codigo") = "FIN-01-01"
'         rstdestino("clasif_codigo") = "ADM"
'         rstdestino("doc_codigo") = "R-110"
'         rstdestino("doc_numero") = correlativo1
'         rstdestino("unidad_codigo") = VAR_COD4     'Ado_datos.Recordset("unidad_codigo")
'         rstdestino("solicitud_codigo") = VAR_SOL   'Ado_datos.Recordset("solicitud_codigo")
'         rstdestino("solicitud_tipo") = VAR_TIPO    '"10"
'
'         rstdestino("beneficiario_codigo") = VAR_BENEF      'Ado_datos.Recordset("beneficiario_codigo")
'         'VAR_BENEF = Ado_datos.Recordset("beneficiario_codigo")
'         rstdestino("fecha_ingreso") = Date
'         rstdestino("tipo_cambio") = GlTipoCambioOficial 'GlTipoCambioMercado
'         rstdestino("tipo_moneda") = "BOB"
'         VAR_MONEDA = "BOB"
'         rstdestino("ingreso_concepto") = "INGRESO POR: " + VAR_GLOSA2  'Ado_datos.Recordset("venta_descripcion")
'         VAR_GLOSA = "INGRESO POR: " + VAR_GLOSA2       'Ado_datos.Recordset("venta_descripcion")
'         If Ado_datos.Recordset("venta_tipo") = "E" Then
'            rstdestino("tipo_comp") = "DYR"
'         Else
'            rstdestino("tipo_comp") = "DEI"
'         End If
'         'CAMBIAR FTE
'         Select Case VAR_ORG
'             Case "111"              'INI SERVICIOS DE PROVISION E INSTALACION
'                 VAR_FTE = "10"
'             Case "112"            'INI SERVICIO DE MANTENIMIENTO - MANTENIMIENTO PREVENTIVO
'                 VAR_FTE = "10"
'             Case "113"            'INI SERVICIO DE REPARACIONES - MANTENIMIENTO CORRECTIVO
'                 VAR_FTE = "10"
'             Case "114"            'INI SERVICIO DE MODERNIZACION
'                 VAR_FTE = "10"
'             Case "211"            'INI APORTES DE CAPITAL
'                 VAR_FTE = "20"
'             Case "311"            'INI BANCO MERCANTIL SANTA CRUZ
'                 VAR_FTE = "30"
'             Case "312"            'INI BANCO DE CREDITO
'                 VAR_FTE = "30"
'             Case "411"            'INI AMT - REPOSICION DE PIEZAS Y PARTES
'                 VAR_FTE = "40"
'             Case Else               'INI OTROS
'                 VAR_FTE = "10"
'        End Select
'         rstdestino("fte_codigo") = VAR_FTE
'         'CAMBIAR FTE
'         'CAMBIAR RUBROS    'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww ya pues
'         'rstdestino("rubro_codigo") = "11200"
'         'VAR_PARTIDA = "11200"
'         'VAR_PARTIDA = "11320"
'         rstdestino("rubro_codigo") = VAR_PARTIDA
'         'CAMBIAR RUBROS
'         rstdestino("cheque_o_trf") = ""
'         rstdestino("Bco_codigo") = "NN"
'         'CAMBIAR CTA
'         rstdestino("cta_codigo") = "NN"
'         VAR_CTA = "NN"
'         'CAMBIAR CTA
'         rstdestino("numero_documento") = "0"
'         rstdestino("unidad_codigo_ant") = VAR_CITE     'Ado_datos.Recordset("unidad_codigo_ant")
'         'VAR_CITE = Ado_datos.Recordset("unidad_codigo_ant")
'         rstdestino("monto_dolares") = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
'         VAR_DOL2 = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
'         rstdestino("monto_bolivianos") = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
'         VAR_BS2 = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
'         rstdestino("monto_recaudado_dolares") = 0
'         rstdestino("monto_recaudado_bolivianos") = 0
'         rstdestino("convenio_codigo") = "NN"
'         rstdestino("pro_codigo_det") = Ado_datos.Recordset("edif_codigo")
'         VAR_PROY2 = Ado_datos.Recordset("edif_codigo")
'         rstdestino("estado_CODIGO") = "APR"
'         'rstdestino("estado_codigo_dr") = "DEI"
'
'         rstdestino("usr_CODIGO") = glusuario
'         rstdestino("fecha_registro") = Date
'         rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
'
'         rstdestino.Update
'         If rstdestino.State = 1 Then rstdestino.Close
'        'db.CommitTrans
'
''          If rstIngresos.State = 1 Then rstIngresos.Close
''          rstIngresos.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
''          rstIngresos.Sort = "ingreso_codigo"
''          rstIngresos.Requery
'
''          rstIngresos.Requery
''          Set AdoIngresos.Recordset = rstIngresos
''          AdoIngresos.Refresh
''          AdoIngresos.Recordset.Find "ultimo = 'S'"
''          If Not (AdoIngresos.Recordset.EOF) Then
''            marca1 = AdoIngresos.Recordset.Bookmark
''            AdoIngresos.Recordset("ultimo") = "N"
''            AdoIngresos.Recordset.Update
''          End If
'
''          AdoIngresos.Recordset.Move marca1 - 1
'
''          marca1 = 0
'      'End If
''   Else
''      MsgBox "ERROR Los datos no están completos, no se realizará la grabación..."
'''      FraOpciones2.Visible = False
'''      FraOpciones.Visible = True
'''      FraIngresosNav.Enabled = True
'''      FraIngresosDat.Enabled = False
'''      AdoIngresos.Refresh
''   End If
''   LblAccion = ""
''AAQQQQQUIIIIIIIIII    JQA

End Sub

Private Sub add_correl()
'  'FALTAAAAA!! org_codigo JQA 2014-07-10
'  Set rstcorrel_ing = New ADODB.Recordset
'  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
'  rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '" & VAR_ORG & "' ", db, adOpenDynamic, adLockOptimistic
'  'rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '111' and ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "'", db, adOpenDynamic, adLockOptimistic
'  If rstcorrel_ing.RecordCount = 0 Then
'     rstcorrel_ing.AddNew
'     rstcorrel_ing("org_codigo") = VAR_ORG
'     rstcorrel_ing("ges_gestion") = glGestion       'Ado_datos.Recordset("ges_gestion")  'Trim(lblges_gestion.Caption)
'     'rstcorrel_ing("correlativo") = 1
'     rstcorrel_ing("correlativo_ingreso") = 1
'     rstcorrel_ing.Update
'     correlativo1 = rstcorrel_ing("correlativo_ingreso")
'     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo_ingreso")
'  Else
'     VARG_ORGD = rstcorrel_ing!org_descripcion
'     rstcorrel_ing("correlativo_ingreso") = rstcorrel_ing("correlativo_ingreso") + 1
'     rstcorrel_ing.Update
'     correlativo1 = rstcorrel_ing("correlativo_ingreso")
'     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo")
'  End If
'  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close

End Sub

Private Sub CmdNOunidad_Click()
    swunidad = 0
    Frmunidad.Visible = False
End Sub

Private Sub CmdOKunidad_Click()
    swunidad = 1
        If swunidad = 1 Then
            Dim rstpagos As New ADODB.Recordset
            Set rstpagos = New ADODB.Recordset
            If rstpagos.State = 1 Then rstpagos.Close
            rstpagos.Open "select * from pagos where GES_gestion = '5000'", db, adOpenKeyset, adLockOptimistic
            rstpagos.AddNew
                rstpagos("ges_gestion") = glGestion     'Ado_datos.Recordset("ges_gestion")
                rstpagos("org_codigo") = DataCombo1.Text   'Ado_datos.Recordset("formulario")
                rstpagos("codigo_pago") = "" 'genera jorge
                rstpagos("codigo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
                rstpagos("formulario") = Ado_datos.Recordset("formulario")
                rstpagos("codigo_unidad") = Ado_datos.Recordset("codigo_unidad")
                rstpagos("monto_bolivianos") = Ado_datos.Recordset("monto_bolivianos")
                rstpagos("estado_compromiso") = "N"
                rstpagos("justificacion") = Ado_datos.Recordset("justificacion_solicitud")
            rstpagos.Update
        End If
End Sub


Private Sub CmdEmail_Click()
    Set rs_aux12 = New ADODB.Recordset
    If rs_aux12.State = 1 Then rs_aux12.Close
    rs_aux12.Open "Select * from gc_beneficiario where (beneficiario_codigo = '" & dtc_benef2A.Text & "') ", db, adOpenStatic   '
    If rs_aux12.RecordCount > 0 Then
        TxtEmail.Text = IIf(IsNull(rs_aux12!beneficiario_email), "-", rs_aux12!beneficiario_email)
        TxtCelular.Text = IIf(IsNull(rs_aux12!beneficiario_telefono_Cel), "0", rs_aux12!beneficiario_telefono_Cel)
    'Else
    End If
    frm_benef.Visible = True
End Sub

Private Sub CmdGrabaCobro_Click()
  NumComp = Ado_datos.Recordset!venta_codigo
  VAR_COBRANZA = Ado_datos16.Recordset!cobranza_prog_codigo
    If TxtMonto.Text >= 1000 And dtc_codigo2A.Text = "0" Then
        MsgBox "No se puede Solicitar una Factura >= Bs.1000, sin NIT, debe registrar el NIT del Beneficiario... ", , "Atención"
        Exit Sub
    End If
    'Dim MyVar, MyCheck
    'MyVar = "53"    ' Assign value.
    'MyCheck = IsNumeric(MyVar)    ' Returns True.
    Dim MyPos, MyPos2 As Integer
    Dim ARROBA, punto As String
    If IsNumeric(dtc_codigo2A.Text) Then
    Else
        MsgBox "El NIT del Cliente a Facturar es Incorrecto, debe corregir y luego vuelva a Intentar ...", vbExclamation, "Atención"
        Exit Sub
    End If
    If (glusuario <> "VBELLIDO") Then
        MyPos = 0
        MyPos2 = 0
        ARROBA = "@"
        punto = "."
        MyPos = InStr(1, dtc_email2A.Text, ARROBA, 1)
        MyPos2 = InStr(1, dtc_email2A.Text, punto, 1)
        'MyPos = Instr(4, SearchString, SearchChar, 1)
        If (Len(dtc_email2A.Text) > 10) Then
            If (MyPos > 0) And (MyPos2 > 0) Then
            Else
                MsgBox "Debe Registrar el EMail del Cliente a Facturar, !! Vuelva a Intentar ...", vbExclamation, "Atención"
                Exit Sub
            End If
        Else
            MsgBox "Debe Registrar el EMail del Cliente a Facturar, !! Vuelva a Intentar ...", vbExclamation, "Atención"
            Exit Sub
        End If
    End If
  If dtc_codigo4A = "" Then
    MsgBox "Debe Elejir " + Lbl_Cobrador.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
'  If TxtMonto = "" Or TxtMonto = "0" Or TxtMonto = "0.00" Then
'    MsgBox "Debe Registrar el " + lbl_monto.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
'    Exit Sub
'  End If
  If TxtObs = "" Then
    MsgBox "Debe Registrar el " + lbl_obs.Caption + " de la Cobranza, !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If txtDoc = "" Then
    MsgBox "Debe Registrar el " + lblccertif.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
'  If DTPFechaConf = "" Then
'    MsgBox "Debe Registrar la " + lblfechaCertif.Caption + " de la Cobranza, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'    Exit Sub
'  End If
  Select Case CmbEmision.Text
    Case "FACTURA FISICA"
        VAR_EMISION = "28"
    Case "CORREO ELECTRONICO"
        VAR_EMISION = "29"
    Case "WHATSAPP"
        VAR_EMISION = "30"
    Case Else
        VAR_EMISION = "28"
  End Select
  
  'If swnuevo = 2 Then
  'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
'  If DTPFechaProg.Visible = False Then
'    If TxtCmpbte = "" Or TxtCmpbte = "0" Then
'       MsgBox "Debe Registrar el " + lbl_factura.Caption + " a emitir al Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'      Exit Sub
'    End If
'  End If
  'fin PARA COBRANZA WWWWWWWWWWWWWWWWWWW
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "select sum(cobranza_programada_bs) as totbs2, sum (cobranza_programada_dol) as totdl2 from ao_ventas_cobranza_prog where venta_codigo=" & Ado_datos.Recordset!venta_codigo & "  ", db, adOpenKeyset, adLockOptimistic
    If IsNull(rs_aux3!totbs2) Then
        If CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
            'MsgBox "No puede programar un <" + lbl_monto.Caption + "> que sobrepase el <" + lbl_totalBs.Caption + "> . !! Vuelva a Intentar ...", vbExclamation, "Atención"
            MsgBox "No puede programar un <" + lbl_monto.Caption + "> que sobrepase el Monto total del Contrato . !! Vuelva a Intentar ...", vbExclamation, "Atención"
            If rs_aux3.State = 1 Then rs_aux3.Close
            Exit Sub
        End If
    Else
        If swnuevo = 1 Then
            If (rs_aux3!totbs2) + CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
                'MsgBox "No puede programar un <" + lbl_monto.Caption + "> que sobrepase el <" + lbl_totalBs.Caption + "> . !! Vuelva a Intentar ...", vbExclamation, "Atención"
                MsgBox "No puede programar un <" + lbl_monto.Caption + "> que sobrepase el Monto total del Contrato . !! Vuelva a Intentar ...", vbExclamation, "Atención"
                If rs_aux3.State = 1 Then rs_aux3.Close
                Exit Sub
            End If
        Else
'            If (rs_aux3!totbs2) - VAR_MBS2 + CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
'                MsgBox "No puede programar un <" + lbl_monto.Caption + "> que sobrepase el <" + lbl_totalBs.Caption + "> . !! Vuelva a Intentar ...", vbExclamation, "Atención"
'                If rs_aux3.State = 1 Then rs_aux3.Close
'                Exit Sub
'            End If
        End If
    End If
  'valida = 1
  'If valida = 1 And dtc_codigo4A <> "" Then
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
    db.BeginTrans
    If swnuevo = 1 Then
      Set rs_aux1 = New ADODB.Recordset
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from ao_ventas_cabecera where venta_codigo=" & Ado_datos.Recordset!venta_codigo & "  ", db, adOpenKeyset, adLockOptimistic
      If rs_aux1.RecordCount > 0 Then
         correldet2 = rs_aux1!correl_cobro_prog + 1
         If rs_aux1!correl_cobro_prog > 1 Then
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            rs_aux2.Open "Select * from ao_ventas_cobranza_prog where venta_codigo=" & Ado_datos.Recordset!venta_codigo & " and cobranza_prog_codigo = " & rs_aux1!correl_cobro_prog & " ", db, adOpenStatic
            If rs_aux2.RecordCount > 0 Then
                If DTPFechaProg.Value <= rs_aux2!cobranza_fecha_prog Then
                    MsgBox "No puede registrar una " + lbl_fechas.Caption + " menor o igual a la anterior. !! Vuelva a Intentar ...", vbExclamation, "Atención"
                    If rs_aux1.State = 1 Then rs_aux1.Close
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    db.CommitTrans
                    Exit Sub
                End If
            End If

         End If
         rs_aux1!correl_cobro_prog = rs_aux1!correl_cobro_prog + 1
         rs_aux1.Update
      End If
      'Ado_datos16.Recordset.AddNew
      Ado_datos16.Recordset!cobranza_prog_codigo = correldet2
      Ado_datos16.Recordset!venta_codigo = Ado_datos.Recordset("venta_codigo")
      Ado_datos16.Recordset!ges_gestion = glGestion      'Ado_datos.Recordset("ges_gestion")
    End If
    If swnuevo = 2 Then
      If Ado_datos16.Recordset!cobranza_prog_codigo > 1 Then
        correldet2 = Ado_datos16.Recordset!cobranza_prog_codigo - 1
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "Select * from ao_ventas_cobranza_prog where venta_codigo=" & Ado_datos.Recordset!venta_codigo & " and cobranza_prog_codigo = " & correldet2 & " ", db, adOpenStatic
        If rs_aux2.RecordCount > 0 Then
          If DTPFechaProg.Value <= rs_aux2!cobranza_fecha_prog Then 'DTPFechaProg.Value
          'If DTPFechaProg.Value <= rs_aux2!cobranza_fecha_prog Then
              MsgBox "No puede registrar una " + lbl_fechas.Caption + " menor o igual a la anterior. !! Vuelva a Intentar ...", vbExclamation, "Atención"
              If rs_aux2.State = 1 Then rs_aux2.Close
              db.CommitTrans
              Exit Sub
          End If
        End If
      End If
    End If
      Ado_datos16.Recordset!beneficiario_codigo = dtc_benef2A.Text                                 'Codigo Beneficiario/Cliente
      Ado_datos16.Recordset!beneficiario_codigo_resp = dtc_codigo4A.Text                                                     'Codigo Cobrador
      'Ado_datos16.Recordset!nombre_cobrador = dtc_desc4A.Text   '+ " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
      Ado_datos16.Recordset!cobranza_programada_bs = CDbl(TxtMonto.Text)                                  'Monto Programado Bs
      Ado_datos16.Recordset!cobranza_programada_dol = CDbl(TxtMonto.Text) / GlTipoCambioMercado        'Monto Programado en Dolares
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
      Ado_datos16.Recordset!cobranza_descuento_bs = 0                                 'Descuento Bs
      Ado_datos16.Recordset!cobranza_descuento_dol = 0                                    'Descuento Dol
      Ado_datos16.Recordset!cobranza_total_bs = CDbl(TxtMonto.Text)   'Ado_datos16.Recordset!cobranza_deuda_bs - Ado_datos16.Recordset!cobranza_descuento_bs               'Monto Total Bs
      Ado_datos16.Recordset!cobranza_total_dol = CDbl(TxtMonto.Text) / GlTipoCambioMercado  'Ado_datos16.Recordset!cobranza_deuda_dol - Ado_datos16.Recordset!cobranza_descuento_dol               'Monto Total Dol
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
      If Ado_datos16.Recordset!cobranza_programada_bs <> 0 Then
            Ado_datos16.Recordset!Literal = Literal(CStr(Ado_datos16.Recordset!cobranza_programada_bs)) + " BOLIVIANOS"
            'Ado_datos16.Recordset!Literal = Literal(CStr(Ado_datos.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
      End If
      Ado_datos16.Recordset!cobranza_fecha_cobro = DTPFechaCobro.Value                                'Fecha de Cobranza cobranza_fecha_conformidad
      'Ado_datos16.Recordset!cobranza_fecha_conformidad = DTPFechaConf.Value                                'Fecha de Cobranza
'      Call acumulaMont(Ado_datos16.Recordset("ges_gestion"), Ado_datos16.Recordset("venta_codigo"))

      Ado_datos16.Recordset!cobranza_requisito_plazo = "S"
      Ado_datos16.Recordset!cobranza_concepto_plazo = txt_plazo.Text
      
'      If Chk_plazo.Value = 1 Then
'        lbl_plazo.Visible = True
'        txt_plazo.Visible = True
'        Ado_datos16.Recordset!cobranza_requisito_plazo = "S"
'        Ado_datos16.Recordset!cobranza_concepto_plazo = "CERTIFICADO DE MANTENIMIENTO R-307 Nro. " + txtDoc
'      Else
'        lbl_plazo.Visible = False
'        txt_plazo.Visible = False
'        Ado_datos16.Recordset!cobranza_requisito_plazo = "N"
'        Ado_datos16.Recordset!cobranza_concepto_plazo = txt_plazo.Text
'      End If
      Ado_datos16.Recordset!nro_fojas = IIf(txt_fojas.Text = "", "1", txt_fojas.Text)
      Ado_datos16.Recordset!cobranza_observaciones = TxtObs.Text
      Ado_datos16.Recordset!proceso_codigo = "TEC"
      Ado_datos16.Recordset!subproceso_codigo = "TEC-02"
      Ado_datos16.Recordset!etapa_codigo = "TEC-02-02"
      Ado_datos16.Recordset!clasif_codigo = "TEC"
      Ado_datos16.Recordset!doc_codigo = "R-110"        '"R-307"
      If Ado_datos.Recordset!unidad_codigo = "DNREP" Or Ado_datos.Recordset!unidad_codigo = "DREPS" Or Ado_datos.Recordset!unidad_codigo = "DREPB" Or Ado_datos.Recordset!unidad_codigo = "DREPC" Then
        Ado_datos16.Recordset!doc_numero = IIf(Ado_datos.Recordset!doc_numero = "", "0", Ado_datos.Recordset!doc_numero)
      Else
        If Ado_datos.Recordset!unidad_codigo = "DNINS" Or Ado_datos.Recordset!unidad_codigo = "DINSS" Or Ado_datos.Recordset!unidad_codigo = "DINSB" Or Ado_datos.Recordset!unidad_codigo = "DINSC" Then
            Ado_datos16.Recordset!doc_numero = IIf(Ado_datos.Recordset!doc_numero = "", "0", Ado_datos.Recordset!doc_numero)
        Else
            Ado_datos16.Recordset!doc_numero = IIf(txtDoc = "", "0", txtDoc)
        End If
      End If
      Ado_datos16.Recordset!doc_codigo_crono = "R-360"
      Ado_datos16.Recordset!doc_numero_crono = IIf(txtDoc = "", "0", txtDoc)    'Ado_datos.Recordset("venta_codigo")
      Ado_datos16.Recordset!poa_codigo = "3.1.2"
      Ado_datos16.Recordset!cobranza_fecha_prog = DTPFechaProg.Value           'Fecha Programada de Cobranza
      Ado_datos16.Recordset!trans_codigo = VAR_EMISION
      Ado_datos16.Recordset!estado_codigo = "REG"
      Ado_datos16.Recordset!usr_codigo = glusuario
      Ado_datos16.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
      Ado_datos16.Recordset!hora_registro = Format(Time, "hh:mm:ss")
      Ado_datos16.Recordset.Update
    db.CommitTrans
  If swnuevo = 1 Then
    'Call abre_solicitud_lista
    'rc_Cobranza.Requery
    'Ado_datos16.Refresh
    'Ado_datos16.Recordset.MoveLast
  End If
' db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.unimed_codigo = ao_ventas_cabecera.unimed_codigo_cobr,ao_ventas_cabecera.venta_monto_total_bs = aa_datos_venta_cabecera.tot_bs,ao_ventas_cabecera.venta_monto_total_dol = aa_datos_venta_cabecera.tot_dol,ao_ventas_cabecera.venta_cantidad_total = aa_datos_venta_cabecera.COBRANZA,ao_ventas_cabecera.venta_saldo_p_cobrar_bs = aa_datos_venta_cabecera.tot_bs - ao_ventas_cabecera.venta_monto_cobrado_bs From aa_datos_venta_cabecera WHERE aa_datos_venta_cabecera.venta_codigo = ao_ventas_cabecera.venta_codigo and ao_ventas_cabecera.venta_codigo = " & NumComp & ""
' db.Execute "update aa_datos_venta_detalle set aa_datos_venta_detalle.venta_precio_unitario_bs = (aa_datos_venta_cabecera.tot_bs / aa_datos_venta_cabecera.COBRANZA) / aa_datos_venta_detalle_equipos_cant.eqp, aa_datos_venta_detalle.venta_precio_total_bs = aa_datos_venta_cabecera.tot_bs / aa_datos_venta_detalle_equipos_cant.eqp, aa_datos_venta_detalle.venta_precio_total_dol = aa_datos_venta_cabecera.tot_dol / aa_datos_venta_detalle_equipos_cant.eqp, aa_datos_venta_detalle.venta_precio_unitario_dol = ((aa_datos_venta_cabecera.tot_bs / aa_datos_venta_cabecera.COBRANZA) * " & GlTipoCambioOficial & ") / aa_datos_venta_detalle_equipos_cant.eqp From aa_datos_venta_cabecera, aa_datos_venta_detalle_equipos_cant where aa_datos_venta_detalle.venta_codigo = aa_datos_venta_cabecera.venta_codigo and aa_datos_venta_detalle.venta_codigo = aa_datos_venta_detalle_equipos_cant.venta_codigo and aa_datos_venta_cabecera.venta_codigo =  " & NumComp & ""
' db.Execute "update ao_ventas_cabecera set venta_cantidad_total = " & txtCantCobr.Text & " where venta_codigo = " & NumComp & ""
 
    'Momentaneamente SUSPENDIDO
'    db.Execute "tp_actualiza_datos_venta " & NumComp
    Call ABRIR_TABLAS_AUX
    Call ABRIR_DETALLE
    ' db.Execute "update aa_datos_venta_detalle set aa_datos_venta_detalle.venta_precio_unitario_bs = aa_datos_venta_cabecera.tot_bs / aa_datos_venta_cabecera.COBRANZA ,aa_datos_venta_detalle.venta_precio_total_bs = aa_datos_venta_cabecera.tot_bs,aa_datos_venta_detalle.venta_precio_total_dol = aa_datos_venta_cabecera.tot_dol,aa_datos_venta_detalle.venta_precio_unitario_bs = (aa_datos_venta_cabecera.tot_bs / aa_datos_venta_cabecera.COBRANZA) * " & GlTipoCambioOficial & "From aa_datos_venta_cabecera where aa_datos_venta_detalle.venta_codigo = aa_datos_venta_cabecera.venta_codigo and aa_datos_venta_cabecera.venta_codigo =  " & NumComp & ""
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    FraNavega.Enabled = True
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = True
    'BtnImprimir1.Visible = True
    'BtnImprimir4.Visible = True
    FrmDetalle.Visible = True
    FrmCobranza.Visible = True
    FrmCobros.Enabled = False
    TxtCobrador.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
    swnuevo = 0
    gestion0 = glGestion
    'gestion0 = Ado_datos.Recordset("ges_gestion")
    'NumComp = Ado_datos.Recordset("correl_venta")
    nroventa = Ado_datos.Recordset("venta_codigo")

'        DTPFechaProg.Visible = True
'        DTPFechaCobro.Visible = False
'        DTPFechaConf.Visible = True
'        DTPFechaProg.Enabled = True
    TxtMonto.Enabled = True
    TxtDsctoTot.Enabled = True
    TxtObs.Enabled = True

     If OptFilGral1.Value = True And Ado_datos.Recordset!estado_cancelado = "N" Then
        'Call OptFilGral2_Click
        Call OptFilGral1_Click        'Pendientes
     Else
        'Call OptFilGral1_Click
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
     'VAR_SW = ""
        rs_datos.Find "venta_codigo = " & NumComp & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
     'VAR_SW = ""
        rs_datos.MoveLast
     End If
    If (DtgCobro.SelBookmarks.Count <> 0) Then
        DtgCobro.SelBookmarks.Remove 0
     End If
     If Ado_datos16.Recordset.RecordCount > 0 Then
     'VAR_SW = ""
        rs_datos16.Find "cobranza_prog_codigo = " & VAR_COBRANZA & "   ", , , 1
        DtgCobro.SelBookmarks.Add (rs_datos16.Bookmark)
     Else
     'VAR_SW = ""
        rs_datos16.MoveLast
     End If

End Sub

Private Sub BtnImprimir2_Click()
  ' INI - SOLICITUD DE ANULACION DE FACTURA
    If Ado_datos16.Recordset.RecordCount > 0 Then
       If Ado_datos16.Recordset!estado_ac = "APR" Then
            'MsgBox "Ya fue solicitada la ANULACION la Factura... , verifique si es la cuota correcta ...", vbYesNo, "Confirmando"      '+ rs_aux15!cobranza_nro_factura
            sino = MsgBox("Ya fue solicitada la ANULACION la Factura...¿Desea Re-Imprimir la Solicitud?.", vbYesNo, "Confirmando")
            If sino = vbYes Then
                'Set rs_aviso_cob = New ADODB.Recordset
                'If rs_aviso_cob.State = 1 Then rs_aviso_cob.Close
                'rs_aviso_cob.Open "Select * from fc_correl where tipo_tramite = 'aviso_cob'", db, adOpenStatic
                'If rs_aviso_cob.RecordCount > 0 Then
                    aviso_cob = Ado_datos16.Recordset!correl_ac
                    Dim iResult As Variant  ', i%, y%
                    cry_ac.WindowShowRefreshBtn = True
                    cry_ac.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
                    cry_ac.StoredProcParam(1) = Me.Ado_datos16.Recordset!cobranza_prog_codigo
                    cry_ac.ReportFileName = App.Path & "\reportes\ventas\ar_solicita_ANL_factura.rpt"
                    'cry_ac.ReportFileName = App.Path & "\reportes\ventas\ar_solicita_ANL_factura_PRUEBA.rpt"
                    '
                    cry_ac.Formulas(1) = "correl = '" & aviso_cob & "' "
                    iResult = cry_ac.PrintReport
                    If iResult <> 0 Then MsgBox cry_ac.LastErrorNumber & " : " & cry_ac.LastErrorString, vbCritical, "Error de impresión"
                'End If
            End If
            Exit Sub
       End If
       If rs_aux15.State = 1 Then rs_aux15.Close
       rs_aux15.Open "SELECT * FROM av_ventas_cuotas_con_factura WHERE correl_prog = " & Ado_datos16.Recordset!correl_prog & " ", db, adOpenKeyset, adLockOptimistic, adCmdText
       If rs_aux15.RecordCount > 0 Then
            sino = MsgBox("¿Esta seguro de solicitar la ANULACION de la Factura Nro." + rs_aux15!cobranza_nro_factura, vbYesNo, "Confirmando")
            If sino = vbYes Then
                FraAnula.Visible = True
                'Else
                '    sino = MsgBox("Ya se genero un Aviso de cobranza anteriormente, ¿Desea reimprimir este aviso?", vbYesNo, "Confirmando")
                '    cry_ac.Formulas(1) = "correl = '" & Ado_datos16.Recordset!correl_ac & "' "
            End If
       Else
            MsgBox "No se puede PROCESAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
       End If
    End If
  ' FIN - SOLICITUD DE ANULACION DE FACTURA

'   consulta = "aa_facs_sin_cobrar '" & Ado_datos.Recordset!edif_codigo & "'"
'   If rs_aux15.State = 1 Then rs_aux15.Close
'   rs_aux15.Open consulta, db, adOpenKeyset, adLockOptimistic, adCmdText
'   If rs_aux15.RecordCount > 1 Then
'   sino = MsgBox("Este cliente tiene mas deudas, ¿Desea imprimir el detalle?", vbYesNo, "Confirmando")
'    If sino = vbYes Then
'      cry_deuda.ReportFileName = App.Path & "\reportes\ventas\aa_facs_sin_cobrar.rpt"
'       cry_deuda.WindowShowRefreshBtn = True
'      cry_deuda.StoredProcParam(0) = Me.Ado_datos.Recordset!edif_codigo
'      iResult = cry_deuda.PrintReport
'  If iResult <> 0 Then MsgBox cry_deuda.LastErrorNumber & " : " & cry_ac.LastErrorString, vbCritical, "Error de impresión"
'
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If

'AVISO DE COBRANZA
'If Ado_datos16.Recordset.RecordCount > 0 Then
'  sino = MsgBox("¿Esta seguro de imprimir el aviso de cobranza?", vbYesNo, "Confirmando")
'
'   If sino = vbYes Then
'    Set rs_aviso_cob = New ADODB.Recordset
'    If rs_aviso_cob.State = 1 Then rs_aviso_cob.Close
'    rs_aviso_cob.Open "Select * from fc_correl where tipo_tramite = 'aviso_cob'", db, adOpenStatic
'
'    If rs_aviso_cob.RecordCount > 0 Then
'    cry_ac.ReportFileName = App.Path & "\reportes\ventas\ar_aviso_cobranza.rpt"
'
'      If Ado_datos16.Recordset!estado_ac = "REG" Then
'       aviso_cob = rs_aviso_cob!numero_correlativo + 1
'       db.Execute "update fc_correl set numero_correlativo = " & aviso_cob & " where tipo_tramite = 'aviso_cob'"
'       Ado_datos16.Recordset!correl_ac = aviso_cob
'       Ado_datos16.Recordset!estado_ac = "APR"
'       Ado_datos16.Recordset.Update
'       cry_ac.Formulas(1) = "correl = '" & aviso_cob & "' "
''       sino = MsgBox("¿Desea enviar a facturacion este registro?", vbYesNo, "Confirmando")
''        If sino = vbYes Then
''         Call BtnAprobar2_Click
''        End If
'      Else
'      sino = MsgBox("Ya se genero un Aviso de cobranza anteriormente, ¿Desea reimprimir este aviso?", vbYesNo, "Confirmando")
'      cry_ac.Formulas(1) = "correl = '" & Ado_datos16.Recordset!correl_ac & "' "
'      End If
'
'    End If
'
'    Dim iResult As Variant  ', i%, y%
'    'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R-105_kardex.rpt"
'    'cry_ac.ReportFileName = App.Path & "\reportes\ventas\ar_aviso_cobranza.rpt"
'    cry_ac.WindowShowRefreshBtn = True
'    cry_ac.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
'    cry_ac.StoredProcParam(1) = Me.Ado_datos16.Recordset!cobranza_prog_codigo
'
'    'Literal por el Total de la Compra
'    iResult = cry_ac.PrintReport
'    If iResult <> 0 Then MsgBox cry_ac.LastErrorNumber & " : " & cry_ac.LastErrorString, vbCritical, "Error de impresión"
'   End If
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
'   consulta = "aa_facs_sin_cobrar '" & Ado_datos.Recordset!edif_codigo & "'"
'   If rs_aux15.State = 1 Then rs_aux15.Close
'   rs_aux15.Open consulta, db, adOpenKeyset, adLockOptimistic, adCmdText
'   If rs_aux15.RecordCount > 1 Then
'   sino = MsgBox("Este cliente tiene mas deudas, ¿Desea imprimir el detalle?", vbYesNo, "Confirmando")
'    If sino = vbYes Then
'      cry_deuda.ReportFileName = App.Path & "\reportes\ventas\aa_facs_sin_cobrar.rpt"
'       cry_deuda.WindowShowRefreshBtn = True
'      cry_deuda.StoredProcParam(0) = Me.Ado_datos.Recordset!edif_codigo
'      iResult = cry_deuda.PrintReport
'
'  If iResult <> 0 Then MsgBox cry_deuda.LastErrorNumber & " : " & cry_ac.LastErrorString, vbCritical, "Error de impresión"
'
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
'End If
  
'  'SOLICITUD DE FACTURACION
'    'ElseIf optRep0010.Value = True And opt_4.Value = True Then
'        CryUnidad.ReportFileName = App.Path & "\Reportes\Ventas\fr_solicitud_factura_cobrador.rpt"
'        titulo2 = "MODULO COBRANZAS"
'        subtitulo2 = "SOLICITUD DE FACTURACION - R-110"
'        CryUnidad.Formulas(2) = "Titulo = '" & titulo2 & "'"
'        CryUnidad.Formulas(3) = "SubTitulo = '" & subtitulo2 & "'"
'        iResult = CryUnidad.PrintReport
'        If iResult <> 0 Then
'            MsgBox CryUnidad.LastErrorNumber & " : " & CryUnidad.LastErrorString, vbCritical + vbOKOnly, "Error..."
'     '   End If
End Sub

Private Sub BtnAnlDetalle_Click()
 If Ado_datos.Recordset!estado_codigo = "REG" Then
   sino = MsgBox("Está seguro de ANULAR este registro", vbYesNo + vbQuestion, "Atención ...")
   If sino = vbYes Then
'     ado_datos14.Recordset.Delete
'     ado_datos14.Recordset.Update
'     rs_datos14.Requery
'     ado_datos14.Refresh
'     'cerea
'     ado_datos14.Refresh
      'db.Execute "update ao_ventas_detalle set ao_ventas_detalle.estado_codigo = 'ANL' Where ao_ventas_detalle.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_detalle.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_detalle.venta_codigo_det = " & ado_datos14.Recordset("venta_codigo_det") & " "
      db.Execute "update ao_ventas_detalle set ao_ventas_detalle.estado_codigo = 'ANL' Where ao_ventas_detalle.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_detalle.venta_codigo_det = " & Ado_datos14.Recordset("venta_codigo_det") & " "
   End If
  Else
    MsgBox "Los Bienes del registro Aprobado o Anulado, NO pueden ser ANULADOS !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub valida_campos2()
  'FALTA VERIFICAR SI EXISTE EN ORGANIZACION DE ZONAS...
 If Ado_datos.Recordset!unidad_codigo = "DNMAN" Or Ado_datos.Recordset!unidad_codigo = "DMANS" Or Ado_datos.Recordset!unidad_codigo = "DMANB" Or Ado_datos.Recordset!unidad_codigo = "DMANC" Then
    If dtc_codigo7.Text = "" Or dtc_codigo7.Text = "0" Then
      MsgBox "Debe Elejir: Zona Piloto !! , Vuelva a Intentar por favor ...", vbExclamation, "Atención"
      VAR_VALD = "ERR"
      Exit Sub
    End If

    'Responsable/Supervisor del Servicio Técnico:
    If dtc_codigo4.Text = "" Then
      MsgBox "Debe Elejir: Responsable/Supervisor del Servicio Técnico !! , Vuelva a Intentar ...", vbExclamation, "Atención"
      VAR_VALD = "ERR"
      Exit Sub
    End If
'  If LblOrden = "" Or LblOrden = "0" Then
'    MsgBox "Debe Registrar previamente el Edificio, en la Organizacion de Zonas Piloto ...!! , Vuelva a Intentar ...", vbExclamation, "Atención"
'    VAR_VAL = "ERR"
'    VAR_VALD = "ERR"
'    Exit Sub
'  End If

  'TOTAL PERIODOS
  If txt_cant = "" Then
    MsgBox "Debe Registrar: Total Periodos !! , Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VALD = "ERR"
    Exit Sub
  End If
  'PERIODICIDAD
  If cmd_unimed_tec = "" Then
    MsgBox "Debe Elejir: Periodicidad !! , Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VALD = "ERR"
    Exit Sub
  End If
  If IsNull(lbl_fecha_fin) Then
    MsgBox "Debe Elejir Fecha Inicio Cronograma !! , Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VALD = "ERR"
    Exit Sub
  End If
  'FECHA INICIO Y FIN DEL CRONOGRAMA
  'If Val(Format(lbl_fecha_fin.Caption, "dd/mm/yyyy")) <= Val(Format(lbl_fecha_ini.Caption, "dd/mm/yyyy")) Then
  If CDate(Format(lbl_fecha_fin, "dd/mm/yyyy")) <= CDate(Format(lbl_fecha_ini, "dd/mm/yyyy")) Then
    MsgBox "La Fecha de Inicio debe ser MENOR a la Fecha de Fin del Cronograma!! , Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VALD = "ERR"
    Exit Sub
  End If
  If cmb_mes_ini_tec = "" Then
    MsgBox "Debe Elejir: Mes Inicio Cronograma !! , Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VALD = "ERR"
    Exit Sub
  End If
  Select Case RTrim(cmb_mes_ini_tec)
        Case "ENERO"
            VAR_MESINI2 = 1
        Case "FEBRERO"
            VAR_MESINI2 = 2
        Case "MARZO"
            VAR_MESINI2 = 3
        Case "ABRIL"
            VAR_MESINI2 = 4
        Case "MAYO"
            VAR_MESINI2 = 5
        Case "JUNIO"
            VAR_MESINI2 = 6
        Case "JULIO"
            VAR_MESINI2 = 7
        Case "AGOSTO"
            VAR_MESINI2 = 8
        Case "SEPTIEMBRE"
            VAR_MESINI2 = 9
        Case "OCTUBRE"
            VAR_MESINI2 = 10
        Case "NOVIEMBRE"
            VAR_MESINI2 = 11
        Case "DICIEMBRE"
            VAR_MESINI2 = 12
  End Select
  If Month(CDate(Format(lbl_fecha_ini, "dd/mm/yyyy"))) <> 12 And VAR_MESINI2 <> 1 Then
    If Val(VAR_MESINI2) < Month(CDate(Format(lbl_fecha_ini, "dd/mm/yyyy"))) Then
        MsgBox "El MES de Inicio del Crono. NO puede ser MENOR al de la Fecha de Inicio del Cronograma!! , Vuelva a Intentar ...", vbExclamation, "Atención"
        VAR_VALD = "ERR"
        Exit Sub
    End If
  End If
 End If
End Sub

Private Sub CmdGrabaDet_Click()
' If Ado_datos.Recordset!estado_codigo = "REG" Then
'  VAR_VALD = "OK"
'  Call valida_campos2
'  If VAR_VALD = "ERR" Then
'      Exit Sub
'  Else
'        NumComp = Ado_datos.Recordset!venta_codigo
'        If Option11.Value = True Then
'            'PROGRAMAR en Meses PARES
'            VAR_IMPAR = "2"
'        Else
'          ' Programar Meses IMPARES
'            VAR_IMPAR = "1"
'        End If
'        Set rs_datos10 = New ADODB.Recordset
'        If rs_datos10.State = 1 Then rs_datos10.Close
'        rs_datos10.Open "Select * from tc_zona_piloto_edif where edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "' ", db, adOpenStatic
'        If rs_datos10.RecordCount > 0 Then
'            VAR_ZONA = rs_datos10!zpiloto_codigo
'            DIA_ORDEN = rs_datos10!zona_edif_orden
'        Else
'        ' wwwwwwwwwwwwwwwwwwwwwwwwwww Graba tc_zona_piloto_edif
'          Set rs_aux18 = New ADODB.Recordset
'          If rs_aux18.State = 1 Then rs_aux18.Close
'          SQL_FOR = "Select ISNULL(max(zona_edif_orden),0) as Orden from tc_zona_piloto_edif where zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' "
'          rs_aux18.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'          If rs_aux18.RecordCount > 0 Then
'              VAR_ORDEN = IIf(IsNull(rs_aux18!Orden), 1, rs_aux18!Orden + 1)
'          Else
'              VAR_ORDEN = 1
'          End If
'          'db.Execute "SELECT Txt_campo1.Text  = ISNULL(MAX(zona_edif_orden),0)+1 FROM tc_zona_piloto_edif where zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' "
'          db.Execute "insert into  tc_zona_piloto_edif(zpiloto_codigo, edif_codigo, zona_edif_orden, zona_codigo, beneficiario_codigo, beneficiario_codigo_rep, beneficiario_codigo_cobr, zorden_cambio, observaciones, estado_codigo, fecha_registro, usr_codigo, mes_par_impar) " & _
'          "values (" & dtc_codigo7.Text & ", '" & Ado_datos.Recordset!edif_codigo & "', " & VAR_ORDEN & ", '0',         '0',                '0',                    '0',                    '0',            '-',            'REG',      GETDATE(),      'ADMIN', '" & VAR_IMPAR & "')"
'
'          db.Execute "update gc_edificaciones set tomado= 'S' where edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "' "
'
'            VAR_ZONA = dtc_codigo7.Text
'            DIA_ORDEN = VAR_ORDEN
'        End If
'      ' wwwwwwwwwwwwwwwwwwwwwwwwwww Graba tc_zona_piloto_edif
'
'        'WWWWW GENERA CRONOGRAMA DIARIO WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
''        FrmABMDet2.Enabled = False
''        FrmABMDet.Enabled = False
''        fraOpciones.Enabled = False
''        'Screen.MousePointer = vbHourglass
'
'        FInicio = Format(Ado_datos.Recordset!fecha_inicio_tec, "dd/mm/yyyy")
'        FFin = Format(Ado_datos.Recordset!fecha_fin_tec, "dd/mm/yyyy")
'        CANTOT = IIf(IsNull(Ado_datos.Recordset!cantidad_periodos_tec), 12, Ado_datos.Recordset!cantidad_periodos_tec)
'        VAR_MED = IIf(IsNull(Ado_datos.Recordset!unimed_codigo_tec), "MES", Ado_datos.Recordset!unimed_codigo_tec)
'        'VAR_ZONA = dtc_codigo7.Text                         'Ado_datos.Recordset!zpiloto_codigo
'        VAR_UNITEC = Ado_datos.Recordset!unidad_codigo      'Ado_datos.Recordset!unidad_codigo_tec
'        VAR_TECCOD = Ado_datos.Recordset!solicitud_codigo    'Ado_datos.Recordset!tec_plan_codigo
'        VAR_EDIF = RTrim(dtc_desc3.Text)
'        VAR_LUN = "SI"                                        'Ado_datos.Recordset!lunes_cambia
'        VAR_PRIM = "SI"                                        'Ado_datos.Recordset!primero_mes
'
'        VAR_EMES = "Error: No se encontró el Mes de Inicio del Cronograma, verifique y vuelva a intentar..."
'        dtc_codigo5.Text = "0"
'        'WWWWWWWWWWWWWWWWWWWWWWWWWWW OJOOOOOOOOOOOO CARGAR TO_CRONOGRAMA_VS_VENTAS  -   JQA 2022-AGO
'        'ABRIR TO_CRONOGRAMA_MENSUAL    -
'        Set rs_aux16 = New ADODB.Recordset
'        If rs_aux16.State = 1 Then rs_aux16.Close
'        rs_aux16.Open "Select * from to_cronograma_mensual where ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' and  fmes_correl = " & IIf(IsNull(Ado_datos.Recordset!mes_codigo), 1, Ado_datos.Recordset!mes_codigo) & " and zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & "  ", db, adOpenStatic        '
'        If rs_aux16.RecordCount > 0 Then
'        End If
'        'db.Execute "update to_cronograma_ventas set fmes_plan= 'S' where edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "' "
'
'      Set rs_datos9 = New ADODB.Recordset
'      If rs_datos9.State = 1 Then rs_datos9.Close
'      'rs_datos9.Open "Select * from to_cronograma WHERE estado_detalle = 'APR' AND unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   ", db, adOpenStatic
'      rs_datos9.Open "Select * from ao_ventas_cabecera WHERE estado_crono = 'APR' AND unidad_codigo = '" & VAR_UNITEC & "' and solicitud_codigo = " & VAR_TECCOD & "   ", db, adOpenStatic
'      If rs_datos9.RecordCount > 0 Then
'           MsgBox "El Cronograma ya existe, verifique y vuelva a intentar ...", vbExclamation, "Validación de Registro"
'           'Frame2.Visible = False
'           'ProgressBar1.Visible = False
'           Exit Sub
'      Else
'        ' estado_activo = 'ANL'
'        'DIA_ORDEN = Ado_datos.Recordset!zona_edif_orden
'        NumComp = Ado_datos.Recordset!venta_codigo
'        MControl = Ado_datos.Recordset!mes_inicio_crono_tec
'        VAR_MESINI2 = IIf(IsNull(Ado_datos.Recordset!mes_inicio_crono_nro), 1, Ado_datos.Recordset!mes_inicio_crono_nro)
'        FControl = FInicio
'        CONT4 = 0
'        UMED_NRO = Ado_datos.Recordset!unimed_codigo_nro     ' Fijo MES=1, BMES=2, TMES=3
'        VAR_CONT = 1
'        VAR_MES = Month(FControl)
'        UMED_NRO2 = VAR_MESINI2      'UMED_NRO
'        'Frame2.Visible = True
'        'ProgressBar1.Visible = True
''        With ProgressBar1
''            .Max = CANTOT     'rs_datos9.RecordCount
''            .Min = 0
''            .Value = 0
''        End With
'        gestion0 = Year(FControl)
'        While CANTOT >= VAR_CONT And FFin >= FControl   'UNIMED veces (12, 24, etc.)
'            If UMED_NRO2 = 13 And gestion0 <> Year(FControl) Then
'                UMED_NRO2 = 1
'                'gestion0 = Year(FControl)
'             End If
'          gestion0 = Year(FControl)
'
'          CONT3 = 0
'          If VAR_MES = UMED_NRO2 Then
'             Set rs_aux1 = New ADODB.Recordset
'             'rs_aux1.Open "select * from to_cronograma_detalle where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   ", db, adOpenKeyset, adLockBatchOptimistic
'             rs_aux1.Open "select * from ao_ventas_detalle where venta_codigo = " & NumComp & " AND par_codigo = '43340'  ", db, adOpenKeyset, adLockBatchOptimistic
'             If rs_aux1.RecordCount > 0 Then
'                 ' De acuerdo a la cantidad de equipos
'                 'NumComp = IIf(IsNull(rs_aux1!bien_cantidad_por_empaque), 2, rs_aux1!bien_cantidad_por_empaque) / 2
'                 VAR_CANTCRO = IIf(IsNull(rs_aux1!bien_cantidad_por_empaque), 2, rs_aux1!bien_cantidad_por_empaque)
'                 rs_aux1.MoveFirst
'                 While Not rs_aux1.EOF
'                     Set rs_aux2 = New ADODB.Recordset
'                     If rs_aux2.State = 1 Then rs_aux2.Close
'                     rs_aux2.Open "select * from to_cronograma_mensual where ges_gestion = '" & gestion0 & "' and fmes_correl = " & VAR_MES & " and zpiloto_codigo = " & VAR_ZONA & "    ", db, adOpenKeyset, adLockOptimistic
'                     If rs_aux2.RecordCount > 0 Then
'                         VAR_AUX2 = rs_aux2!fmes_plan
'                         VAR_COD0 = 0
'                         'UMED_NRO2 = 0
'                         Set rs_aux3 = New ADODB.Recordset
'                         If rs_aux3.State = 1 Then rs_aux3.Close
'                         rs_aux3.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & "   ", db, adOpenKeyset, adLockBatchOptimistic
'                         If rs_aux3.RecordCount > 0 Then
'                             rs_aux3.MoveFirst
'                             While Not rs_aux3.EOF
'                                If cmb_dia.Text = "AUTOMATICO" And dtc_codigo5.Text = "0" Then
'                                   Set rs_aux4 = New ADODB.Recordset
'                                   If rs_aux4.State = 1 Then rs_aux4.Close
'                                   rs_aux4.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & "  AND estado_activo = 'REG'  ", db, adOpenKeyset, adLockBatchOptimistic
'                                   If rs_aux4.RecordCount > 0 Then
'                                    If VAR_COD0 < NumComp And rs_aux3!estado_activo = "REG" Then
'                                       db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux1!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                       db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
'                                       db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                       db.Execute "update to_cronograma_diario set nro_total_horas = " & NumComp & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                       'VAR_COD0 = VAR_COD0 + 1
'                                       VAR_COD0 = VAR_COD0 + NumComp
'                                       CONT3 = 1
'                                       'If VAR_MES Then
'                                       VAR_EMES = "NADA"
'                                       'End If
''                                       If VAR_LUN = "SI" Or VAR_PRIM = "SI" Then
''                                          'TODOS LOS LUNES O EL 1RO. DE CADA MES
''                                          If (rs_aux3!dia_nombre = "LUNES" Or rs_aux3!dia_correl = "1") And rs_aux3!hora_ingreso = "08:00" Then
''                                             rs_aux3.MoveNext
''                                             db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux1!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                                             db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                                             db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                                             VAR_COD0 = VAR_COD0 + 1
''                                             CONT3 = 1
''                                          End If
''                                       End If
'                                       'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
'                                       db.Execute "Update ao_ventas_cabecera Set estado_crono = 'APR' Where VENTA_CODIGO = " & Ado_datos.Recordset!venta_codigo & " "             'unidad_codigo = '" & VAR_UNITEC & "' and solicitud_codigo = " & VAR_TECCOD & "   "
'                                    End If
'                                   Else
'                                        MsgBox "Ya no existen horarios laborales LIBRES, para la gestion: " & gestion0 & ", el Mes: " & VAR_MES & " y la Zona: " & VAR_ZONA, vbInformation, "Información"
'                                        rs_aux3.MoveLast
'                                   End If
'                                Else
'                                   If cmb_dia.Text = rs_aux3!dia_nombre And dtc_codigo5.Text = "0" Then
'    '                                         If rs_aux3!dia_nombre = "SÁBADO" Or rs_aux3!dia_nombre = "DOMINGO" Or rs_aux3!estado_activo = "ANL" Then
'    '                                            db.Execute "update to_cronograma_diario set observaciones = 'DIA NO LABORABLE' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'    '                                            db.Execute "update to_cronograma_diario set estado_activo = 'ANL' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'    '                                         Else
'                                     If VAR_COD0 < NumComp Then     'And rs_aux3!estado_activo = "REG"
'                                         db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux1!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                         db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
'                                         db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                         VAR_COD0 = VAR_COD0 + 1
'                                         CONT3 = 1
'                                         'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
'                                         db.Execute "Update ao_ventas_detalle Set estado_crono = 'APR' Where unidad_codigo = '" & VAR_UNITEC & "' and solicitud_codigo = " & VAR_TECCOD & "   "
'                                     End If
'                                   End If
'                                   If dtc_codigo5.Text = rs_aux3!horario_codigo Then
'                                     If VAR_COD0 < NumComp Then     'And rs_aux3!estado_activo = "REG"
'                                         db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux1!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                         db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
'                                         db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                         VAR_COD0 = VAR_COD0 + 1
'                                         CONT3 = 1
'                                         'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
'                                         db.Execute "Update ao_ventas_detalle Set estado_crono = 'APR' Where unidad_codigo = '" & VAR_UNITEC & "' and solicitud_codigo = " & VAR_TECCOD & "   "
'                                     End If
'                                   End If
'                                End If
'                                rs_aux3.MoveNext
'                                'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
'                             Wend
'                         End If
'                     End If
'                     rs_aux1.MoveNext
'                 Wend
'             End If
'             VAR_CONT = VAR_CONT + 1
'             UMED_NRO2 = UMED_NRO2 + UMED_NRO
''             ProgressBar1.Value = ProgressBar1.Value + 1
'          Else
'            'VAR_EMES = "Error: No se encontró el Mes de Inicio del Cronograma, verifique u vuelva a intentar..."
'          End If
'             Select Case VAR_MES
'                 Case 2
'                     If gestion0 = "2016" Or gestion0 = "2020" Or gestion0 = "2024" Or gestion0 = "2028" Then
'                         Dias_Mes = 29
'                     Else
'                         Dias_Mes = 28
'                     End If
'                 Case 1, 3, 5, 7, 8, 10, 12
'                     Dias_Mes = 31
'                 Case 4, 6, 9, 11
'                     Dias_Mes = 30
'             End Select
'             'rs_aux2!cobranza_fecha_prog = FControl
'             'rs_aux2!cobranza_fecha_cobro = FControl + 10
'             FControl = CDate(FControl) + Dias_Mes
'             VAR_MES = Month(FControl)
'             Select Case VAR_MED
'                Case "MES"    'MENSUAL
''                    UMED_NRO2 = VAR_CONT
''                    If (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) Then
''                        'UMED_NRO2 = (VAR_MES * UMED_NRO) - 1
''                        UMED_NRO2 = VAR_CONT
''                    Else
''                        UMED_NRO2 = VAR_MES * UMED_NRO
''                    End If
'                Case "BMES"    'BIMESTRAL
''                    If (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) Then
''                        UMED_NRO2 = (VAR_CONT * UMED_NRO) - 1
''                    Else
''                        UMED_NRO2 = VAR_CONT * UMED_NRO
''                    End If
'                Case "TMES"    'TRIMESTRAL
'                    'UMED_NRO2 = (UMED_NRO2 + UMED_NRO)
''                    If (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) Then
''                        UMED_NRO2 = (VAR_CONT * UMED_NRO) '- 2
''                        'UMED_NRO2 = (UMED_NRO2 + VAR_MESINI2)
''                    Else
''                        UMED_NRO2 = (VAR_CONT * UMED_NRO) - 1
''                    End If
''                    'UMED_NRO2 = 3
''                    If VAR_MES = UMED_NRO2 Then
''                        UMED_NRO2 = UMED_NRO2 + VAR_MESINI2
''                    End If
''                    UMED_NRO2 = VAR_CONT * UMED_NRO
'                Case "CMES"    'CUATRIMESTRAL
'                Case "QMES"    'CADA 5 MESES
'                Case "SMES"    'SEMESTRAL
'                    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'                Case "ANUAL"    'ANUAL
'             End Select
''             If VAR_MED = "TMES" And CONT3 = 1 Then
''                UMED_NRO2 = (VAR_CONT * UMED_NRO) - 2
''                VAR_CONT = VAR_CONT + 1
''             End If
''                If CONT3 = 1 And VAR_MED = "MES" And (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) And (UMED_NRO = 2) Then
''                    UMED_NRO2 = (VAR_MES * UMED_NRO) - 1
''                Else
''                    UMED_NRO2 = VAR_MES * UMED_NRO
''                End If
''                'If CONT3 = 1 And VAR_MED = "BMES" Then
''                If VAR_MED = "BMES" Then
''                    If (VAR_MESINI2 = 1 Or VAR_MESINI2 = 3 Or VAR_MESINI2 = 5 Or VAR_MESINI2 = 7 Or VAR_MESINI2 = 9 Or VAR_MESINI2 = 11) And (UMED_NRO = 2) Then
''                        UMED_NRO2 = (VAR_CONT * UMED_NRO) - 1
''                        VAR_CONT = VAR_CONT + 1
''                    Else
''                        UMED_NRO2 = VAR_CONT * UMED_NRO
''                        VAR_CONT = VAR_CONT + 1
''                    End If
''                End If
''             'End If
'        Wend
'
'        FrmABMDet2.Enabled = True
'        FrmABMDet.Enabled = True
'        fraOpciones.Enabled = True
''        Screen.MousePointer = vbDefault
'        If VAR_EMES = "NADA" Then
'            MsgBox "El Cronograma fue creado Satisfactoriamente ...", vbInformation, "Información"
''            ProgressBar1.Visible = False
''            Frame2.Visible = False
'        Else
'            MsgBox VAR_EMES, vbInformation, "Información"
'        End If
'        Call ABRIR_DETALLE
'      End If
'''      ProgressBar1.Visible = False
''      Frame2.Visible = False
'      'WWWWW GENERA CRONOGRAMA DIARIO (FIN)
'  End If
' Else
'        MsgBox "NO se puede generar un NUEVO CRONOGRAMA, en un Registro APROBADO o ANULADO !! ", vbExclamation, "Atención!"
' End If
'
End Sub

Private Sub Combo1_LostFocus()
    '----------------------------- GLOSA PARA FACTURAR
    If parametro = "DNMAN" Or parametro = "DMANS" Or parametro = "DMANB" Or parametro = "DMANC" Then
        txt_plazo.Text = "SERVICIO DE MANTENIMIENTO INTEGRAL - CUOTA No " + Str(VAR_COBRANZA)
    End If
    If parametro = "DNREP" Or parametro = "DREPS" Or parametro = "DREPB" Or parametro = "DREPC" Then
        txt_plazo.Text = "SERVICIO DE REPARACION, SEGÚN " + Ado_datos.Recordset!unidad_codigo_ant
    End If
    If parametro = "DNINS" Or parametro = "DINSB" Or parametro = "DINSC" Or parametro = "DINSS" Then
        txt_plazo.Text = "SERVICIO DE INSTALACION, SEGÚN " + Ado_datos.Recordset!unidad_codigo_ant
    End If
    If parametro = "DNMOD" Or parametro = "DMODS" Or parametro = "DMODB" Or parametro = "DMODC" Then
        txt_plazo.Text = "SERVICIO DE MODERNIZACION DE EQUIPOS - CUOTA Nº " + Str(VAR_COBRANZA)
    End If
End Sub

Private Sub Command2_Click()
'ACTUALIZA LITERALES ----------------------------------- WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW Falta habilitar BOton
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "select * From ao_ventas_cabecera "
    'queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    'rs_datos.Sort = "SOLICITUD_codigo"
    'Set rs_datos = rs_datos.DataSource
    rs_datos.MoveFirst
    sino = rs_datos.RecordCount
    While Not rs_datos.EOF
        Set rs_aux13 = New ADODB.Recordset
        If rs_aux13.State = 1 Then rs_aux13.Close
        rs_aux13.Open "select sum(venta_precio_total_bs) as total from ao_ventas_detalle where venta_codigo = '" & rs_datos!venta_codigo & "' ", db, adOpenKeyset, adLockOptimistic
        If rs_aux13!total <> "NULL" Then
              rs_datos!literal_a = Literal(rs_aux13!total)
        End If
      
'      Set rs_aux13 = New ADODB.Recordset
'        If rs_aux13.State = 1 Then rs_aux13.Close
'         rs_aux13.Open "select sum(venta_precio_total_bs) as total from ao_ventas_detalle where venta_codigo = '" & rs_datos!venta_codigo & "' and almacen_tipo ='A'", db, adOpenKeyset, adLockOptimistic
'            If rs_aux13!total <> "NULL" Then
'              rs_datos!literal_a = Literal(rs_aux13!total)
'            End If
'
'      Set rs_aux13 = New ADODB.Recordset
'        If rs_aux13.State = 1 Then rs_aux13.Close
'         rs_aux13.Open "select sum(venta_precio_total_bs) as total from ao_ventas_detalle where venta_codigo = '" & rs_datos!venta_codigo & "' and almacen_tipo ='I'", db, adOpenKeyset, adLockOptimistic
'            If rs_aux13!total <> "NULL" Then
'              rs_datos!literal_i = Literal(rs_aux13!total)
'            End If
'
'      Set rs_aux13 = New ADODB.Recordset
'        If rs_aux13.State = 1 Then rs_aux13.Close
'         rs_aux13.Open "select sum(venta_precio_total_bs) as total from ao_ventas_detalle where venta_codigo = '" & rs_datos!venta_codigo & "' and almacen_tipo ='R'", db, adOpenKeyset, adLockOptimistic
'            If rs_aux13!total <> "NULL" Then
'              rs_datos!literal_r = Literal(rs_aux13!total)
'            End If
'
'  Set rs_aux13 = New ADODB.Recordset
'        If rs_aux13.State = 1 Then rs_aux13.Close
'         rs_aux13.Open "select sum(venta_precio_total_bs) as total from ao_ventas_detalle where venta_codigo = '" & rs_datos!venta_codigo & "' and almacen_tipo ='H'", db, adOpenKeyset, adLockOptimistic
'            If rs_aux13!total <> "NULL" Then
'              rs_datos!literal_h = Literal(rs_aux13!total)
'            End If
        
        rs_datos.Update
        rs_datos.MoveNext
    Wend
End Sub

Private Sub BtnImprimir4_Click()
 If Ado_datos.Recordset.RecordCount > 0 Then
      If Ado_datos14.Recordset.RecordCount > 0 Then
        Dim iResult As Variant, i%, Y%
        Dim co As New ADODB.Command

    '    Dim rs As New ADODB.Recordset
    '    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    '            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly
    '    i = 1
    '    y = 1
        CryV01.ReportFileName = App.Path & "\reportes\Tecnico\tr_recibo_devolucion.rpt"
        'CryV01.WindowShowRefreshBtn = True
        CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
        'CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
     Else
        MsgBox "No se puede Imprimir. Debe registrar datos... " & FrmDetalle.Caption, , "Atención"
     End If
   Else
        MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
   End If

End Sub

'Private Sub CmdDetCabeza_Click()
'    fraOpciones.Visible = False
'    FrmDetalle.Visible = True
'    FrmCobranza.Visible = True
'    FraNavega.Enabled = False
'    If Not (adoDetalleSolicitud.Recordset.BOF) Then adoDetalleSolicitud.Recordset.MoveFirst
'End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_aux2.BoundText
    dtc_desc2.BoundText = dtc_aux2.BoundText
    Dtc_deudor2.BoundText = dtc_aux2.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_aux4.BoundText
    dtc_desc4.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_aux7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_aux7.BoundText
    dtc_desc7.BoundText = dtc_aux7.BoundText
End Sub

Private Sub dtc_benef2A_Click(Area As Integer)
    dtc_codigo2A.BoundText = dtc_benef2A.BoundText
    dtc_desc2A.BoundText = dtc_benef2A.BoundText
    dtc_email2A.BoundText = dtc_benef2A.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    dtc_aux2.BoundText = dtc_codigo2.BoundText
    Dtc_deudor2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo7_Click(Area As Integer)
    dtc_desc7.BoundText = dtc_codigo7.BoundText
    dtc_aux7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
    dtc_aux2.BoundText = dtc_desc2.BoundText
    Dtc_deudor2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc2A_LostFocus()
    '----------------------------- GLOSA PARA FACTURAR
    If parametro = "DNMAN" Or parametro = "DMANS" Or parametro = "DMANB" Or parametro = "DMANC" Then
        txt_plazo.Text = "SERVICIO DE MANTENIMIENTO INTEGRAL - CUOTA Nº " + Str(VAR_COBRANZA)
    End If
    If parametro = "DNREP" Or parametro = "DREPS" Or parametro = "DREPB" Or parametro = "DREPC" Then
        txt_plazo.Text = "SERVICIO DE REPARACION, SEGÚN " + Ado_datos.Recordset!unidad_codigo_ant
    End If
    If parametro = "DNINS" Or parametro = "DINSB" Or parametro = "DINSC" Or parametro = "DINSS" Then
        txt_plazo.Text = "SERVICIO DE INSTALACION, SEGÚN " + Ado_datos.Recordset!unidad_codigo_ant
    End If
    If parametro = "DNMOD" Or parametro = "DMODS" Or parametro = "DMODB" Or parametro = "DMODC" Then
        txt_plazo.Text = "SERVICIO DE MODERNIZACION DE EQUIPOS - CUOTA Nº " + Str(VAR_COBRANZA)
    End If

End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
    dtc_aux4.BoundText = dtc_desc4.BoundText
    VAR_BEN2 = dtc_codigo4.Text
End Sub

Private Sub dtc_desc4_LostFocus()
    dtc_codigo4.Text = VAR_BEN2
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc4A_LostFocus()
    '----------------------------- GLOSA PARA FACTURAR
    If parametro = "DNMAN" Or parametro = "DMANS" Or parametro = "DMANB" Or parametro = "DMANC" Then
        txt_plazo.Text = "SERVICIO DE MANTENIMIENTO INTEGRAL - CUOTA Nº " + Str(VAR_COBRANZA)
    End If
    If parametro = "DNREP" Or parametro = "DREPS" Or parametro = "DREPB" Or parametro = "DREPC" Then
        txt_plazo.Text = "SERVICIO DE REPARACION, SEGÚN " + Ado_datos.Recordset!unidad_codigo_ant
    End If
    If parametro = "DNINS" Or parametro = "DINSB" Or parametro = "DINSC" Or parametro = "DINSS" Then
        txt_plazo.Text = "SERVICIO DE INSTALACION, SEGÚN " + Ado_datos.Recordset!unidad_codigo_ant
    End If
    If parametro = "DNMOD" Or parametro = "DMODS" Or parametro = "DMODB" Or parametro = "DMODC" Then
        txt_plazo.Text = "SERVICIO DE MODERNIZACION DE EQUIPOS - CUOTA Nº " + Str(VAR_COBRANZA)
    End If
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub dtc_desc5_LostFocus()
'    If dtc_codigo11.Text = "C" Or dtc_codigo11.Text = "V" Then
'        If VAR_OS = "SI" Then
'            TxtConcepto.Text = Right(Trim(lbl_titulo.Caption) + " - " + Trim(Txt_campo2.Text) + " -Edificio: " + RTrim(dtc_desc3.Text), Len(Trim(lbl_titulo.Caption) + " - " + Trim(Txt_campo2.Text) + " -Edificio: " + RTrim(dtc_desc3.Text)) - 6)
'        Else
'            TxtConcepto.Text = Right(lbl_titulo.Caption + " -Edificio: " + RTrim(dtc_desc3.Text), Len(lbl_titulo.Caption + " -Edificio: " + RTrim(dtc_desc3.Text)) - 6)
'        End If
'    Else
'        If dtc_codigo11.Text = "E" Then
'            TxtConcepto.Text = lbl_titulo.Caption + " - " + RTrim(dtc_desc3.Text)
'            TxtPlazo.Text = 0
'        Else
'            TxtConcepto.Text = lbl_titulo.Caption + " - " + RTrim(dtc_desc3.Text)   '"VENTA DIRECTA AL CLIENTE"
'            TxtPlazo.Text = 0
'        End If
'    End If
    '-----------------------------
    If parametro = "DNMAN" Or parametro = "DMANS" Or parametro = "DMANB" Or parametro = "DMANC" Then
        TxtConcepto.Text = "Servicio de MANTENIMIENTO INTEGRAL. Edificio: " + dtc_desc3.Text + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
    End If
    If parametro = "DNREP" Or parametro = "DREPS" Or parametro = "DREPB" Or parametro = "DREPC" Then
        TxtConcepto.Text = "Servicio de REPARACIONES. Edificio: " + dtc_desc3.Text + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
    End If
    If parametro = "DNMOD" Or parametro = "DMODS" Or parametro = "DMODB" Or parametro = "DMODC" Then
        TxtConcepto.Text = "Servicio de MODERNIZACION de equipos. Edificio: " + dtc_desc3.Text + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
    End If
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc7.BoundText
    dtc_aux7.BoundText = dtc_desc7.BoundText
End Sub

Private Sub dtc_desc7_LostFocus()
    If dtc_aux7.Text = 1 Then
        'Programar Meses IMPARES y quitar PARES
        VAR_IMPAR = "1"
        Option11.Value = False
        Option10.Value = True
        Option11.Visible = False
        Option10.Visible = True
        LblParImpar = "MESES IMPARES"
    Else
        'PROGRAMAR en Meses PARES y quitar Mes IMPARES
        VAR_IMPAR = "2"
        Option11.Value = True
        Option10.Value = False
        Option11.Visible = True
        Option10.Visible = False
        LblParImpar = "MESES PARES"
    End If
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
End Sub

Private Sub Dtc_deudor2_Click(Area As Integer)
    dtc_codigo2.BoundText = Dtc_deudor2.BoundText
    dtc_aux2.BoundText = Dtc_deudor2.BoundText
    dtc_desc2.BoundText = Dtc_deudor2.BoundText
End Sub

Private Sub dtc_codigo2A_Click(Area As Integer)
    dtc_desc2A.BoundText = dtc_codigo2A.BoundText
    dtc_benef2A.BoundText = dtc_codigo2A.BoundText
    dtc_email2A.BoundText = dtc_codigo2A.BoundText
End Sub

Private Sub dtc_codigo4A_Click(Area As Integer)
    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    DataCombo2.Text = DataCombo1.BoundText
End Sub

Private Sub DataCombo2_Click(Area As Integer)
    DataCombo1.Text = DataCombo2.BoundText
End Sub

Private Sub cmdVerifica_existencia_Click()
' verifica existencia  del almacen
Cant_Alm = 0
AlFrmExistencia_Almacen.Show

DE.dbo_albSacaDetalleMaterial Mid(TxtCodigo, 3, 12), descri_bien, Cant_Alm
Txtcant_alm = Cant_Alm
If Cant_Alm >= TxtCantPedi Then
        optSi = True
    Else
        optNo = True
    End If
End Sub

Private Sub dtc_codigo11_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
End Sub

Private Sub dtc_desc11_LostFocus()
'    If dtc_codigo11.Text = "C" Or dtc_codigo11.Text = "V" Then
'        'TxtCobrado.Visible = False
'        'Label7.Visible = False
'        TxtConcepto.Text = Right(lbl_titulo.Caption + " -Edificio: " + RTrim(dtc_desc3.Text), Len(lbl_titulo.Caption + " -Edificio: " + RTrim(dtc_desc3.Text)) - 6)
''        TxtPlazo.Visible = True
'    Else
'        If dtc_codigo11.Text = "E" Then
'            TxtConcepto.Text = lbl_titulo.Caption + " - " + RTrim(dtc_desc3.Text)
'            TxtPlazo.Text = 0
''            TxtPlazo.Visible = False
'        Else
'        'dtc_codigo2.Text = "VD"
'        'dtc_desc2.Text = "VENTA DIRECTA"
'        'TxtCobrado.Visible = True
'        'Label7.Visible = True
'            TxtConcepto.Text = lbl_titulo.Caption + " - " + RTrim(dtc_desc3.Text)   '"VENTA DIRECTA AL CLIENTE"
'            TxtPlazo.Text = 0
''            TxtPlazo.Visible = False
'        End If
'    End If
End Sub

Private Sub dtccodmanejo_Click(Area As Integer)
    DtCCodigo.BoundText = dtccodmanejo.BoundText
    DtCDescripcion.BoundText = dtccodmanejo.BoundText
    dtcunidadmedida.BoundText = dtccodmanejo.BoundText
    dtccodpeso.BoundText = dtccodmanejo.BoundText
End Sub

Private Sub dtccodpeso_Click(Area As Integer)
    DtCCodigo.BoundText = dtccodpeso.BoundText
    DtCDescripcion.BoundText = dtccodpeso.BoundText
    dtcunidadmedida.BoundText = dtccodpeso.BoundText
    dtccodmanejo.BoundText = dtccodpeso.BoundText
End Sub

Private Sub dtccodpar_Click(Area As Integer)
    dtcdescripar.Text = dtccodpar.BoundText
End Sub

Private Sub dtccodpoa_Click(Area As Integer)
    dtcdespoa.Text = dtccodpoa.BoundText
End Sub

Private Sub dtccodpuesto_Click(Area As Integer)
    dtcdenopuesto.Text = dtccodpuesto.BoundText
End Sub

Private Sub dtccodtipoid_Click(Area As Integer)
    dtcdescrtipoid.BoundText = dtccodtipoid.BoundText
End Sub

Private Sub dtccoduni_Click(Area As Integer)
    dtcdescripuni.Text = dtccoduni.BoundText
End Sub

Private Sub dtccorrcompromiso_Click(Area As Integer)
    dtcfechacompromiso.BoundText = dtccorrcompromiso.BoundText
End Sub

Private Sub dtccorrsol_Click(Area As Integer)
 dtcfechasol.BoundText = dtccorrsol.BoundText
End Sub

Private Sub dtcdenominacionruc_Click(Area As Integer)
    dtcnroruc.BoundText = dtcdenominacionruc.BoundText
End Sub

Private Sub dtcdenopuesto_Click(Area As Integer)
    dtccodpuesto.Text = dtcdenopuesto.BoundText
End Sub

Private Sub DtCDescripcion_Click(Area As Integer)
    DtCCodigo.BoundText = DtCDescripcion.BoundText
    dtcunidadmedida.BoundText = DtCDescripcion.BoundText
    dtccodmanejo.BoundText = DtCDescripcion.BoundText
    dtccodpeso.BoundText = DtCDescripcion.BoundText
End Sub

Private Sub dtc_email2A_Click(Area As Integer)
    dtc_codigo2A.BoundText = dtc_email2A.BoundText
    dtc_benef2A.BoundText = dtc_email2A.BoundText
    dtc_desc2A.BoundText = dtc_email2A.BoundText
End Sub

Private Sub dtcdescripar_Click(Area As Integer)
    dtccodpar.Text = dtcdescripar.BoundText
End Sub

Private Sub dtcdescripuni_Click(Area As Integer)
    dtccoduni.Text = dtcdescripuni.BoundText
End Sub

Private Sub dtcdescrtipoid_Click(Area As Integer)
    dtccodtipoid.BoundText = dtcdescrtipoid.BoundText
End Sub

Private Sub dtcfechacompromiso_Click(Area As Integer)
    dtccorrcompromiso.BoundText = dtcfechacompromiso.BoundText
End Sub

Private Sub dtcfechasol_Click(Area As Integer)
    dtccorrsol.BoundText = dtcfechasol.BoundText
End Sub

Private Sub dtcnroruc_Click(Area As Integer)
    dtcdenominacionruc.Text = dtcnroruc.BoundText
End Sub

Private Sub dtc_desc2_LostFocus()
    'If AdoBeneficiario.Recordset!beneficiario_deudor = "SI" Then
    If Dtc_deudor2.Text = "SI" Then
        Dtc_deudor2.backColor = &HFF&
    Else
        Dtc_deudor2.backColor = &H80000010
    End If

End Sub

Private Sub dtc_desc4A_Click(Area As Integer)
    dtc_codigo4A.BoundText = dtc_desc4A.BoundText
End Sub

Private Sub dtctipodoc_Click(Area As Integer)
    dtcdenodoc.Text = dtctipodoc.BoundText
End Sub

Private Sub dtcunidadmedida_Click(Area As Integer)
    DtCCodigo.BoundText = dtcunidadmedida.BoundText
    DtCDescripcion.BoundText = dtcunidadmedida.BoundText
    dtccodmanejo.BoundText = dtcunidadmedida.BoundText
    dtccodpeso.BoundText = dtcunidadmedida.BoundText
End Sub

Private Sub dtcdespoa_Click(Area As Integer)
    dtccodpoa.Text = dtcdespoa.BoundText
End Sub

Private Sub dtc_desc2A_Click(Area As Integer)
    dtc_codigo2A.BoundText = dtc_desc2A.BoundText
    dtc_benef2A.BoundText = dtc_desc2A.BoundText
    dtc_email2A.BoundText = dtc_desc2A.BoundText
End Sub

'Private Sub DTPfechasol_Change()
'    txtGes_gestion = CStr(Year(DTPfechasol.Value))
'End Sub

'Private Sub DTPfechasol_LostFocus()
'    Set rs_TipoCambio = New ADODB.Recordset
'    If rs_TipoCambio.State = 1 Then rs_TipoCambio.Close
'    rs_TipoCambio.Open "select * from gc_tipo_cambio WHERE Fecha_Cambio='" & DTPfechasol & "'  ", db, adOpenKeyset, adLockReadOnly
'    If rs_TipoCambio.RecordCount > 0 Then
'        txtTDC.Text = rs_TipoCambio!cambio_oficial_compra
'    End If
'    'Ado_datos4.Refresh
'
'    DTPfechaIni.Value = DTPfechasol.Value
''    'validar fecha solicitud OJO JQA 31/12/2014
''    Set rs_TipoCambio = New ADODB.Recordset
''    If rs_TipoCambio.State = 1 Then rs_TipoCambio.Close
''    rs_TipoCambio.Open "select * from gc_tipo_cambio WHERE Fecha_Cambio='" & DTPfechasol & "'  ", db, adOpenKeyset, adLockReadOnly
''    If rs_TipoCambio.RecordCount > 0 Then
''        txtTDC.Text = rs_TipoCambio!cambio_oficial_compra
''    End If
'End Sub

Private Sub Form_Load()
    buscados = 0
    swnuevo = 0
    VAR_SW = ""
    lbl_cerrado = ""
    VAR_EXPOR = "NN"
    'db.Execute "fp_saldos"
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        usuario2 = rs_aux3!beneficiario_codigo
        VAR_DA = rs_aux3!da_codigo
        VAR_DPTO = rs_aux3!depto_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.3"
        VAR_DPTO = "2"
    End If
    If Aux = "DNREP" Or Aux = "DNEME" Or Aux = "DNINS" Then
        lbl_cite = "Cite / O.S."
        VAR_OS = "SI"
    Else
        lbl_cite = "Cod.Admin."
        VAR_OS = "NO"
    End If
    VAR_UORIGEN = Aux
    If Aux = "DNMAN" Then
        Select Case VAR_DA
            Case "1.8"    'Cochabamba
                Aux = "DMANB"
                'VAR_DPTO = "3"
            Case "1.7"    'Santa Cruz
                Aux = "DMANS"
                'VAR_DPTO = "7"
            Case "1.4", "1.2", "1.3"    'La Paz - Tecnico
                Aux = "DNMAN"
                'VAR_DPTO = "2"
            Case "1.9"    ' Chuquisaca
                Aux = "DMANC"
                'VAR_DPTO = "1"
            Case "0"    ' TODO
                Aux = "DNMAN"
                'VAR_DPTO = "2"
         End Select
     End If
     If Aux = "DNREP" Then
        Select Case VAR_DA
            Case "1.8"    'Cochabamba
                Aux = "DREPB"
                VAR_ALMACEN = 20
            Case "1.7"    'Santa Cruz
                Aux = "DREPS"
                VAR_ALMACEN = 21
            Case "1.4", "1.2", "1.3"    'La Paz - Comercial, Tecnico
                Aux = "DNREP"
                VAR_ALMACEN = 9
            Case "1.9"    ' Chuquisaca
                Aux = "DREPC"
                VAR_ALMACEN = 9
            Case "0"    ' TODO
                Aux = "DNREP"
                VAR_ALMACEN = 9
         End Select
     End If
     If Aux = "DNINS" Then
        Select Case VAR_DA
            Case "1.8"    'Cochabamba
                Aux = "DINSB"
                VAR_ALMACEN = 20
            Case "1.7"    'Santa Cruz
                Aux = "DINSS"
                VAR_ALMACEN = 21
            Case "1.4", "1.2", "1.3"    'La Paz - Comercial, Tecnico
                Aux = "DNINS"
                VAR_ALMACEN = 9
            Case "1.9"    ' Chuquisaca
                Aux = "DINSC"
                VAR_ALMACEN = 9
            Case "0"    ' TODO
                Aux = "DNINS"
                VAR_ALMACEN = 9
         End Select
     End If
     'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    db.Execute "UPDATE ao_ventas_cabecera SET venta_monto_origen_bs = venta_monto_total_bs WHERE venta_monto_origen_bs ='0' OR venta_monto_origen_bs IS NULL "
    db.Execute "UPDATE ao_ventas_cabecera SET codigo_empresa = '2' WHERE venta_tipo = 'G' AND codigo_empresa IS NULL "
    db.Execute "UPDATE ao_ventas_cabecera SET codigo_empresa = '1' WHERE venta_tipo <> 'G' AND codigo_empresa IS NULL "
    
    db.Execute "update ao_ventas_cabecera set estado_cancelado = 'Y' Where estado_codigo = 'REG'  "
    db.Execute "update ao_ventas_cabecera set estado_cancelado = 'A' Where estado_codigo = 'ANL'  "
    db.Execute "update ao_ventas_cabecera set estado_cancelado = 'A' Where estado_codigo = 'ERR'  "
    parametro = Aux
    Call OptFilGral1_Click
    Call ABRIR_TABLAS_AUX
    'Call ABRIR_TABLA
    'Call ABRIR_TABLA_AUX2
    'Call ABRIR_TABLA_DET3
    'txt_codigo.Enabled = True
    mbDataChanged = False
    FrmCabecera.Enabled = False
    dg_datos.Enabled = True
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    GlNombFor = "F04"
    'LblUsuario.Caption = GlUsuario
    marca1 = 1
    deta2 = 0
    BtnImprimir2.Visible = False
'    BtnImprimir3.Visible = False

'    FrmEdita.Enabled = False
'    FrmCobros.Enabled = False
'    Cmd_Cliente.Visible = False
    swnuevo = 0
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
    
    Chk_plazo.Value = 0
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'UNIDAD EJECUTORA
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText

    'Beneficiario Personas Nat. y Juridicas - Responsable del Contrato
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from gc_beneficiario where (estado_codigo ='APR' ) order by beneficiario_denominacion", db, adOpenStatic   'and tipoben_codigo <20
    'rs_datos2.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    dtc_aux2.BoundText = dtc_codigo2.BoundText
    Dtc_deudor2.BoundText = dtc_codigo2.BoundText


    'Proyecto de Edificación
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from gc_edificaciones order by edif_denominacion", db, adOpenStatic
    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText

    'Beneficiario Funcionario - Supervisor Tecnico
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "select * from gc_beneficiario where tipoben_codigo = '1' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    'rs_datos4.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    dtc_aux4.BoundText = dtc_codigo4.BoundText

    'Beneficiario Funcionario - Cobrador por Contrato
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' order by beneficiario_denominacion", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText

    'Beneficiario Funcionario - Cobrador por Pago
    Set rs_datos4A = New ADODB.Recordset
    If rs_datos4A.State = 1 Then rs_datos4A.Close
    rs_datos4A.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
    Set ado_datos4A.Recordset = rs_datos4A
    dtc_desc4A.BoundText = dtc_codigo4A.BoundText

'    'Motivos del Proceso
'    Set rs_datos7 = New ADODB.Recordset
'    If rs_datos7.State = 1 Then rs_datos7.Close
'    rs_datos7.Open "Select * from  rc_motivo_proceso where motivo_tipo = '0' ", db, adOpenStatic
'    'Set Ado_datos7.Recordset = rs_datos7
''    dtc_desc7.BoundText = dtc_codigo7.BoundText
    
    'ZONAS PILOTO
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "select * from tc_zonas_piloto order by zpiloto_descripcion ", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText

    'EMPRESA
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_empresas order by codigo_empresa", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText

   ' Factura a nombre de.... BENEFICIARIO
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_beneficiario WHERE estado_codigo = 'APR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_codigo2A.BoundText = dtc_benef2A.BoundText
    dtc_desc2A.BoundText = dtc_benef2A.BoundText
    dtc_email2A.BoundText = dtc_benef2A.BoundText

    'ac_tipo_compra_venta
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "select * from ac_tipo_compra_venta where venta_tipo <> 'L' and venta_tipo <> 'V' and venta_tipo <> 'A' ", db, adOpenStatic
    rs_datos11.Open "select * from ac_tipo_compra_venta where subproceso_codigo = 'TEC-02' or subproceso_codigo = 'TEC-03'  ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText

    'Solo para Equipos (*)
    Set rs_datos15 = New ADODB.Recordset
    If rs_datos15.State = 1 Then rs_datos15.Close
    rs_datos15.Open "select * from ac_bienes where estado_codigo = 'APR' order by bien_descripcion", db, adOpenKeyset, adLockReadOnly
    Set ado_datos15.Recordset = rs_datos15
    ado_datos15.Refresh
   'wwwwwwwwwwwwwwwwwwww
    Set rs_datos17 = New ADODB.Recordset
    If rs_datos17.State = 1 Then rs_datos17.Close
    rs_datos17.Open "select * from ac_bienes_grupo", db, adOpenKeyset, adLockReadOnly
    Set ado_datos17.Recordset = rs_datos17
    ado_datos17.Refresh
'WWWWWWWWWWWWWWWWWWWWWWWWWWWW
End Sub

Private Sub grabar()
  'db.BeginTrans
    If swgrabar = 1 Then
'      'Ado_datos.Recordset("venta_codigo") = Ado_datos.Recordset.RecordCount
'      'rstdestino.AddNew
    End If
       'Ado_datos.Recordset("ges_gestion") = glGestion       'CStr(Year(DTPfechasol.Value))
       'Ado_datos.Recordset("unidad_codigo") = dtc_codigo1.Text
       'Ado_datos.Recordset("solicitud_codigo") = txt_codigo.Caption
       'Ado_datos.Recordset("edif_codigo") = dtc_codigo3.Text
       'Ado_datos.Recordset("depto_codigo") = Left(dtc_codigo3.Text, 1)
       'Ado_datos.Recordset("venta_fecha") = IIf(DTPFechaIni.Value = "", Format(Date, "dd,mm,yyyy"), DTPFechaIni.Value)
       'Ado_datos.Recordset("venta_fecha_inicio") = IIf(DTPFechaIni.Value = "", Format(Date, "dd,mm,yyyy"), DTPFechaIni.Value)
       'Ado_datos.Recordset("venta_fecha_fin") = DTPFechaFin.Value
       'Ado_datos.Recordset("venta_tipo") = dtc_codigo11.Text                'E=Efectivo, C=Credito
       'Ado_datos.Recordset("beneficiario_codigo") = dtc_codigo2.Text        'CLIENTE
       'Ado_datos.Recordset("beneficiario_codigo_resp") = dtc_codigo4.Text   'Vendedor
       'Ado_datos.Recordset("beneficiario_codigo_cobr") = dtc_codigo5.Text   'Cobrador
       'Ado_datos.Recordset("venta_descripcion") = Trim(TxtConcepto.Text)
       'CONT2 = 365 / 30 * Ado_datos.Recordset!venta_cantidad_total
       'Ado_datos.Recordset("venta_plazo_dias_calendario") = IIf(TxtPlazo.Text = "", CONT2, TxtPlazo.Text)
        'GlTipoCambioOficial As Currency        'GlTipoCambioMercado As Currency        'GlTipoCambioGestion As Currency
       'Ado_datos.Recordset("venta_tipo_cambio") = GlTipoCambioMercado        'Val(txtTDC.Text)venta_tipo_cambio
       'Ado_datos.Recordset("tipoben_codigo") = IIf(Dtc_aux2.Text = "", "2", Dtc_aux2.Text)      'Tipo de Beneficiario
       
'       Ado_datos.Recordset("unimed_codigo_cobr") = cmd_unimed2.Text
'       Ado_datos.Recordset("venta_cantidad_cobr") = txtCantCobr.Text
'       Ado_datos.Recordset("mes_inicio_crono") = RTrim(cmb_mes_ini.Text)
'       VAR_MED2 = Ado_datos.Recordset!unimed_codigo_cobr
'       VAR_COBR2 = Ado_datos.Recordset!venta_cantidad_cobr
'       MControl = Ado_datos.Recordset!mes_inicio_crono
'       Ado_datos.Recordset("proceso_codigo") = "TEC"
'       Ado_datos.Recordset("subproceso_codigo") = "TEC-02"
'       Ado_datos.Recordset("etapa_codigo") = "TEC-02-02"
'       Ado_datos.Recordset("clasif_codigo") = "TEC"
'       Ado_datos.Recordset("doc_codigo") = "R-123"
'       'Ado_datos.Recordset("doc_numero") = "0"
'       Ado_datos.Recordset("poa_codigo") = "3.2.3"
     Select Case dtc_codigo1.Text             'dtc_codigo2.Text
        Case "DNINS", "DNAJS", "DINSS", "DINSB", "DINSC"    'VENTA DE SERVICIOS INSTTALACIONES
            VAR_CITE = Txt_campo2.Text
            VAR_CODDOC = "R-362"
            VAR_PRO = "COM"
            VAR_SUB = "COM-03"
            VAR_ETAPA = "COM-03-01"
            VAR_CLASIF = "ADM"
            'rs_datos!proceso_codigo = "COM"
            'rs_datos!subproceso_codigo = "COM-03"
            'rs_datos!etapa_codigo = "COM-03-01"
            'rs_datos!clasif_codigo = "TEC"
            'rs_datos!doc_codigo = "R-362"
        Case "DNMAN", "DMANS", "DMANB", "DMANC"             '10. SERVICIO MANTENIMIENTO PREVENTIVO
            VAR_CITE = Trim(dtc_aux3.Text) + "-" + Trim(Year(Ado_datos.Recordset!venta_fecha_inicio))
            VAR_CODDOC = "R-362"
            VAR_PRO = "TEC"
            VAR_SUB = "TEC-02"
            VAR_ETAPA = "TEC-02-02"
            VAR_CLASIF = "TEC"
'            VAR_CITE = Trim(Mid(Ado_datos.Recordset!edif_codigo, 7, 5)) + "-" + Trim(Year(Ado_datos.Recordset!venta_fecha_inicio))
'            rs_datos!unidad_codigo_ant = VAR_CITE
            'Ado_datos.Recordset("proceso_codigo") = "TEC"
            'Ado_datos.Recordset("subproceso_codigo") = "TEC-02"
            'Ado_datos.Recordset("etapa_codigo") = "TEC-02-02"
            'Ado_datos.Recordset("clasif_codigo") = "TEC"
            'Ado_datos.Recordset("doc_codigo") = "R-123"
            'Ado_datos.Recordset("doc_numero") = "0"
            'Ado_datos.Recordset("poa_codigo") = "3.2.3"
            'VAR_CITE = Txt_campo2.Text
        Case "DNREP", "DREPS", "DREPB", "DREPC"             'REPARACIONES
            VAR_CITE = Txt_campo2.Text
            VAR_CODDOC = "R-306"
            VAR_PRO = "TEC"
            VAR_SUB = "TEC-03"
            VAR_ETAPA = "TEC-03-02"
            VAR_CLASIF = "TEC"
            If dtc_codigo11.Text = "R" Then
                VAR_TRANS = "5"
            Else
                VAR_TRANS = "20"
            End If
            'rs_datos!proceso_codigo = "TEC"
            'rs_datos!subproceso_codigo = "TEC-03"
            'rs_datos!etapa_codigo = "TEC-03-02"
            'rs_datos!clasif_codigo = "TEC"
            'rs_datos!doc_codigo = "R-306"
'                rs_datos!unidad_codigo_ant = Txt_campo2.Text
        Case "DNEME"                                        '10 EMERGENCIAS
            VAR_CITE = Txt_campo2.Text
            VAR_CODDOC = "R-306"
            VAR_PRO = "TEC"
            VAR_SUB = "TEC-04"
            VAR_ETAPA = "TEC-04-01"
            VAR_CLASIF = "TEC"
                'rs_datos!proceso_codigo = "TEC"
                'rs_datos!subproceso_codigo = "TEC-04"
                'rs_datos!etapa_codigo = "TEC-04-01"
                'rs_datos!clasif_codigo = "TEC"
                'rs_datos!doc_codigo = "R-306"
'                rs_datos!unidad_codigo_ant = Txt_campo2.Text
     End Select
        'VAR_CODDOC = rs_datos!doc_codigo
'       Ado_datos.Recordset("saldo_p_cobrar") = Ado_datos.Recordset("monto_total_bS") - Ado_datos.Recordset("deuda_cobrada")
       If Ado_datos.Recordset!estado_codigo = "REG" Then
            'Ado_datos.Recordset("estado_codigo") = "REG"
            VAR_ESTADO = "REG"
       Else
            VAR_ESTADO = Ado_datos.Recordset!estado_codigo
       End If
       'Ado_datos.Recordset("usr_codigo") = glusuario
       'Ado_datos.Recordset("fecha_registro") = Format(Date, "dd/mm/yyyy")
       'Ado_datos.Recordset("hora_registro") = Format(Time, "hh/mm/ss")
       'Ado_datos.Recordset("usuario_aprueba") = ""
        'Ado_datos.Recordset("fecha_aprueba") = ""
        'Ado_datos.Recordset.Update
        '
       If Ado_datos.Recordset!depto_codigo = Left(Ado_datos.Recordset!edif_codigo, 1) Then
            VAR_DPTO = Ado_datos.Recordset!depto_codigo
       Else
            VAR_DPTO = Left(Ado_datos.Recordset!edif_codigo, 1)
       End If
        'GRABA DATOS DEL CONTRATO DE VENTA
        VAR_MED2 = cmd_unimed2.Text                                      'Ado_datos.Recordset!unimed_codigo_cobr
        VAR_COBR2 = IIf(txtCantCobr.Text = "", 1, Val(txtCantCobr.Text)) 'Ado_datos.Recordset!venta_cantidad_cobr
        MControl = RTrim(cmb_mes_ini.Text)                               'Ado_datos.Recordset!mes_inicio_crono
        gestion0 = Trim(Str(Year(FInicio)))
        If Ado_datos.Recordset!ges_gestion <> gestion0 Then
            db.Execute "UPDATE ao_ventas_cabecera SET ges_gestion = '" & Year(FInicio) & "' WHERE venta_codigo = " & NumComp & " "
        End If
        db.Execute "update ao_ventas_cabecera set mes_codigo = '" & Month(FInicio) & "' WHERE venta_codigo = " & NumComp & " "
    
        db.Execute "UPDATE ao_ventas_cabecera SET venta_fecha = '" & FInicio & "', venta_fecha_inicio = '" & FInicio & "', venta_fecha_fin = '" & FFin & "', venta_tipo = '" & dtc_codigo11.Text & "', unidad_codigo_ant = '" & VAR_CITE & "',  " & _
        " beneficiario_codigo = '" & dtc_codigo2.Text & "', beneficiario_codigo_RESP = '" & usuario2 & "', beneficiario_codigo_cobr = '" & dtc_codigo5.Text & "', venta_descripcion = '" & Trim(TxtConcepto.Text) & "', edif_codigo_corto ='" & dtc_aux3.Text & "'  " & _
        " WHERE venta_codigo = " & NumComp & " "
    
        db.Execute "UPDATE ao_ventas_cabecera SET venta_tipo_cambio = " & GlTipoCambioMercado & ", tipoben_codigo = " & IIf(dtc_aux2.Text = "", 2, Val(dtc_aux2.Text)) & ", codigo_empresa = " & dtc_codigo8.Text & ", venta_tipo = '" & dtc_codigo11.Text & "', " & _
        " mes_inicio_crono = '" & MControl & "', venta_cantidad_cobr = " & VAR_COBR2 & ", unimed_codigo_cobr = '" & VAR_MED2 & "', venta_cantidad_total = " & VAR_COBR2 & ", doc_codigo = '" & VAR_CODDOC & "', depto_codigo = '" & VAR_DPTO & "'  " & _
        " WHERE venta_codigo = " & NumComp & " "
        
        VAR_BEN2 = IIf(dtc_codigo4.Text = "", "0", dtc_codigo4.Text)
        'GRABA DATOS DEL CONTRATO PARA CRONOGRAMA
        If Ado_datos.Recordset!unidad_codigo = "DNMAN" Or Ado_datos.Recordset!unidad_codigo = "DMANS" Or Ado_datos.Recordset!unidad_codigo = "DMANB" Or Ado_datos.Recordset!unidad_codigo = "DMANC" Then
            db.Execute "UPDATE ao_ventas_cabecera SET beneficiario_codigo_tec = '" & VAR_BEN2 & "', zpiloto_codigo= " & dtc_codigo7.Text & ", cantidad_periodos_tec= " & txt_cant.Text & ", unimed_codigo_tec = '" & cmd_unimed_tec.Text & "', " & _
            " fecha_inicio_tec = '" & lbl_fecha_ini.Value & "', fecha_fin_tec = '" & lbl_fecha_fin.Value & "', mes_inicio_crono_tec = '" & cmb_mes_ini_tec.Text & "', mes_par_impar = '" & dtc_aux7.Text & "' " & _
            " WHERE venta_codigo = " & NumComp & " "
        End If
        ' GRABA DATOS GENERALES
        'VAR_PRO, VAR_SUB, VAR_ETAPA, VAR_CLASIF, VAR_DOCS
        db.Execute "UPDATE ao_ventas_cabecera SET proceso_codigo = '" & VAR_PRO & "', subproceso_codigo = '" & VAR_SUB & "', etapa_codigo = '" & VAR_ETAPA & "', clasif_codigo = '" & VAR_CLASIF & "', poa_codigo = '3.2.3', " & _
        " estado_codigo = '" & VAR_ESTADO & "', usr_codigo = '" & glusuario & "', fecha_registro = '" & Format(Date, "dd/mm/yyyy") & "', hora_registro = '" & Format(Time, "hh/mm/ss") & "' " & _
        " WHERE venta_codigo = " & NumComp & " "
        
        'GENERA CORREL CONTRATO POR DEPTO INI
        Set rs_aux7 = New ADODB.Recordset
        If rs_aux7.State = 1 Then rs_aux7.Close
        rs_aux7.Open "Select correl_contrato as Codigo from gc_departamento where depto_codigo = '" & VAR_DPTO & "'    ", db, adOpenStatic        '
        If Not rs_aux7.EOF Then
            'VAR_CONTR = IIf(IsNull(rs_aux7!Codigo), 1, CDbl(rs_aux7!Codigo) + 1)
            If IsNull(rs_aux7!Codigo) Then
                VAR_CONTR = 1
            Else
                VAR_CONTR = IIf(IsNull(rs_aux7!Codigo), 1, CDbl(rs_aux7!Codigo) + 1)
            End If
        End If
        'db.Execute "update ao_ventas_cabecera set venta_codigo_new = " & VAR_CONTR & " Where ao_ventas_cabecera.venta_codigo = " & NumComp & "  And ao_ventas_cabecera.ges_gestion = " & gestion0 & " "
        db.Execute "update ao_ventas_cabecera set venta_codigo_new = " & VAR_CONTR & " Where ao_ventas_cabecera.venta_codigo = " & NumComp & "  And ao_ventas_cabecera.ges_gestion = " & gestion0 & " "
        db.Execute "update gc_departamento set correl_contrato = " & VAR_CONTR & " Where depto_codigo = '" & Left(VAR_PROY2, 1) & "' "
        'GENERA CORREL CONTRATO POR DEPTO FIN

    'Ado_datos.Recordset.Requery
    'If rstdestino.State = 1 Then rstdestino.Close
    'db.CommitTrans
    If Ado_datos.Recordset.RecordCount > 0 Then
       marca1 = Ado_datos.Recordset.Bookmark
       'Call OptFilGral1_Click
       'Ado_datos.Refresh
       'Ado_datos.Recordset.Move marca1 - 1
        If swgrabar = 1 Then
            Ado_datos.Refresh
            Ado_datos.Recordset.MoveLast
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'  If glPersNew = "P" Then
'    frmmo_formulario_M1.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre = rs_Personal!pers_nombres
'    frmmo_formulario_M1.Dtc_Pers_Cargo = rs_Personal!cargo_codigo
'  End If
'  If glPersNew = "L" Then
'    frmmo_formulario_M1.Dtc_doc_id_lab = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell_lab = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2apell_lab = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre_lab = rs_Personal!pers_nombres
'  End If
'  glPersNew = "N"
    db.Execute "fp_saldos"
End Sub

Private Sub opt_salir_Click()
    fra_reportes.Visible = False
End Sub

Private Sub opt_vigentes_eqp3_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Dim iResult As Variant, i%, Y%
        Dim co As New ADODB.Command
    '    Dim rs As New ADODB.Recordset
    '    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    '            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly
    '    i = 1
    '    y = 1
        Select Case Me.Ado_datos.Recordset!unidad_codigo
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
          Case "DVTA"
              var_titulo = "Módulo Comercial"
        End Select

        CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_lista_de_ventas_zpiloto_txt.rpt"
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True

        CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
        
        CryV01.Formulas(1) = "titulo = '" & var_titulo & "' "
        CryV01.Formulas(2) = "subtitulo = '" & lbl_titulo.Caption & "' "
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
        fra_reportes.Visible = False
    Else
        MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
    End If
End Sub

Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
  
    Select Case VAR_DPTO
        Case "2"
            Select Case VAR_UORIGEN
                Case "DNMAN"
                    If glusuario = "ADMIN" Or glusuario = "VBELLIDO" Or glusuario = "SQUISPE" Or glusuario = "CSALINAS" Then
                        queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo LIKE  '%MAN%')) "
                    Else
                        queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo LIKE  '%MAN%')) "
                    End If
                Case "DNREP"
                    If glusuario = "ADMIN" Or glusuario = "KBETANCOURTH" Or glusuario = "LNAVA" Or glusuario = "FFLORES" Or glusuario = "CARIZACA" Or glusuario = "SQUISPE" Or glusuario = "CSALINAS" Or glusuario = "ARODRIGUEZ" Or glusuario = "RGIL" Or glusuario = "LMORALES" Or glusuario = "GMORA" Then
                        queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo LIKE  '%REP%')) "
                    Else
                        queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo LIKE  '%REP%')) "
                    End If
                Case "DNINS"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNINS')) "
                Case "DNEME"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNEME')) "
                Case Else
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN')) "
            End Select
        Case "7"
            Select Case VAR_UORIGEN
                Case "DNMAN"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '8' OR  Left(edif_codigo, 1) = '9' OR Left(edif_codigo, 1) = '1' ) ) "
                Case "DNREP"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNREP') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '8' OR  Left(edif_codigo, 1) = '9' OR Left(edif_codigo, 1) = '1' ) ) "
                Case "DNINS"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNINS') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '8' OR  Left(edif_codigo, 1) = '9' OR Left(edif_codigo, 1) = '1' ) ) "
                Case "DNEME"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNEME') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '8' OR  Left(edif_codigo, 1) = '9' OR Left(edif_codigo, 1) = '1' ) ) "
                Case Else
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '8' OR  Left(edif_codigo, 1) = '9' OR Left(edif_codigo, 1) = '1' ) ) "
            End Select
        Case "3"
            Select Case VAR_UORIGEN
                Case "DNMAN"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4' ) ) "
                Case "DNREP"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNREP') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4'  ) ) "
                Case "DNINS"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNINS') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4' ) ) "
                Case "DNEME"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNEME') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4'  ) ) "
                Case Else
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4'  ) ) "
            End Select
        Case "1"
            Select Case VAR_UORIGEN
                Case "DNMAN"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
                Case "DNREP"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNREP') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
                Case "DNINS"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNINS') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
                Case "DNEME"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNEME') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
                Case Else
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
            End Select
        Case "4"
            Select Case VAR_UORIGEN
                Case "DNMAN"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4' ) ) "
                Case "DNREP"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNREP') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4'  ) ) "
                Case "DNINS"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNINS') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4' ) ) "
                Case "DNEME"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNEME') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4'  ) ) "
                Case Else
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4'  ) ) "
            End Select
        Case "6"
            Select Case VAR_UORIGEN
                Case "DNMAN"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '6' ) ) "
                Case "DNREP"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNREP') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '6'  ) ) "
                Case "DNINS"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNINS') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '6' ) ) "
                Case "DNEME"
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNEME') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '6'  ) ) "
                Case Else
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '6'  ) ) "
            End Select
        Case Else
            'queryinicial = "Select * from ao_solicitud where (estado_codigo = 'REG' AND Left(edif_codigo, 1) = '" & VAR_DPTOC & "' AND (unidad_codigo = '" & parametro & "' OR unidad_codigo = '" & VAR_UORIGEN & "')) "
            queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' or (unidad_codigo LIKE  '%MAN%'))) "
    End Select
  'End If
'    If VAR_DPTO = "2" Then
'        queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') "
'    Else
'        'queryinicial = "select * From av_ventas_cabecera WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG' AND unidad_codigo='" & VAR_UORIGEN & "' AND depto_codigo = '" & VAR_DPTO & "')) "
'        'queryinicial = "select * From av_ventas_cabecera WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG' AND unidad_codigo='" & VAR_UORIGEN & "' AND left(edif_codigo,1) = '" & VAR_DPTO & "')) "
'        'queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND left(edif_codigo,1) = '" & VAR_DPTO & "' AND (unidad_codigo = '" & parametro & "' OR unidad_codigo='" & VAR_UORIGEN & "' )) "
'        queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND left(edif_codigo,1) = '" & VAR_DPTO & "' AND (unidad_codigo = '" & parametro & "' OR unidad_codigo='" & VAR_UORIGEN & "' )) "
'    End If
'    'queryinicial = "select * From av_ventas_cabecera WHERE estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "SOLICITUD_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
    db.Execute "fp_saldos"
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
  'If glusuario = "ADMIN" Or glusuario = "VBELLIDO" Then
  '      queryinicial = "select * From av_ventas_cabecera WHERE (unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN') "
  'Else
    Select Case VAR_DPTO
        Case "2"
            Select Case VAR_UORIGEN
                Case "DNMAN"
                    If glusuario = "ADMIN" Or glusuario = "VBELLIDO" Or glusuario = "SQUISPE" Or glusuario = "CSALINAS" Then
                        'queryinicial = "select * From av_ventas_cabecera WHERE (unidad_codigo = 'DMANS' or unidad_codigo = 'DNMAN' or unidad_codigo = 'DMANB' or unidad_codigo = 'DMANC') "
                        queryinicial = "select * From av_ventas_cabecera WHERE (unidad_codigo LIKE  '%MAN%') "
                    Else
                        queryinicial = "select * From av_ventas_cabecera WHERE (unidad_codigo LIKE  '%MAN%') "
                    End If
                Case "DNREP"
                    If glusuario = "ADMIN" Or glusuario = "KBETANCOURTH" Or glusuario = "LNAVA" Or glusuario = "FFLORES" Or glusuario = "CARIZACA" Or glusuario = "SQUISPE" Or glusuario = "CSALINAS" Or glusuario = "ARODRIGUEZ" Or glusuario = "RGIL" Or glusuario = "LMORALES" Or glusuario = "GMORA" Or glusuario = "ASANTIVAÑEZ" Then
                        queryinicial = "select * From av_ventas_cabecera WHERE (unidad_codigo LIKE  '%REP%') "
                    Else
                        queryinicial = "select * From av_ventas_cabecera WHERE (unidad_codigo LIKE  '%REP%') "
                    End If
                Case "DNINS"
                    queryinicial = "select * From av_ventas_cabecera WHERE (unidad_codigo LIKE  '%INS%') "
                Case "DNEME"
                    queryinicial = "select * From av_ventas_cabecera WHERE (unidad_codigo LIKE  '%EME%') "
                Case Else
                    queryinicial = "select * From av_ventas_cabecera WHERE (unidad_codigo LIKE  '%MAN%') "
            End Select
        Case "7"
            Select Case VAR_UORIGEN
                Case "DNMAN"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%MAN%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '8' OR  Left(edif_codigo, 1) = '9' OR Left(edif_codigo, 1) = '1' ) ) "
                Case "DNREP"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%REP%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '8' OR  Left(edif_codigo, 1) = '9' OR Left(edif_codigo, 1) = '1' ) ) "
                Case "DNINS"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%INS%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '8' OR  Left(edif_codigo, 1) = '9' OR Left(edif_codigo, 1) = '1' ) ) "
                Case "DNEME"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%EME%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '8' OR  Left(edif_codigo, 1) = '9' OR Left(edif_codigo, 1) = '1' ) ) "
                Case Else
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%MAN%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '8' OR  Left(edif_codigo, 1) = '9' OR Left(edif_codigo, 1) = '1' ) ) "
            End Select
        Case "3"
            Select Case VAR_UORIGEN
                Case "DNMAN"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%MAN%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4' ) ) "
                Case "DNREP"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%REP%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4'  ) ) "
                Case "DNINS"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%INS%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4' ) ) "
                Case "DNEME"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%EME%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4'  ) ) "
                Case Else
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%MAN%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4'  ) ) "
            End Select
            'queryinicial = "Select * from ao_solicitud where (estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "' AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR  Left(edif_codigo, 1) = '4' )) "
        Case "1"
            Select Case VAR_UORIGEN
                Case "DNMAN"
                    'queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE '%MAN%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
                Case "DNREP"
                    'queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNREP') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE '%REP%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
                Case "DNINS"
                    'queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNINS') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE '%INS%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
                Case "DNEME"
                    'queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNEME') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE '%EME%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
                Case Else
                    'queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "' or unidad_codigo = 'DNMAN') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE '%MAN%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ) ) "
            End Select
            'queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "' or unidad_codigo = 'DMMAN') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR  Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' ))"
        Case "4"
            Select Case VAR_UORIGEN
                Case "DNMAN"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%MAN%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4' ) ) "
                Case "DNREP"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%REP%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4'  ) ) "
                Case "DNINS"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%INS%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4' ) ) "
                Case "DNEME"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%EME%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4'  ) ) "
                Case Else
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%MAN%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '4'  ) ) "
            End Select
        Case "6"
            Select Case VAR_UORIGEN
                Case "DNMAN"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%MAN%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '6' ) ) "
                Case "DNREP"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%REP%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '6'  ) ) "
                Case "DNINS"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%INS%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '6' ) ) "
                Case "DNEME"
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%EME%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '6'  ) ) "
                Case Else
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo LIKE  '%MAN%') AND (Left(edif_codigo, 1) = '" & VAR_DPTO & "' OR   Left(edif_codigo, 1) = '6'  ) ) "
            End Select
        Case Else
            'queryinicial = "Select * from ao_solicitud where (estado_codigo = 'REG' AND Left(edif_codigo, 1) = '" & VAR_DPTOC & "' AND (unidad_codigo = '" & parametro & "' OR unidad_codigo = '" & VAR_UORIGEN & "')) "
            queryinicial = "select * From av_ventas_cabecera WHERE (unidad_codigo = '" & parametro & "' or (unidad_codigo LIKE '%MAN%')) "
    End Select
  'End If
    'queryinicial = "select * From av_ventas_cabecera WHERE (unidad_codigo = '" & parametro & "') "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "SOLICITUD_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
    db.Execute "fp_saldos"
End Sub

Private Sub TxtCantPedi_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtcaracteristicas_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtMonto_bolivianos_contra_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub TxtMonto_bolivianos_contra_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If (Len(Trim(TxtMonto_bolivianos_contra.Text)) > 0) Then
       Txtmonto_dolares_contra.Text = IIf(TxtMonto_bolivianos_contra.Text > 0, TxtMonto_bolivianos_contra.Text / TxtTipo_cambio, 0)
    Else
       Txtmonto_dolares_contra.Text = 0
    End If
  End If
End Sub

Private Sub TxtMonto_bolivianos_KeyPress(KeyAscii As Integer)
'solo numeros y , .
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
    Else
      KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub txtjustifica_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtMonto_bolivianos_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If (Len(Trim(TxtMonto_bolivianos.Text)) > 0) Then
       Txtmonto_dolares.Text = IIf(TxtMonto_bolivianos.Text > 0, TxtMonto_bolivianos.Text / TxtTipo_cambio, 0)
    Else
       Txtmonto_dolares.Text = 0
    End If
  End If

End Sub

Private Sub Txtmonto_dolares_contra_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtmonto_dolares_contra_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If Len(Trim(Txtmonto_dolares_contra.Text)) > 0 Then
      TxtMonto_bolivianos_contra.Text = IIf(Txtmonto_dolares_contra.Text > 0, Txtmonto_dolares_contra * TxtTipo_cambio, 0)
    Else
      TxtMonto_bolivianos_contra.Text = 0
    End If
  End If
End Sub

Private Sub Txtmonto_dolares_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtmonto_dolares_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If Len(Trim(Txtmonto_dolares.Text)) > 0 Then
      TxtMonto_bolivianos.Text = IIf(Txtmonto_dolares.Text > 0, Txtmonto_dolares * TxtTipo_cambio, 0)
    Else
      TxtMonto_bolivianos.Text = 0
    End If
  End If
End Sub

Private Sub Txtobservaciones_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtsolpeso_KeyPress(KeyAscii As Integer)
'solo numeros y , .
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then

    Else
      KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub txtterref_KeyPress(KeyAscii As Integer)
    If KeyAscii < 58 And KeyAscii > 47 Then
        KeyAscii = Asc(UCase(Chr(0)))
    Else
        If UCase(Chr(KeyAscii)) = "S" Or UCase(Chr(KeyAscii)) = "N" Or KeyAscii = 8 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Else
            KeyAscii = Asc(UCase(Chr(0)))
            MsgBox "Debe escribir solo 'N' o 'S'", vbOKOnly, "Error..."
        End If
    End If
End Sub

Private Sub cerea()
  txt_venta = " "
  dtc_codigo4.Text = " "
  Dtcpaternosol.Text = " "  'dtc_codigo4.BoundText
'  dtcmaternosol.Text = " "
'  dtcnombresol.Text = " "
  txtCantTotal = "0"
  TxtMontoBs = "0"
  TxtMontoUs = "0"
  TxtConcepto = ""
  dtc_codigo2 = ""
  dtc_desc2 = ""
  txtTDC.Text = GlTipoCambioOficial

'  DtCDenominacion_moneda = ""
'  TxtMonto_bolivianos = 0
'  Txtmonto_dolares = 0
'  TxtMonto_bolivianos_contra = 0
'  Txtmonto_dolares_contra = 0
'  DtCOrg_descripcion = ""
'  txtjustifica = ""
'  txt_venta = ""
'  txtterref = ""
End Sub

Private Sub acumulaMont(ges, Nro)
  Set rstacumdet = New ADODB.Recordset
  If rstacumdet.State = 1 Then rstacumdet.Close
  Set rs_datos19 = New ADODB.Recordset
  If rs_datos19.State = 1 Then rs_datos19.Close
'  LblGestion
'  lblcorrelVenta
'  lblNroVenta
  'rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as VAR_COBR2 from ao_ventas_detalle where ges_gestion = '" & ges & "' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
  If (Ado_datos.Recordset!unidad_codigo = "DNMAN" Or Ado_datos.Recordset!unidad_codigo = "DMANS" Or Ado_datos.Recordset!unidad_codigo = "DMANB" Or Ado_datos.Recordset!unidad_codigo = "DMANC") Then
    rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot0 from ao_ventas_detalle where venta_codigo = " & Nro & " and par_codigo = '43340'", db, adOpenKeyset, adLockOptimistic
  Else
    rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot0 from ao_ventas_detalle where venta_codigo = " & Nro & " and par_codigo <> '43340'", db, adOpenKeyset, adLockOptimistic
  End If
  If IsNull(rstacumdet!totbs) Then
    VAR_AUX = 0
    VAR_AUX2 = 0
    VAR_CANT = 1
  Else
    VAR_AUX = Round(rstacumdet!totbs, 2)
    VAR_AUX2 = Round(rstacumdet!totdl, 2)
    VAR_CANT = rstacumdet!cantot0
  End If

  'rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza_prog where ges_gestion = '" & ges & "' and estado_codigo = 'APR' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
  rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza where estado_codigo = 'APR' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  'rs_datos19.Open "select cobranza_bs as totbs2, cobranza_dol as totdl2 from FV_VENTAS_COBRANZA_TESORERIA where venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  If IsNull(rs_datos19!totbs2) Then
    Cobrobs = 0
    VAR_COBR = 0
  Else
    Cobrobs = Round(rs_datos19!totbs2, 2)
    VAR_COBR = Round(rs_datos19!totdl2, 2)
  End If

  VAR_Bs = VAR_AUX - Cobrobs
  VAR_Dol = VAR_AUX2 - VAR_COBR
  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & VAR_AUX & " , ao_ventas_cabecera.venta_monto_total_dol = " & VAR_AUX2 & ", ao_ventas_cabecera.venta_cantidad_total = " & VAR_CANT & ", ao_ventas_cabecera.venta_monto_cobrado_bs = " & Cobrobs & ", ao_ventas_cabecera.venta_monto_cobrado_dol = " & VAR_COBR & ",  ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_Bs & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_Dol & "  Where ao_ventas_cabecera.venta_codigo = " & Nro & " "

  TxtMontoBs.Text = VAR_AUX
  TxtCobrado.Text = Cobrobs
  TxtBstotal.Text = VAR_Bs

  If rstacumdet.State = 1 Then rstacumdet.Close

End Sub

Private Sub opt_vigentes_eqp2_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Dim iResult As Variant, i%, Y%
        Dim co As New ADODB.Command

    '    Dim rs As New ADODB.Recordset
    '    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    '            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly
    '    i = 1
    '    y = 1
        Select Case Me.Ado_datos.Recordset!unidad_codigo
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN", "DMANS", "DMANB", "DMANC"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP", "DREPS", "DREPB", "DREPC"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
          Case "DVTA"
              var_titulo = "Módulo Comercial"
        End Select

        CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_lista_de_ventas_zpiloto.rpt"
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True
        'CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        'CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
        'CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
        CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
        
        CryV01.Formulas(1) = "titulo = '" & var_titulo & "' "
        CryV01.Formulas(2) = "subtitulo = '" & lbl_titulo.Caption & "' "
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
        fra_reportes.Visible = False
    Else
        MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
    End If
End Sub

Private Sub opt_vigentes_eqp1_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Dim iResult As Variant, i%, Y%
        Dim co As New ADODB.Command

    '    Dim rs As New ADODB.Recordset
    '    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    '            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly
    '    i = 1
    '    y = 1
        Select Case Me.Ado_datos.Recordset!unidad_codigo
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
          Case "DVTA"
              var_titulo = "Módulo Comercial"
        End Select

        CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_lista_de_ventas.rpt"
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True
        'CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        'CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
        'CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
        CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
        
        CryV01.Formulas(1) = "titulo = '" & var_titulo & "' "
        CryV01.Formulas(2) = "subtitulo = '" & lbl_titulo.Caption & "' "
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
        fra_reportes.Visible = False
    Else
        MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
    End If

End Sub


Private Sub opt_todos_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Dim iResult As Variant, i%, Y%
        Dim co As New ADODB.Command

    '    Dim rs As New ADODB.Recordset
    '    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    '            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly

        Select Case Me.Ado_datos.Recordset!unidad_codigo
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN", "DMANS", "DMANB", "DMANC"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP", "DREPS", "DREPB", "DREPC"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
          Case "DVTA"
              var_titulo = "Módulo Comercial"
        End Select

        CryV02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_de_ventas_estado_todo.rpt"
        CryV02.WindowShowPrintSetupBtn = True
        CryV02.WindowShowRefreshBtn = True
        'CryV02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        'CryV02.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
        'CryV02.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
        CryV02.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
        
        CryV02.Formulas(1) = "titulo = '" & var_titulo & "' "
        CryV02.Formulas(2) = "subtitulo = '" & lbl_titulo.Caption & "' "
        iResult = CryV02.PrintReport
        If iResult <> 0 Then MsgBox CryV02.LastErrorNumber & " : " & CryV02.LastErrorString, vbCritical, "Error de impresión"
        fra_reportes.Visible = False
    Else
        MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
    End If

End Sub

Private Sub opt_vigentes_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Dim iResult As Variant, i%, Y%
        Dim co As New ADODB.Command

    '    Dim rs As New ADODB.Recordset
    '    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    '            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly

        Select Case Me.Ado_datos.Recordset!unidad_codigo
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
          Case "DVTA"
              var_titulo = "Módulo Comercial"
        End Select

        CryV02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_de_ventas_estado.rpt"
        CryV02.WindowShowPrintSetupBtn = True
        CryV02.WindowShowRefreshBtn = True
        'CryV02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        'CryV02.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
        'CryV02.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
        CryV02.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
        
        CryV02.Formulas(1) = "titulo = '" & var_titulo & "' "
        CryV02.Formulas(2) = "subtitulo = '" & lbl_titulo.Caption & "' "
        iResult = CryV02.PrintReport
        If iResult <> 0 Then MsgBox CryV02.LastErrorNumber & " : " & CryV02.LastErrorString, vbCritical, "Error de impresión"
        fra_reportes.Visible = False
    Else
        MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
    End If
End Sub

Private Sub Option1_Click()
    If Ado_datos.Recordset!estado_codigo = "APR" Then       'And Ado_datos.Recordset!estado_cancelado = "S"
        Dim iResult As Variant, i%, Y%
        Dim co As New ADODB.Command
        CryR02.ReportFileName = App.Path & "\reportes\Tecnico\tr_certificado_cumplim_contrato.rpt"
        'CryR02.WindowShowRefreshBtn = True
        CryR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CryR02.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
    
        iResult = CryR02.PrintReport
        If iResult <> 0 Then MsgBox CryR02.LastErrorNumber & " : " & CryR02.LastErrorString, vbCritical, "Error de impresión"
     Else
        MsgBox "El trámite debe estar CERRADO, para emitir el Certificado de Cumplimiento de Contrato, revise y vuelva a intentar !! ", vbExclamation, "Atención!"
     End If
     fra_reportes.Visible = False
End Sub

Private Sub Option10_Click()
    'Programar Meses IMPARES y quitar PARES
    VAR_IMPAR = "1"
    Option11.Value = False
    Option10.Value = True
    LblParImpar = "MESES IMPARES"
End Sub

Private Sub Option11_Click()
    'PROGRAMAR en Meses PARES y quitar Mes IMPARES
    VAR_IMPAR = "2"
    Option11.Value = True
    Option10.Value = False
    LblParImpar = "MESES PARES"
End Sub

Private Sub Option2_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Dim iResult As Variant, i%, Y%
        Dim co As New ADODB.Command
    '    Dim rs As New ADODB.Recordset
    '    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    '            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly
    '    i = 1
    '    y = 1
        Select Case Me.Ado_datos.Recordset!unidad_codigo
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
          Case "DVTA"
              var_titulo = "Módulo Comercial"
        End Select

        CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_lista_de_ventas_zpiloto_txt.rpt"
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True

        CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
        
        CryV01.Formulas(1) = "titulo = '" & var_titulo & "' "
        CryV01.Formulas(2) = "subtitulo = '" & lbl_titulo.Caption & "' "
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
        fra_reportes.Visible = False
    Else
        MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
    End If
End Sub

Private Sub sstab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
        Case 0
        Case 1
            Select Case Ado_datos.Recordset!mes_par_impar
                Case 0
                    'NO ASIGNADO
                    VAR_IMPAR = "0"
                    Option11.Value = False
                    Option10.Value = False
                    LblParImpar = "NO ASIGNADO"
                Case 1
                    'Programar Meses IMPARES y quitar PARES
                    VAR_IMPAR = "1"
                    Option11.Value = False
                    Option10.Value = True
                    LblParImpar = "MESES IMPARES"
                Case 2
                    'PROGRAMAR en Meses PARES y quitar Mes IMPARES
                    VAR_IMPAR = "2"
                    Option11.Value = True
                    Option10.Value = False
                    LblParImpar = "MESES PARES"
                Case Else
                    'NO ASIGNADO
                    VAR_IMPAR = "0"
                    Option11.Value = False
                    Option10.Value = False
                    LblParImpar = "NO ASIGNADO"
            End Select
            If txt_cant.Text = "" Or txt_cant.Text = "0" Then       'Nro.de Preriodos
                txt_cant.Text = txtCantCobr.Text
            End If
            If cmd_unimed_tec.Text = "" Or cmd_unimed_tec.Text = "0" Then      'Periodicidad
                cmd_unimed_tec.Text = cmd_unimed2.Text
            End If
            lbl_fecha_ini.Value = IIf(IsNull(lbl_fecha_ini.Value), DTPfechaIni.Value, lbl_fecha_ini.Value)       'Fecha Inicio Crono.)
            lbl_fecha_fin.Value = IIf(IsNull(lbl_fecha_fin.Value), DTPfechaFin.Value, lbl_fecha_fin.Value)       'Fecha Fin Crono.)
            
            If cmb_mes_ini_tec.Text = "" Or cmb_mes_ini_tec.Text = "0" Then                                   'Mes de Inicio Crono.
                cmb_mes_ini_tec.Text = cmb_mes_ini.Text
            End If
            'tc_zona_piloto_edif
            If IsNull(Ado_datos.Recordset!zpiloto_codigo) Or Ado_datos.Recordset!zpiloto_codigo = "0" Then
                Set rs_datos10 = New ADODB.Recordset
                If rs_datos10.State = 1 Then rs_datos10.Close
                rs_datos10.Open "Select * from tc_zona_piloto_edif where edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "' ", db, adOpenStatic
                If rs_datos10.RecordCount > 0 Then
                    dtc_codigo7.Text = rs_datos10!zpiloto_codigo
                    dtc_desc7.BoundText = dtc_codigo7.BoundText
                    dtc_aux7.BoundText = dtc_codigo7.BoundText
                'Else
                    
                End If
                'Set Ado_datos1.Recordset = rs_datos1
                'dtc_desc1.BoundText = dtc_codigo1.BoundText
            Else
                
            End If
            dtc_desc4.BoundText = dtc_codigo4.BoundText
            dtc_aux4.BoundText = dtc_codigo4.BoundText
        Case 2
        Case 3
        Case Else
    End Select
End Sub

Private Sub TxtCobrado_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Then      '(KeyAscii = 8) Or '(0..9)
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub TxtDsctoTot_LostFocus()
    If TxtDsctoTot.Text = "" Or TxtDsctoTot.Text = "0" Or TxtDsctoTot.Text = "0.00" Then
        TxtMonto.Text = "0"
    Else
        TxtMonto.Text = Round(CDbl(TxtDsctoTot.Text) * GlTipoCambioMercado, 2)
    End If
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
  '? . , 09
  ',.01234856789
End Sub

Private Sub TxtMonto_LostFocus()
    If TxtMonto.Text = "" Or TxtMonto.Text = "0" Or TxtMonto.Text = "0.00" Then
        TxtDsctoTot.Text = "0"
    Else
        TxtDsctoTot.Text = Round(CDbl(TxtMonto.Text) / GlTipoCambioMercado, 2)
    End If
End Sub

Private Sub TxtPlazo_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]" Or KeyAscii = 8, KeyAscii, 0)
End Sub

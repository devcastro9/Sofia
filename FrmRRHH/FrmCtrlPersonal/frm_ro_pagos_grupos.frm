VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_ro_pagos_grupos 
   Caption         =   "RRHH - Planillas"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   13155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFC0&
      Height          =   6255
      Left            =   960
      TabIndex        =   17
      Top             =   1920
      Width           =   3975
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   5880
         Width           =   3825
         _ExtentX        =   6747
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
         BackColor       =   16777152
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
         Caption         =   " <-- Inicio                        Gerencia General                          Fin -->"
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
      Begin TrueOleDBGrid60.TDBGrid grdPrincipal 
         Height          =   5520
         Left            =   120
         OleObjectBlob   =   "frm_ro_pagos_grupos.frx":0000
         TabIndex        =   31
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame fraDatosProponente 
      Caption         =   "Liquidación "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   4920
      TabIndex        =   30
      Top             =   1680
      Width           =   8175
      Begin VB.CommandButton cmdImprimeLiquida 
         Caption         =   "&Imprimir"
         Height          =   600
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Imprime orden de liquidación y cronograma..."
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdPagoAprob 
         Caption         =   "A&probar"
         Height          =   600
         Left            =   3600
         TabIndex        =   46
         ToolTipText     =   "Aprobar orden de liquidación"
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdPagoDesaprob 
         Caption         =   "Desaprobar"
         Height          =   600
         Left            =   3600
         TabIndex        =   54
         ToolTipText     =   "Desaprueba la orden de liquidación"
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdPagoDeven 
         Caption         =   "Genera &Devengado"
         Height          =   600
         Left            =   4560
         TabIndex        =   48
         ToolTipText     =   "Genera el devengado de la liquidación seleccionada"
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdPagoEditar 
         Caption         =   "&Modificar"
         Height          =   600
         Left            =   1605
         TabIndex        =   47
         ToolTipText     =   "Modifica datos de liquidación..."
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdPagoAnular 
         Caption         =   "&Eliminar"
         Height          =   600
         Left            =   2610
         TabIndex        =   45
         ToolTipText     =   "Elimina la liquidación seleccionada"
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdPagoNuevo 
         Caption         =   "&Adicionar"
         Height          =   600
         Left            =   600
         TabIndex        =   44
         ToolTipText     =   "Permite adicionar liquidaciones..."
         Top             =   240
         Width           =   1005
      End
      Begin Crystal.CrystalReport CRCrono 
         Left            =   120
         Top             =   2880
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   1
         WindowControlBox=   -1  'True
         WindowMaxButton =   0   'False
         WindowMinButton =   0   'False
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCancelBtn=   0   'False
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.CommandButton cmdPagoAnulaDev 
         Caption         =   "Anula D&evengado"
         Height          =   600
         Left            =   5520
         TabIndex        =   56
         ToolTipText     =   "Anula la aprobación de orden de liquidación y su devengado."
         Top             =   240
         Width           =   1005
      End
      Begin TrueOleDBGrid60.TDBGrid grdLiquida 
         Height          =   2400
         Left            =   120
         OleObjectBlob   =   "frm_ro_pagos_grupos.frx":2979
         TabIndex        =   60
         Top             =   840
         Width           =   7935
      End
      Begin VB.Label Label23 
         Caption         =   "Totales:"
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
         Left            =   3480
         TabIndex        =   35
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblTotalUSLiq 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$US:"
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
         Left            =   4320
         TabIndex        =   34
         Top             =   3360
         Width           =   1830
      End
      Begin VB.Label lblTotalBSLiq 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bs.:"
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
         Left            =   6240
         TabIndex        =   33
         Top             =   3360
         Width           =   1830
      End
      Begin VB.Label lblEstadoLiquida 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblEstadoLiquida"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   3360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblNroLiquida 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro. registros:"
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
         Left            =   120
         TabIndex        =   36
         Top             =   3360
         Width           =   3135
      End
   End
   Begin VB.Frame fraFiltra 
      Height          =   1215
      Left            =   1080
      TabIndex        =   8
      Top             =   600
      Width           =   3735
      Begin VB.CommandButton cmdFiltro 
         Height          =   480
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Filtra datos sobre grupos de liquidación."
         Top             =   720
         Width           =   525
      End
      Begin VB.CommandButton cmdOrdAZ 
         Height          =   480
         Left            =   765
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ordena en forma ascendente."
         Top             =   720
         Width           =   525
      End
      Begin VB.CommandButton cmdOrdZA 
         Height          =   480
         Left            =   1290
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ordena en forma descendente."
         Top             =   720
         Width           =   525
      End
      Begin VB.CommandButton cmdBusca 
         Height          =   480
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Busca por la columna seleccionada."
         Top             =   720
         Width           =   525
      End
      Begin VB.CommandButton cmdActualiza 
         Height          =   480
         Left            =   2325
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Actualiza los datos de grupos de liquidación."
         Top             =   720
         Width           =   525
      End
      Begin VB.CommandButton cmdImprime 
         Height          =   480
         Left            =   2865
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Imprime el grid de grupos de liquidación."
         Top             =   720
         Width           =   525
      End
      Begin MSDataListLib.DataCombo cboUnidadSol 
         Height          =   315
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Unidad Sol.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraToolBarMain 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7155
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   1020
      Begin VB.CommandButton cmdRelCronoCompro 
         Caption         =   "Crono. vs Compro."
         Height          =   735
         Left            =   120
         TabIndex        =   59
         Top             =   2520
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Sale del formulario"
         Top             =   6240
         Width           =   765
      End
      Begin VB.CommandButton cmdAnular 
         Caption         =   "A&nular"
         Height          =   720
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Anula el grupo de liquidación, su comprometido y desaprueba la adjudicación."
         Top             =   1680
         Width           =   765
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Adicionar"
         Height          =   720
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "No habilitado en este modulo..."
         Top             =   240
         Width           =   765
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "M&odificar"
         Height          =   720
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Modifica datos del grupo de liquidación..."
         Top             =   960
         Width           =   765
      End
   End
   Begin VB.Frame fraBeneficiario 
      Caption         =   "Beneficiario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   4920
      TabIndex        =   37
      Top             =   5400
      Width           =   8175
      Begin VB.CommandButton cmdBenAnularTodo 
         Caption         =   "Eliminar Todo"
         Height          =   600
         Left            =   5040
         TabIndex        =   58
         ToolTipText     =   "Elimina todos los beneficiario de la liquidación"
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdDatosContrato 
         Caption         =   "Datos &Contrato"
         Height          =   600
         Left            =   240
         TabIndex        =   55
         ToolTipText     =   "Permite confirmar datos de contrato..."
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdBenConf 
         Caption         =   "Reg. Con&formidad"
         Height          =   600
         Left            =   6000
         TabIndex        =   53
         ToolTipText     =   "Registra conformidad de producto..."
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdBenMonto 
         Caption         =   "Captura Mon&tos"
         Height          =   600
         Left            =   2925
         TabIndex        =   52
         ToolTipText     =   "Captura montos de pago..."
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdBenNuevo 
         Caption         =   "Adicionar &Beneficiario"
         Height          =   600
         Left            =   1920
         TabIndex        =   51
         ToolTipText     =   "Adiciona beneficiario al pago..."
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdBenAnular 
         Caption         =   "Eliminar Se&lecionado"
         Height          =   600
         Left            =   3930
         TabIndex        =   50
         ToolTipText     =   "Elimina beneficiario seleccionado"
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton cmdBenFact 
         Caption         =   "Fact&ura?"
         Height          =   600
         Left            =   7005
         TabIndex        =   49
         ToolTipText     =   "Registra si el beneficiario emite factura..."
         Top             =   240
         Width           =   1005
      End
      Begin TrueOleDBGrid60.TDBGrid grdBeneficiario 
         Height          =   2280
         Left            =   120
         OleObjectBlob   =   "frm_ro_pagos_grupos.frx":52F4
         TabIndex        =   61
         Top             =   960
         Width           =   7935
      End
      Begin VB.Label lblEstadoBeneficiario 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblEstadoBeneficiario"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1560
         TabIndex        =   43
         Top             =   3240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblNroBeneficiario 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro. registros:"
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
         Left            =   120
         TabIndex        =   42
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label Label16 
         Caption         =   "Totales:"
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
         Left            =   3360
         TabIndex        =   41
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblTotalUS 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$US:"
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
         Left            =   4200
         TabIndex        =   40
         Top             =   3240
         Width           =   1830
      End
      Begin VB.Label lblTotalBS 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bs.:"
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
         Left            =   6120
         TabIndex        =   39
         Top             =   3240
         Width           =   1830
      End
   End
   Begin VB.Label lbl_titulo2 
      Alignment       =   2  'Center
      Caption         =   "PROCESO DE LIQUIDACIÓN DE CONSULTOR INDIVIDUAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   11895
   End
   Begin VB.Label lblEstadoGrupoLiq 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblEstadoGrupoLiq"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2280
      TabIndex        =   29
      Top             =   8400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblCodUniSol 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblCodGrupo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   27
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblDesGrupo 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   26
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Grupo Liquidación:"
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
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "Form.:"
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
      Height          =   255
      Left            =   6960
      TabIndex        =   24
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblUsuario 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuario: XXXXX"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7680
      TabIndex        =   2
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label lblFormulario 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   23
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label lblGestion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   22
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label30 
      Caption         =   "Gestión:"
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
      Height          =   255
      Left            =   4920
      TabIndex        =   21
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lbl_titulo 
      Alignment       =   2  'Center
      Caption         =   "PROCESO DE LIQUIDACIÓN DE CONSULTOR INDIVIDUAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   11895
   End
   Begin VB.Label lblDA 
      Caption         =   "D.A.: XXXXX"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label lblDesUniSol 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   20
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Label Label28 
      Caption         =   "Uni. Sol.:"
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
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblNroPrincipal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro. registros:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   18
      Top             =   8160
      Width           =   3735
   End
End
Attribute VB_Name = "frm_ro_pagos_grupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQLs As String ' usado para la elaboración de los querys
Dim filtro As String ' usado para la elaboración de los querys
Dim nro_reg As Integer
Dim rs_grdPrincipal As ADODB.Recordset ' usado para navegar sobre el grid principal
Dim rs_grdLiquida As ADODB.Recordset ' usado para navegar sobre el grid
Dim rs_grdBeneficiario As ADODB.Recordset ' usado para navegar sobre el grid

Private Sub cboUnidadSol_Change()
    cboUnidadSol.ToolTipText = cboUnidadSol.Text
    
    If cboUnidadSol.BoundText <> "" Then
        rs_grdPrincipal.Filter = "codigo_unidad like " & Chr(39) & cboUnidadSol.BoundText & Chr(39)
        filtro = "Unidad -> " & Chr(39) & cboUnidadSol.BoundText & Chr(39) ' concatenamos la cadena de filtración
        lblNroPrincipal.Caption = "Nro. de liq.: " & rs_grdPrincipal.RecordCount & " Filtro ( " & filtro & " )"
        Call pl_PersonalizaGridPrincipal
        
    End If
    
End Sub

Private Sub cboUnidadSol_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
      Case 46 ' si presiono suprimir
        cboUnidadSol.BoundText = ""
        Call cmdActualiza_Click ' se refresca el grid principal
      Case 13 ' si presiono enter
        SendKeys "{Tab}"
    End Select

End Sub

Private Sub cmdActualiza_Click()
    Dim Gestion As String
    Dim CodUni  As String
    Dim CodGrupo  As Integer
    
    Screen.MousePointer = vbHourglass
    
    cboUnidadSol.BoundText = ""
    If rs_grdPrincipal.RecordCount > 0 Then
        ' se guarda los datos de registro para poder ubicar luego el registro
        Gestion = rs_grdPrincipal!ges_gestion
        CodUni = rs_grdPrincipal!codigo_unidad
        CodGrupo = rs_grdPrincipal!codigo_grupo
    End If
    Call pl_RefrescaListaPrincipal ' se refresca el recordset para mostrar los datos originales
    
    If Len(Gestion) > 0 Then ' si se tiene un registro activo
        rs_grdPrincipal.Find " ges_gestion ='" & Gestion & "'" ' el puntero de registro se ubica en la posicion guardada
        rs_grdPrincipal.Find " codigo_unidad ='" & CodUni & "'" 'el puntero de registro se ubica en la posicion guardada
        rs_grdPrincipal.Find " codigo_grupo =" & CodGrupo ' el puntero de registro se ubica en la posicion guardada
    End If
    Call grdPrincipal_RowColChange(0, 0)
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdAnulaDev_Click()
    Call pl_OpcionesGenericas("Tool_AnulaOrdLiq", "Liquida")
End Sub

Private Sub cmdAnular_Click()
    Call pl_OpcionesGenericas("Tool_Anular", "GrupoLiq")
End Sub

Private Sub cmdBenAnularTodo_Click()
    Call pl_OpcionesGenericas("Tool_AnularTodo", "Beneficiario")
End Sub

Private Sub cmdBenConf_Click()
    Call pl_OpcionesGenericas("Tool_Conformidad", "Beneficiario")
End Sub

Private Sub cmdBenAnular_Click()
    Call pl_OpcionesGenericas("Tool_Anular", "Beneficiario")
End Sub

Private Sub cmdBenFact_Click()
    Call pl_OpcionesGenericas("Tool_Factura", "Beneficiario")
End Sub

Private Sub cmdBenMonto_Click()
    Call pl_OpcionesGenericas("Tool_Monto", "Beneficiario")
End Sub

Private Sub cmdBenNuevo_Click()
    Call pl_OpcionesGenericas("Tool_Nuevo", "Beneficiario")
End Sub

Private Sub cmdBenNuevoSegun_Click()
    Call pl_OpcionesGenericas("Tool_NuevoSegun", "Beneficiario")
End Sub

Private Sub cmdBusca_Click()
    If rs_grdPrincipal.RecordCount > 0 Then
        Call pg_BuscaTdbGrid(grdPrincipal, rs_grdPrincipal, grdPrincipal.Columns(grdPrincipal.Col).DataField)
        Call pl_PersonalizaGridPrincipal
      Else
        MsgBox "No existen registros para búscar.", vbInformation, "Aviso"
    End If

End Sub

Private Sub cmdDatosContrato_Click()
    Call pl_OpcionesGenericas("Tool_Contrato", "Beneficiario")
End Sub

Private Sub cmdEditar_Click()
    Call pl_OpcionesGenericas("Tool_Editar", "GrupoLiq")
End Sub

Private Sub cmdFiltro_Click()
    'PROPÓSITO      : Realiza la filtración de una especificación sobre la columna de la celda activa
    
    Dim CadFiltro As String ' usada para almacenar la cedena de filtración
    Dim micriterio As String ' critetio de filtración
    Dim CampoAct As Integer ' nombre del campo activo
    Dim ColFiltro As String ' usada para almacenar la columna activa por el cual se filtrará
    Dim a As Integer ' usada para ver el formato de cadena a filtrar
    
    If rs_grdPrincipal.RecordCount > 0 Then
    
        micriterio = "Digite " & LCase(grdPrincipal.Columns(grdPrincipal.Col).Caption) & " a filtrar"
        CadFiltro = pg_QuitaEspBlanco(UCase(InputBox(micriterio, "Filtración")))
        ' verificamos que la cadena sea del tipo *a* donde a representa cualquier secuencia de caracteres
        Select Case Len(CadFiltro)
          Case Is >= 3
            If Left(CadFiltro, 1) = "*" And Right(CadFiltro, 1) <> "*" Then
                ' completamos la cadena al tipo *a*
                CadFiltro = CadFiltro & "*"
              Else
                ' es del tipo a* o a que son cadenas validas
                'CadFiltro = "*" & CadFiltro
            End If
          Case 2
            If Left(CadFiltro, 1) = "*" And Right(CadFiltro, 1) <> "*" Then
                CadFiltro = CadFiltro & "*"
              Else
                If Left(CadFiltro, 1) = "*" And Right(CadFiltro, 1) = "*" Then
                    ' si ambos son *
                    CadFiltro = ""
                  Else
                    ' es del tipo a* o a que son cadenas validas
                    'CadFiltro = "*" & CadFiltro
                End If
            End If
          Case 1
            If CadFiltro = "*" Then CadFiltro = ""
        End Select
        
        On Error GoTo EtiqError
        
        If Len(CadFiltro) > 0 Then ' si introdujo una cadena a filtrar
            CampoAct = grdPrincipal.Col
            ColFiltro = grdPrincipal.Columns(grdPrincipal.Col).DataField
            ' verificamos si la longitud coincide con el tamaño del campo
            If Len(CadFiltro) <= rs_grdPrincipal.Fields(ColFiltro).DefinedSize Or rs_grdPrincipal.Fields(ColFiltro).Type = 3 Then
                If rs_grdPrincipal.Filter = 0 Then ' es la primera filtración
                    rs_grdPrincipal.Filter = ColFiltro & " like " & Chr(39) & CadFiltro & Chr(39)
                    filtro = grdPrincipal.Columns(CampoAct).Caption & " -> " & Chr(39) & CadFiltro & Chr(39) ' concatenamos la cadena de filtración
                  Else
                    rs_grdPrincipal.Filter = rs_grdPrincipal.Filter & " AND " & ColFiltro & " like " & Chr(39) & CadFiltro & Chr(39)
                    filtro = filtro & ", " & grdPrincipal.Columns(CampoAct).Caption & " -> " & Chr(39) & CadFiltro & Chr(39)  ' concatenamos la cadena de filtración
                End If
                lblNroPrincipal.Caption = "Nro. de Liq: " & rs_grdPrincipal.RecordCount & " Filtro ( " & filtro & " )"
                
                If rs_grdPrincipal.RecordCount = 0 Then ' no se encontraron coincidencias
                    Call grdPrincipal_RowColChange(0, 0)
                    MsgBox "No se encontró ninguna coincidencia con " & filtro, vbInformation, "Información"

                End If
                grdPrincipal.SetFocus
                
              Else ' la longitud de la cadena a filtrar es mayor a la longitud del campo
                MsgBox "La longitud de la cadena a filtrar -> " & CadFiltro & " es mayor a la longitud del permitido por " & grdPrincipal.Columns(CampoAct).Caption, vbInformation, "Información"
            End If
            grdPrincipal.SetFocus
          Else ' solo tiene el foco
            grdPrincipal.SetFocus
        End If
      Else
        MsgBox "No existen registros para ser filtrados.", vbInformation, "Aviso"
    End If
    
    Call pl_PersonalizaGridPrincipal

    On Error GoTo 0 ' desactiva el manejador de errores
    Exit Sub
    
EtiqError:
    Select Case Err.Number
      Case -2147352571
        MsgBox "Error: No se pueden filtrar los datos, los tipos no coinciden." & Chr(13) & Chr(13) & "No se realizo la filtración de datos." & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description, vbCritical, "Error"
      Case Else ' si se produjo otro tipo de error
        MsgBox "Error: No se realizo la filtración de datos." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
    End Select

End Sub

Private Sub cmdImprime_Click()
    
    If rs_grdPrincipal.RecordCount > 0 Then
        Call pg_Imprimir(grdPrincipal, grdPrincipal.Caption)
      Else
        MsgBox "No existen registros para ser impresos.", vbInformation, "Aviso"
    End If

End Sub

Private Sub cmdImprimeLiquida_Click()
    If rs_grdPrincipal.RecordCount > 0 Then
        If rs_grdLiquida.RecordCount > 0 Then
            ''** llama al formulario
            lblEstadoBeneficiario.Tag = rs_grdLiquida!NUMERO_PAGO & ""
            frm_ro_PagosPrintOrdenPago.Show vbModal
          Else
            MsgBox "No existe resgistros para liquidación." & Chr(13) & "Corrija el error e intente imprimir nuevamente.", vbInformation, "Aviso"
        End If
      Else
        MsgBox "No existen registros para ser impresos.", vbInformation, "Aviso"
    End If

End Sub

Private Sub cmdNuevo_Click()
    Call pl_OpcionesGenericas("Tool_Nuevo", "GrupoLiq")
End Sub

Private Sub cmdOrdAZ_Click()
    If rs_grdPrincipal.RecordCount > 0 Then
        Call pg_OrdenaTdbGrid(grdPrincipal, rs_grdPrincipal, True)
        Call pl_PersonalizaGridPrincipal
      Else
        MsgBox "No existen registros para ordenar.", vbInformation, "Aviso"
    End If

End Sub

Private Sub cmdOrdZA_Click()
    If rs_grdPrincipal.RecordCount > 0 Then
        Call pg_OrdenaTdbGrid(grdPrincipal, rs_grdPrincipal, False)
        Call pl_PersonalizaGridPrincipal
      Else
        MsgBox "No existen registros para ordenar.", vbInformation, "Aviso"
    End If

End Sub

Private Sub cmdPagoAnulaDev_Click()
    Call pl_OpcionesGenericas("Tool_AnulaDevengar", "Liquida")
End Sub

Private Sub cmdPagoanular_Click()
    Call pl_OpcionesGenericas("Tool_Anular", "Liquida")
End Sub

Private Sub cmdPagoAprob_Click()
    Call pl_OpcionesGenericas("Tool_Aprobar", "Liquida")
End Sub

Private Sub cmdPagoDesaprob_Click()
    Call pl_OpcionesGenericas("Tool_Desaprobar", "Liquida")
End Sub

Private Sub cmdPagoDeven_Click()
    Call pl_OpcionesGenericas("Tool_Devengar", "Liquida")
End Sub

Private Sub cmdPagoEditar_Click()
    Call pl_OpcionesGenericas("Tool_Editar", "Liquida")
End Sub

Private Sub cmdPagoNuevo_Click()
    Call pl_OpcionesGenericas("Tool_Nuevo", "Liquida")
End Sub

Private Sub cmdRelCronoCompro_Click()
'    If rs_grdPrincipal.RecordCount > 0 Then
'        If Not (rs_grdPrincipal.EOF Or rs_grdPrincipal.BOF) Then
'            frm_ro_HistRelCronoCompro.xGes_Gestion = rs_grdPrincipal!ges_gestion
'            frm_ro_HistRelCronoCompro.xCodigo_Unidad = rs_grdPrincipal!codigo_unidad
'            frm_ro_HistRelCronoCompro.xCodigo_Grupo = rs_grdPrincipal!codigo_grupo
'        End If
'    End If
'    If rs_grdBeneficiario.RecordCount > 0 Then
'        If Not (rs_grdBeneficiario.EOF Or rs_grdBeneficiario.BOF) Then
'            frm_ro_HistRelCronoCompro.xNumero_Pago = rs_grdBeneficiario!NUMERO_PAGO
'            'frm_ro_HistRelCronoCompro.Xcodigo_beneficiario = rs_grdBeneficiario!codigo_beneficiario
'        End If
'    End If
'    frm_ro_HistRelCronoCompro.Show vbModal

End Sub

Private Sub cmdSalir_Click()
    Call pl_OpcionesGenericas("Tool_Salir", "GrupoLiq")
End Sub

Private Sub Form_Load()
    
    ' obtiene direccion administrativa en funcion del usuario
    Set rstTemp = New ADODB.Recordset
    SQLs = "SELECT gc_usuarios.DA, fc_direccion_administrativa.descripcion_DA "
    SQLs = SQLs & "FROM gc_usuarios INNER JOIN fc_direccion_administrativa ON gc_usuarios.DA = fc_direccion_administrativa.DA "
    SQLs = SQLs & "WHERE gc_usuarios.usr_usuario = '" & GlUsuario & "' AND gc_usuarios.Usr_Activo = 1 "
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        GldaCodigo = rstTemp!da & ""
        GldaDescrip = rstTemp!descripcion_DA & ""
      Else
        GldaCodigo = ""
        GldaDescrip = ""
        MsgBox "Error: No existe relación entre el usuario y una Dirección Administrativa." & Chr(13) & "Esto puede causar muchos errores." & Chr(13) & "Anote el error y comuniquese con el admisnitrador del sistema.", vbError, "Aviso"
    End If
    
    Call pl_Llena_Combos_Base 'llena los combos base
    
    Call pl_RefrescaListaPrincipal 'refresca la lista principal
    
    Call pl_ValoresDefecto
    
    Call grdPrincipal_RowColChange(0, 0)

'''/***
''DE.Edson.Open
''DE.Edson.Execute "SET DATEFORMAT dmy"
'''**/

	Call SeguridadSet(Me)
End Sub

Private Sub pl_Llena_Combos_Base()
    ' llena los combos y listas base para la carga del formulario
    
    Dim rstTemp As ADODB.Recordset ' usado para la carga de los combos de base
    
    ' unidad solicitante - para realizar filtro sobre el grid principal
    Set rstTemp = New ADODB.Recordset
    SQLs = "select codigo_unidad, codigo_unidad + ' - '+ uni_descripcion_larga as des_unidad from fc_unidad_ejecutora where uni_activo = 'S' ORDER BY codigo_unidad"
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        Set cboUnidadSol.RowSource = rstTemp
        cboUnidadSol.BoundColumn = "codigo_unidad"
        cboUnidadSol.ListField = "des_unidad"
        
      Else
        MsgBox "El catalogo de unidad solicitante no esta actualizado.", vbInformation, "Aviso"
    End If
    
    Set rstTemp = Nothing

End Sub

Private Sub pl_ValoresDefecto()
    'PROPOSITO:             Permite establecer los valores por defecto de los elementos del formulario
    
    lblUsuario.Caption = "Usuario: " & GlUsuario ' usuario
    lblDA.Caption = "Dir. Adm.: " & GldaDescrip ' descripcion direc. adm.
    
    Select Case GldaCodigo
      Case "01" ' si da es DGAARRYHH
        
        Me.Caption = "SAF - Proceso de Liquidación de Consultor"
        lbl_titulo.Caption = "PROCESO DE LIQUIDACIÓN CONSULTOR"
      
      Case "52" ' si da es DAP
        
        Me.Caption = "SAF - Proceso de Liquidación de Consultor"
        lbl_titulo.Caption = "PROCESO DE LIQUIDACIÓN CONSULTOR"
      
      Case "00" ' es para recursos humanos RRYHH
        
        Me.Caption = "SAF - Proceso de Liquidación de Consultor RRyHH"
        lbl_titulo.Caption = "PROCESO DE LIQUIDACION CONSULTOR A LARGO PLAZO"
    
    End Select
    
End Sub

Private Sub pl_RefrescaListaPrincipal()
  ' Onjetivo: Procedimiento refrescar o actualizar la lista principal de solicitudes
  
    On Error GoTo 0 ' activamos el manejador de errores
  
    Screen.MousePointer = vbHourglass
    cboUnidadSol.BoundText = ""
    
    SQLs = "SELECT ao_pagos_grupos.codigo_unidad, ao_pagos_grupos.codigo_solicitud, ao_pagos_grupos.codigo_grupo, ao_pagos_grupos.descripcion_grupo, "
    SQLs = SQLs & "'ModPago' = case when ao_pagos_grupos.modalidad_pago = 'P' then 'Planilla' else 'Invividual' end, ao_pagos_grupos.estado_aprobado, fc_unidad_ejecutora.Uni_descripcion_larga, ac_tipo_tramite.denominacion_tipo, ao_pagos_grupos.ges_gestion, ao_pagos_grupos.modalidad_pago, "
    SQLs = SQLs & "ao_pagos_grupos.formulario , ao_pagos_grupos.da, ao_pagos_grupos.numero_consultoria, ao_pagos_grupos.correl_grupo_da "

    SQLs = SQLs & "FROM ac_tipo_tramite INNER JOIN ao_pagos_grupos ON ac_tipo_tramite.tipo_formulario = ao_pagos_grupos.formulario LEFT OUTER JOIN fc_unidad_ejecutora ON "
    SQLs = SQLs & "ao_pagos_grupos.codigo_unidad = fc_unidad_ejecutora.codigo_unidad "
    

    Select Case glProceso
      Case "F05"
        SQLs = SQLs & "WHERE ao_pagos_grupos.estado_aprobado <> 'E' AND ao_pagos_grupos.formulario = 'F05' and ao_pagos_grupos.da = '" & GldaCodigo & "'"
      Case "F10"
        SQLs = SQLs & "WHERE ao_pagos_grupos.estado_aprobado <> 'E' AND ao_pagos_grupos.formulario = 'F10' "
    End Select
    
    SQLs = SQLs & "ORDER BY ao_pagos_grupos.ges_gestion, ao_pagos_grupos.codigo_unidad, ao_pagos_grupos.codigo_grupo"

    Set rs_grdPrincipal = New ADODB.Recordset
    rs_grdPrincipal.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    Set grdPrincipal.DataSource = rs_grdPrincipal
    Call pl_PersonalizaGridPrincipal
    lblNroPrincipal.Caption = "Nro. grupo de liquidación: " & rs_grdPrincipal.RecordCount
    
    Screen.MousePointer = vbDefault
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub pl_PersonalizaGridPrincipal()
    'TITULO:                Procedimiento pl_PersonalizaGridPrincipal
    'PROPOSITO:             Personalizar los captions, anchos, etc. del grid
    'EJEMPLO DE LLAMADA:    call pl_PersonalizaGridPrincipal
        
    Dim i As Integer
    
    ' define ancho de columnas y titulo de la cabecera
    grdPrincipal.Columns(0).Width = 750 ' codigo unidad
    grdPrincipal.Columns(0).Caption = "Unidad"
    grdPrincipal.Columns(1).Width = 700 ' codigo solicitud
    grdPrincipal.Columns(1).Caption = "Solicitud Original"
    grdPrincipal.Columns(2).Width = 600 ' cod grupo
    grdPrincipal.Columns(2).Caption = "Código Grupo"
    grdPrincipal.Columns(3).Width = 1800 ' des grupo
    grdPrincipal.Columns(3).Caption = "Descripción Grupo"
    grdPrincipal.Columns(4).Width = 900 ' Modalidad pago
    grdPrincipal.Columns(4).Caption = "Modalidad Pago"
    grdPrincipal.Columns(5).Width = 800 ' estado aprobado
    grdPrincipal.Columns(5).Caption = "Est. Aprobado"
    
    For i = 6 To rs_grdPrincipal.Fields.Count - 1
        grdPrincipal.Columns(i).Visible = False
        grdPrincipal.Columns(i).AllowSizing = False
    Next i
    
    
End Sub

Private Sub pl_OpcionesGenericas(TipoOpcion As String, Proceso As String)
    'TITULO:                Procedimiento pl_OpcionesGenericas
    'PROPOSITO:             Ejecuta una opcion del toolbar
    'EJEMPLO DE LLAMADA:    call pl_OpcionesGenericas(TipoOpcion)
    'ENTRADAS:              TipoOpcion = Opción a elegir (Grabar,Editar, etc.)
                            ' Realiza una acción según TipoOpcion

    Dim Cad As String '
    Dim swGuardar As Integer ' usado para saber si efectivamente se almaceno o elimino los datos en la base
                          ' swGuarda -> 0 si se realizo el proceso satisfactoriamente
                          ' swGuarda -> 1 si se produjo un evento de cancelar por parte del usuario en el proceso
                          ' swGuarda -> 2 si se produjo un error de integridad de la base de datos en el servidor por el proceso
    Dim RegPuntero As Long ' usada para guardar el código de registro para poder apuntar el el registro seleccionado luego de un refresh
    Dim fechax As String
    Dim horax  As String
    Dim i As Integer
    
    On Error GoTo EtiqError
    
    Select Case Proceso
      
      ' ********************************************************
      ' opciones genericas de la ficha: GRUPOS
      ' ********************************************************
      
      Case "GrupoLiq" ' procesa la ficha GRUPOS
        Select Case TipoOpcion
          Case "Tool_Nuevo"
            ' no se procesa en este modulo
            
          Case "Tool_Editar"
            
            If rs_grdPrincipal.RecordCount > 0 Then
                
                ' ***********************
                ' se llama al formulario
                ' ***********************
                Screen.MousePointer = vbHourglass
                lblEstadoGrupoLiq.Caption = "E" ' se esta en modo de edicion del reg. actual
                frm_ro_LiquidaAdiGrupo.Show vbModal
                
                If lblEstadoGrupoLiq.Caption <> "E" Then  ' si el proceso la modificacion
                    Cad = lblCodUniSol.Caption   ' codigo unidad
                    RegPuntero = Val(lblCodGrupo.Caption) ' codigo grupo
                    Call pl_RefrescaListaPrincipal
                    rs_grdPrincipal.Find "codigo_unidad = '" & Cad & "'" ' se posisicona en el registro editado
                    rs_grdPrincipal.Find "codigo_grupo = " & RegPuntero ' se posisicona en el registro editado
                    Call grdPrincipal_RowColChange(0, 0)
                    MsgBox "Los cambios se realizaron satisfactoriamente", vbInformation, "Aviso"
                    grdPrincipal.SetFocus
                End If
                lblEstadoGrupoLiq.Caption = "" ' no se esta editando ni adicionando registros
              Else
                MsgBox "No existen registro de grupos de liquidación para ser editado.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
            
          Case "Tool_Anular"
            
            If rs_grdPrincipal.RecordCount > 0 Then ' si existen registros
                
                If fl_VerificaEliminaGrupo Then
                    Screen.MousePointer = vbHourglass
                    Cad = "Se eliminará el grupo de liquidación, se anulará el comprometido,se desaprueba la adjudicación y registro de contrato:" & Chr(13)
                    Cad = Cad & "Grupo:[" & lblCodGrupo.Caption & " - " & lblDesGrupo.Caption & "]" & Chr(13) & "Unidad: [" & lblCodUniSol.Caption & " - " & lblDesUniSol.Caption & "]." & Chr(13) & "Desea continuar no podrá revertir el proceso?"
                    If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de elimnación") Then
                        'JQ QR
                        'DE.dbo_rp_PagosBorraGrupo rs_grdPrincipal!ges_gestion, rs_grdPrincipal!codigo_unidad, rs_grdPrincipal!codigo_grupo, rs_grdPrincipal!numero_consultoria
                        Call pl_RefrescaListaPrincipal
                        Call grdPrincipal_RowColChange(0, 0)
                        MsgBox "La información se eliminó satisfactoriamente", vbInformation, "Aviso"
                    End If
                  Else
                    grdPrincipal.SetFocus
                    
                End If
                Screen.MousePointer = vbDefault
              Else
                MsgBox "No existen registros para ser eliminados.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
                    
          Case "Tool_Salir"
            Unload Me
            
        End Select
    
      ' ********************************************************
      ' opciones genericas de la: LIQUIDACION
      ' ********************************************************
      
      Case "Liquida" ' procesa la ficha LIQUIDACION
        
        Select Case TipoOpcion
          Case "Tool_Nuevo"
            ' ***********************
            ' se llama al formulario
            ' ***********************
            Screen.MousePointer = vbHourglass
            lblEstadoLiquida.Caption = "N" ' se esta en modo de adicion de nuevo registro
            frm_ro_LiquidaAdiPago.Show vbModal
            
            If lblEstadoLiquida.Caption <> "N" Then ' si el proceso realizo la adicion de un registro
                RegPuntero = Val(lblEstadoLiquida.Caption)  ' codigo
                Call grdPrincipal_RowColChange(0, 0)
                rs_grdLiquida.Find "numero_pago = " & RegPuntero ' se posisicona en el registro editado
                Call grdLiquida_RowColChange(0, 0)
                MsgBox "Los cambios se realizaron satisfactoriamente", vbInformation, "Aviso"
                grdLiquida.SetFocus
            End If
            lblEstadoLiquida.Caption = "" ' no se esta editando ni adicionando registros
            
          Case "Tool_Editar"
            
            If rs_grdLiquida.RecordCount > 0 Then

                ' ***********************
                ' se llama al formulario
                ' ***********************
                Screen.MousePointer = vbHourglass
                lblEstadoLiquida.Caption = "E" ' se esta en modo de edicion del reg. actual
                lblEstadoLiquida.Tag = rs_grdLiquida!NUMERO_PAGO ' guara el nuemro de pago q se usa como parametro
                frm_ro_LiquidaAdiPago.Show vbModal

                If lblEstadoLiquida.Caption <> "E" Then   ' si el proceso realizo la edicion o moedidificcacion de los montos
                    RegPuntero = rs_grdLiquida!NUMERO_PAGO   ' numero pago
                    Call grdPrincipal_RowColChange(0, 0)
                    rs_grdLiquida.Find "numero_pago = " & RegPuntero ' se posisicona en el registro editado
                    Call grdLiquida_RowColChange(0, 0)
                    MsgBox "Los cambios se realizaron satisfactoriamente", vbInformation, "Aviso"
                    grdLiquida.SetFocus
                End If
                lblEstadoLiquida.Caption = "" ' no se esta editando ni adicionando registros
              Else
                MsgBox "No existen el registro de liquidación para ser editado.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
            
          Case "Tool_Anular"
            
            If rs_grdPrincipal.RecordCount > 0 Then ' si existen registros
                If fl_VerificaEliminaLiquida Then
                    Screen.MousePointer = vbHourglass
                    Cad = "Se eliminará el registro de liquidación número: [" & rs_grdLiquida!NUMERO_PAGO & "] del grupo: [" & lblCodGrupo.Caption & " - " & lblDesGrupo.Caption & "], Unidad: [" & rs_grdPrincipal!codigo_unidad & "]."
                    If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de elimnación") Then
                        'JQ QR
                        'De.dbo_ap_PagosBorraPago rs_grdPrincipal!ges_gestion, rs_grdPrincipal!codigo_unidad, rs_grdPrincipal!codigo_grupo, rs_grdLiquida!NUMERO_PAGO
                        Call pl_RefrescaLiquidacion
                        Call grdLiquida_RowColChange(0, 0)
                        MsgBox "La información se eliminó satisfactoriamente", vbInformation, "Aviso"
                    End If
                    Screen.MousePointer = vbDefault
                  Else
                    grdPrincipal.SetFocus
                End If
              Else
                MsgBox "No existen registros para ser eliminados.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
          
          Case "Tool_Aprobar"
          
            If rs_grdLiquida.RecordCount > 0 Then
                If rs_grdBeneficiario.RecordCount > 0 Then
                    
                    If fl_VerificaAprobar Then
                        Screen.MousePointer = vbHourglass
                        Cad = "Se aprobará la liquidación correspondiente a: " & Chr(13)
                        Cad = Cad & "Número de liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "] [" & rs_grdLiquida!Concepto & "]" & Chr(13)
                        Cad = Cad & "Beneficiarios: [" & rs_grdBeneficiario.RecordCount & "]" & Chr(13)
                        Cad = Cad & "Desea aprobar la liquidación?. No podrá modificar mas datos."
                        
                        If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de aprobación") Then
                            ' se actualiza las banderas de aprobacion a 'S' para q no puedan ser modificadas
                            'JQ QR
                            'De.dbo_ap_GetServDateTime fechax, horax
                            
                            If rs_grdPrincipal!ESTADO_APROBADO <> "S" Or rs_grdPrincipal!ESTADO_APROBADO <> "A" Then
                                SQLs = "UPDATE ao_pagos_grupos SET estado_aprobado ='S', usr_aprueba = '" & GlUsuario & "', fecha_aprueba = '" & fechax & "', hora_aprueba = '" & horax & "' "
                                SQLs = SQLs & "WHERE ges_gestion = '" & lblGestion.Caption & "' and codigo_unidad = '" & lblCodUniSol.Caption & "' and codigo_grupo = " & Val(lblCodGrupo.Caption)
                                'JQ QR
                                'De.dbo_apGeneralSearching SQLs
                            End If
                            
                            SQLs = "UPDATE ao_pagos_cronograma SET estado_aprobado ='S', usr_aprueba = '" & GlUsuario & "', fecha_aprueba = '" & fechax & "', hora_aprueba = '" & horax & "' "
                            SQLs = SQLs & "WHERE ges_gestion = '" & lblGestion.Caption & "' and codigo_unidad = '" & lblCodUniSol.Caption & "' and codigo_grupo = " & Val(lblCodGrupo.Caption) & " and numero_pago = " & rs_grdLiquida!NUMERO_PAGO
                            'JQ QR
                            'De.dbo_apGeneralSearching SQLs
                            
                            If rs_grdPrincipal!ESTADO_APROBADO <> "S" And rs_grdPrincipal!ESTADO_APROBADO <> "A" Then
                                fechax = lblGestion.Caption
                                Cad = lblCodUniSol.Caption
                                RegPuntero = rs_grdPrincipal!codigo_grupo
                                nro_reg = rs_grdLiquida!NUMERO_PAGO
                                ' se posisicona en el registro editado
                                Call pl_RefrescaListaPrincipal
                                rs_grdPrincipal.Find "ges_gestion = '" & fechax & "'"
                                rs_grdPrincipal.Find "codigo_unidad ='" & Cad & "'"
                                rs_grdPrincipal.Find "codigo_grupo =" & RegPuntero
                                Call grdPrincipal_RowColChange(0, 0)
                                rs_grdLiquida.Find "numero_pago =" & nro_reg
                                Call grdLiquida_RowColChange(0, 0)
                                MsgBox "Se aprobo la liquidación correspondiente.", vbInformation, "Aviso"
'                                grdLiquida.SetFocus
                              Else
                                RegPuntero = rs_grdLiquida!NUMERO_PAGO
                                Call grdPrincipal_RowColChange(0, 0)
                                rs_grdLiquida.Find "numero_pago =" & RegPuntero ' se posiciona en el registro editado
                                Call grdLiquida_RowColChange(0, 0)
                                MsgBox "Se aprobo la liquidación correspondiente.", vbInformation, "Aviso"
                            End If
                            
                         End If
                        grdBeneficiario.SetFocus
                        Screen.MousePointer = vbDefault
                    End If
                    
                  Else
                    MsgBox "No existe registro de beneficiario para ser procesado.", vbInformation, "Aviso"
                End If
              Else
                MsgBox "No existe registro de liquidación para ser procesado.", vbInformation, "Aviso"
            End If
            
          Case "Tool_Desaprobar"
          
            If rs_grdLiquida.RecordCount > 0 Then
                If rs_grdBeneficiario.RecordCount > 0 Then
                    Screen.MousePointer = vbHourglass
                    Cad = "Se desaprobará la liquidación correspondiente a: " & Chr(13)
                    Cad = Cad & "Número de liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "] [" & rs_grdLiquida!Concepto & "]" & Chr(13)
                    Cad = Cad & "Beneficiarios: [" & rs_grdBeneficiario.RecordCount & "]" & Chr(13)
                    Cad = Cad & "Desea desaprobar la liquidación?. Podrá modificar los datos."
                    
                    If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación") Then
                        'JQ QR
                        'De.dbo_ap_GetServDateTime fechax, horax
                        ' se actualiza las banderas de aprobacion a 'N'
                        SQLs = "SELECT * FROM ao_pagos_cronograma WHERE estado_DEVENGADO ='N' AND ges_gestion = '" & lblGestion.Caption & "' and codigo_unidad = '" & lblCodUniSol.Caption & "' and codigo_grupo = " & Val(lblCodGrupo.Caption)
                        Set rstTemp = New ADODB.Recordset
                        rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
                        If rstTemp.RecordCount = 0 Then
                            SQLs = "UPDATE ao_pagos_grupos SET estado_aprobado ='N', usr_aprueba = '" & GlUsuario & "', fecha_aprueba = '" & fechax & "', hora_aprueba = '" & horax & "' "
                            SQLs = SQLs & "WHERE ges_gestion = '" & lblGestion.Caption & "' and codigo_unidad = '" & lblCodUniSol.Caption & "' and codigo_grupo = " & Val(lblCodGrupo.Caption)
                            'JQ QR
                            'De.dbo_apGeneralSearching SQLs
                            
                            SQLs = "UPDATE ao_pagos_cronograma SET estado_aprobado ='N', usr_aprueba = '" & GlUsuario & "', fecha_aprueba = '" & fechax & "', hora_aprueba = '" & horax & "' "
                            SQLs = SQLs & "WHERE ges_gestion = '" & lblGestion.Caption & "' and codigo_unidad = '" & lblCodUniSol.Caption & "' and codigo_grupo = " & Val(lblCodGrupo.Caption) & " and numero_pago = " & rs_grdLiquida!NUMERO_PAGO
                            'JQ QR
                            'De.dbo_apGeneralSearching SQLs
                            
                            fechax = lblGestion.Caption
                            Cad = lblCodUniSol.Caption
                            RegPuntero = rs_grdPrincipal!codigo_grupo
                            nro_reg = rs_grdLiquida!NUMERO_PAGO
                            ' se posiciona en el registro editado
                            Call pl_RefrescaListaPrincipal
                            rs_grdPrincipal.Find "ges_gestion = '" & fechax & "'"
                            rs_grdPrincipal.Find "codigo_unidad ='" & Cad & "'"
                            rs_grdPrincipal.Find "codigo_grupo =" & RegPuntero
                            Call grdPrincipal_RowColChange(0, 0)
                            rs_grdLiquida.Find "numero_pago =" & nro_reg
                            Call grdLiquida_RowColChange(0, 0)
                            MsgBox "Se desaprobo la liquidación correpondiente.", vbInformation, "Aviso"
'                            grdLiquida.SetFocus
                          Else
                        
                            SQLs = "UPDATE ao_pagos_cronograma SET estado_aprobado ='N', usr_aprueba = '" & GlUsuario & "', fecha_aprueba = '" & fechax & "', hora_aprueba = '" & horax & "' "
                            SQLs = SQLs & "WHERE ges_gestion = '" & lblGestion.Caption & "' and codigo_unidad = '" & lblCodUniSol.Caption & "' and codigo_grupo = " & Val(lblCodGrupo.Caption) & " and numero_pago = " & rs_grdLiquida!NUMERO_PAGO
                            'JQ QR
                            'De.dbo_apGeneralSearching SQLs
                            
                            RegPuntero = rs_grdLiquida!NUMERO_PAGO
                            Call grdPrincipal_RowColChange(0, 0)
                            rs_grdLiquida.Find "numero_pago =" & RegPuntero ' se posiciona en el registro editado
                            Call grdLiquida_RowColChange(0, 0)
                            MsgBox "Se desaprobo la liquidación correspondiente", vbInformation, "Aviso"
                        
                        End If
                    End If
                    
                    grdBeneficiario.SetFocus
                    Screen.MousePointer = vbDefault
                  Else
                    MsgBox "No existe registro de beneficiario para ser procesado.", vbInformation, "Aviso"
                End If
              Else
                MsgBox "No existe registro de liquidación para ser procesado.", vbInformation, "Aviso"
            End If
          
          Case "Tool_Devengar"
            
            If fl_VerificaDevengar Then
                Screen.MousePointer = vbHourglass
                Cad = "Se devengará la liquidación correspondiente a: " & Chr(13)
                Cad = Cad & "Número de liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "] [" & rs_grdLiquida!Concepto & "]" & Chr(13)
                Cad = Cad & "Beneficiarios: [" & rs_grdBeneficiario.RecordCount & "]" & Chr(13)
                Cad = Cad & "Desea devengar la liquidación?."
                
                If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de aprobación") Then
                    
                    ' genera el devengado del comprometido para todo el pago
                    Call pl_GeneraDevengado
                    
                    RegPuntero = rs_grdLiquida!NUMERO_PAGO
                    Call grdPrincipal_RowColChange(0, 0)
                    rs_grdLiquida.Find "numero_pago =" & RegPuntero ' se posiciona en el registro editado
                    Call grdLiquida_RowColChange(0, 0)
                    
                    If rs_grdLiquida!estado_devengado = "S" Then
                        MsgBox "Se ha generado el Devengado del Comprometido con todo éxito", vbInformation, "Aviso"
                    End If
                    
                 End If
                 grdBeneficiario.SetFocus
                 Screen.MousePointer = vbDefault
            End If
          
          Case "Tool_AnulaDevengar"
            
            If rs_grdPrincipal.RecordCount > 0 Then
                If fl_VerificaAnulaOrdLiq Then
                    Screen.MousePointer = vbHourglass
                    Cad = "Se anulará la orden de liquidación y el devengado correspondiente a: " & Chr(13)
                    Cad = Cad & "Número de liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "] [" & rs_grdLiquida!Concepto & "]" & Chr(13)
                    Cad = Cad & "Beneficiarios: [" & rs_grdBeneficiario.RecordCount & "]" & Chr(13)
                    Cad = Cad & "Desea anular la orden de liquidación y su devengado correspondiente?."
                    
                    If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de aprobación") Then
                        
                        ' anula la orden de liquidación y el devengado del comprometido para todo el pago
                        'JQ QR
                        'De.dbo_ap_PagosAnulaOrdLiq_c lblGestion.Caption, lblCodUniSol.Caption, Val(lblCodGrupo.Caption), rs_grdLiquida!NUMERO_PAGO, rs_grdLiquida!correlativo_reg
                        RegPuntero = rs_grdLiquida!NUMERO_PAGO
                        Call pl_RefrescaLiquidacion
                        MsgBox "Se anulo la Orden de Liquidación y el Devengado del Comprometido con todo éxito.", vbInformation, "Aviso"
                     
                     End If
                     grdBeneficiario.SetFocus
                     Screen.MousePointer = vbDefault
                End If
              Else
                MsgBox "No existen registros para procesar.", vbInformation, "Aviso"
            End If
            
          Case "Tool_Salir"
            Unload Me
            
        End Select
    
      ' ********************************************************
      ' opciones genericas de: BENEFICIARIOS
      ' ********************************************************
      
      Case "Beneficiario" ' procesa la cuadro BENEFICIARIOS
        Select Case TipoOpcion
          Case "Tool_Contrato"
            If rs_grdBeneficiario.RecordCount > 0 Then
                    
                ' ***********************
                ' se llama al formulario
                ' ***********************
                Screen.MousePointer = vbHourglass
                grdBeneficiario.SetFocus
                lblEstadoBeneficiario.Caption = rs_grdBeneficiario!Nombre & " " & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno  ' codigo de beneficario usadocomo parametro
                lblEstadoBeneficiario.Tag = rs_grdBeneficiario!codigo_beneficiario ' codigo de beneficario usadocomo parametro
               
                frm_ro_ConfirmaFechasContrato.Show vbModal

              Else
                MsgBox "No existen el registro de beneficiario para ser editado.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If

          Case "Tool_Nuevo"
            If rs_grdLiquida.RecordCount > 0 Then
                ' adiciona un beneficiario al pago si no tiene beneficiarios o modalidad planilla
                If rs_grdBeneficiario.RecordCount = 0 Or rs_grdPrincipal!modalidad_pago = "P" Then
                    lblEstadoBeneficiario.Caption = "N"
                    '***************************
                    'llama al formulario
                    '***************************
                    Screen.MousePointer = vbHourglass
                    grdBeneficiario.Tag = rs_grdLiquida!NUMERO_PAGO ' guarda el numero de pago como paremtro
                    lblEstadoBeneficiario.Tag = rs_grdLiquida!correlativo_reg
                    frm_ro_SelecBenLiquida.Show vbModal
                    
                    If lblEstadoBeneficiario.Caption <> "N" Then ' es distinto se guardo el codigo de beneficiario
                        RegPuntero = rs_grdLiquida!NUMERO_PAGO ' se guarda para luego ser ubicado
                        Cad = lblEstadoBeneficiario.Caption ' se gurada el codigo beneficiario
                        
                        Call pl_RefrescaLiquidacion
                        rs_grdLiquida.Find "numero_pago = " & RegPuntero
                        Call grdLiquida_RowColChange(0, 0)
                        rs_grdBeneficiario.Find "codigo_beneficiario = '" & Cad & "'"  ' se posisicona en el registro editado
                        MsgBox "Se adiciono beneficiario(s).", vbInformation, "Aviso"
                        grdBeneficiario.SetFocus
                        
                    End If
                    
                    lblEstadoBeneficiario.Caption = "" ' no se esta editando ni adicionando registros
                  
                  Else
                    MsgBox "La modalidad de liquidación es planilla individual.", vbInformation, "Aviso"
                    grdLiquida.SetFocus
                    
                End If
              
              Else
                MsgBox "No existen registrado número de Liquidación.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
            
          Case "Tool_Monto"
            
            If rs_grdBeneficiario.RecordCount > 0 Then
            
                ' verifica si el comprobante de pago existe y esta aprobado por tesoreria
                Set rstTemp = New ADODB.Recordset
                SQLs = "SELECT r.aprobotesoreria FROM ac_ben_comprdeven r  WHERE r.codigo_beneficiario = '" & rs_grdBeneficiario!codigo_beneficiario & "' and "
                SQLs = SQLs & "r.gp_codigo_unidad = '" & lblCodUniSol.Caption & "' and "
                SQLs = SQLs & "r.gp_codigo_grupo = '" & lblCodGrupo.Caption & "' and "
                SQLs = SQLs & "r.ges_Gestion     = '" & lblGestion.Caption & "' and "
                SQLs = SQLs & "r.tipocomprobante = 'COM' AND "
                SQLs = SQLs & "r.APROBOTESORERIA IN ('N','S') "

                rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
                
                If rstTemp.RecordCount > 0 Then
                    If rstTemp!APROBOTESORERIA = "S" Then ' EN AC_BEN_COMPRDEVEN
                        
                        ' ***********************
                        ' se llama al formulario
                        ' ***********************
                        Screen.MousePointer = vbHourglass
                        grdBeneficiario.SetFocus
                        lblEstadoBeneficiario.Caption = "E" ' se esta en modo de edicion del reg. actual
                        grdLiquida.Tag = rs_grdLiquida!correlativo_reg ' numero correlativo
                        grdBeneficiario.Tag = rs_grdLiquida!NUMERO_PAGO ' numero de pago
                        lblEstadoBeneficiario.Tag = rs_grdBeneficiario!codigo_beneficiario ' codigo de beneficario usadocomo parametro

                        frm_ro_LiquidaMontoBen.Show vbModal

                        If lblEstadoBeneficiario.Caption <> "E" Then  ' si el proceso realizo alguna modificacion
                            RegPuntero = rs_grdLiquida!NUMERO_PAGO ' nujero de pago
                            Cad = rs_grdBeneficiario!codigo_beneficiario ' codigo
                            pl_RefrescaLiquidacion
                            rs_grdLiquida.Find "numero_pago =" & RegPuntero
                            Call grdLiquida_RowColChange(0, 0)
                            rs_grdBeneficiario.Find "codigo_beneficiario = '" & Cad & "'"  ' se posisicona en el registro editado
                            
                            MsgBox "Los cambios se realizaron satisfactoriamente", vbInformation, "Aviso"
                            grdBeneficiario.SetFocus
                                                    
                        End If
                        lblEstadoBeneficiario.Caption = "" ' no se esta editando ni adicionando registros
                        
                      Else
                        MsgBox "El comprobante de liquidación correspondiente a [" & rs_grdBeneficiario!Nombre & " " & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & "]." & Chr(13) & " No se encuentra aprobado en presupuestos." & Chr(13) & Chr(13) & "Verifique el proceso....gracias.", vbCritical, "Error"
                        grdBeneficiario.SetFocus
                    End If
                  Else
                    MsgBox "No existen registro de comprobante de liquidación correspondiente a [" & rs_grdBeneficiario!Nombre & " " & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & "]." & Chr(13) & " para ser procesado." & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
                    grdBeneficiario.SetFocus
                End If
              Else
                MsgBox "No existen el registro de beneficiario para ser editado.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
            
          Case "Tool_Anular"
            
            If rs_grdBeneficiario.RecordCount > 0 Then ' si existen registros
                If rs_grdLiquida!estado_devengado & "" <> "S" Then
                    Screen.MousePointer = vbHourglass
                    Cad = "Se eliminará el registro del beneficiario [" & rs_grdBeneficiario!Nombre & " " & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & "] correspondiente a:" & Chr(13)
                    Cad = Cad & "Unidad: [" & lblCodUniSol.Caption & " - " & lblDesUniSol.Caption & "]" & Chr(13) & "Grupo: [" & lblCodGrupo.Caption & " - " & lblDesGrupo.Caption & "]" & Chr(13) & "Nro. liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "]." & Chr(13)
                    If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de elimnación") Then
                        'JQ QR
                        'De.dbo_ap_PagosBorraPagoBenef rs_grdPrincipal!ges_gestion, rs_grdPrincipal!codigo_unidad, rs_grdPrincipal!codigo_grupo, rs_grdLiquida!NUMERO_PAGO, rs_grdBeneficiario!codigo_beneficiario
                        
                        RegPuntero = rs_grdLiquida!NUMERO_PAGO
                        Call pl_RefrescaLiquidacion
                        rs_grdLiquida.Find "numero_pago =" & RegPuntero
                        Call grdLiquida_RowColChange(0, 0)
                        
                        Call pl_RefrescaBeneficiario
                        
                        MsgBox "La información se eliminó satisfactoriamente", vbInformation, "Aviso"
                    End If
                    Screen.MousePointer = vbDefault
                  Else
                    MsgBox "No puede eliminar el registro del beneficiario [" & rs_grdBeneficiario!Nombre & " " & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & "] por que tiene devengado generado.", vbInformation, "Aviso"
                    grdBeneficiario.SetFocus
                End If
              Else
                MsgBox "No existen registros para ser eliminados.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
          
          Case "Tool_AnularTodo"
            
            If rs_grdBeneficiario.RecordCount > 0 Then ' si existen registros
                If rs_grdLiquida!estado_devengado & "" <> "S" Then
                    Screen.MousePointer = vbHourglass
                    Cad = "Se eliminará TODOS los registros de beneficiarios correspondiente a:" & Chr(13)
                    Cad = Cad & "Unidad: [" & lblCodUniSol.Caption & " - " & lblDesUniSol.Caption & "]" & Chr(13) & "Grupo: [" & lblCodGrupo.Caption & " - " & lblDesGrupo.Caption & "]" & Chr(13) & "Nro. liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "]." & Chr(13)
                    If vbYes = MsgBox(Cad, vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación de elimnación") Then
                        'JQ QR
                        'De.dbo_apGeneralSearching "DELETE ao_pagos_cronograma_detalle WHERE ges_gestion = '" & rs_grdPrincipal!ges_gestion & "' and codigo_unidad   = '" & rs_grdPrincipal!codigo_unidad & "' and codigo_grupo = " & rs_grdPrincipal!codigo_grupo & " and numero_pago = " & rs_grdLiquida!NUMERO_PAGO
                        'De.dbo_apGeneralSearching "UPDATE ao_pagos_cronograma set tipo_moneda = '', monto_us = 0, monto_bs = 0 WHERE ges_gestion = '" & rs_grdPrincipal!ges_gestion & "' and codigo_unidad   = '" & rs_grdPrincipal!codigo_unidad & "' and codigo_grupo = " & rs_grdPrincipal!codigo_grupo & " and numero_pago = " & rs_grdLiquida!NUMERO_PAGO
                        
                        RegPuntero = rs_grdLiquida!NUMERO_PAGO
                        Call pl_RefrescaLiquidacion
                        rs_grdLiquida.Find "numero_pago =" & RegPuntero
                        Call grdLiquida_RowColChange(0, 0)
                        
                        Call pl_RefrescaBeneficiario
                        
                        MsgBox "La información se eliminó satisfactoriamente", vbInformation, "Aviso"
                    End If
                    Screen.MousePointer = vbDefault
                  Else
                    MsgBox "No puede eliminar los beneficiarios por que tiene devengado generado.", vbInformation, "Aviso"
                    grdBeneficiario.SetFocus
                End If
              Else
                MsgBox "No existen registros para ser eliminados.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
            End If
                    
          Case "Tool_Conformidad"
            
            If rs_grdLiquida.RecordCount > 0 Then
                If rs_grdLiquida!ESTADO_APROBADO <> "S" Then ' verificamos
                    
                    If rs_grdBeneficiario.RecordCount > 0 Then
                    
                        If fl_VerificaConformidad Then

                            ' *********************************************
                            ' se llama al formulario q permite resgistrar cite
                            ' *********************************************
                            Screen.MousePointer = vbHourglass
                            lblEstadoBeneficiario.Caption = "C" ' parametro para registrar conformidad
                            grdBeneficiario.Tag = rs_grdLiquida!NUMERO_PAGO
                            frm_ro_LiquidaConformidad.Show vbModal
                            lblEstadoBeneficiario.Caption = ""
                            
                            Call grdLiquida_RowColChange(0, 0)
                            grdLiquida.SetFocus
                        End If
                      Else
                        MsgBox "No se tiene registro de beneficiarios.", vbInformation, "Aviso"
                        grdLiquida.SetFocus
                    End If
                  Else
                    MsgBox "La liquidación número [" & rs_grdLiquida!NUMERO_PAGO & "] se encuentra aprobada.", vbInformation, "Aviso"
                    grdPrincipal.SetFocus
                End If
              Else
                MsgBox "No existe registro de liquidación para registrar conformidad.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
                
            End If
            
          Case "Tool_Factura"
            
            If rs_grdLiquida.RecordCount > 0 Then
                If rs_grdLiquida!ESTADO_APROBADO <> "S" Then ' verificamos si tiene comprobante presupuestario
                    If rs_grdBeneficiario.RecordCount > 0 Then
                    
                        ' *********************************************
                        ' se llama al formulario q permite resgistrar si emite factura
                        ' *********************************************
                        Screen.MousePointer = vbHourglass
                        lblEstadoBeneficiario.Caption = "F" 'parametro para procesar registro de emite o no factura
                        grdBeneficiario.Tag = rs_grdLiquida!NUMERO_PAGO
                        frm_ro_LiquidaConformidad.Show vbModal
                        lblEstadoBeneficiario.Caption = ""
                        
                        Call grdLiquida_RowColChange(0, 0)
                        grdPrincipal.SetFocus
                      Else
                        MsgBox "No se tiene registro de beneficiarios.", vbInformation, "Aviso"
                        grdLiquida.SetFocus
                    End If
                  Else
                    MsgBox "La liquidación número [" & rs_grdLiquida!NUMERO_PAGO & "] se encuentra aprobada.", vbInformation, "Aviso"
                    grdPrincipal.SetFocus
                  
                End If
                
              Else
                MsgBox "No existe registro de liquidación para registrar si emite factura.", vbInformation, "Aviso"
                grdPrincipal.SetFocus
                
            End If
          
          
          Case "Tool_Salir"
            Unload Me
            
        End Select
    
    End Select
      
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    ' si se produjo otro tipo de error
    MsgBox "Error: Se produjo un error." & Chr(13) & Chr(13) & "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Function fl_VerificaEliminaGrupo() As Boolean
    'TITULO:                Función fl_VerificaEliminaGrupo
    'PROPOSITO:             Verifica los datos para procesar la elimnacion
    'EJEMPLO DE LLAMADA:    fl_VerificaEliminaGrupo
    Dim rstTemp As ADODB.Recordset ' usado para la carga de los combos de base
    
    fl_VerificaEliminaGrupo = True
    
    ' verifica si tiene compromisos de pago aprobados
    SQLs = "select * from ac_ben_comprDeven where gp_ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and gp_codigo_unidad ='" & rs_grdPrincipal!codigo_unidad & "' and gp_codigo_grupo =" & rs_grdPrincipal!codigo_grupo & " and tipoComprobante ='COM' and aprobotesoreria='S'"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        MsgBox "No puede eliminar el grupo [" & lblCodGrupo.Caption & "][" & lblDesGrupo.Caption & "] de liquidación por tener compromiso de pago APROBADO." & Chr(13) & "Comuniquese con el administrador del sistema.", vbInformation, "Aviso"
        grdPrincipal.SetFocus
        fl_VerificaEliminaGrupo = False
        Exit Function
    End If
        
    ' verificamos si tiene algun pago devengado
    SQLs = "select * from ao_pagos_cronograma where ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and codigo_unidad='" & rs_grdPrincipal!codigo_unidad & "' and codigo_grupo=" & rs_grdPrincipal!codigo_grupo & " and estado_devengado ='S'"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    If rstTemp.RecordCount > 0 Then
        MsgBox "No puede eliminar el grupo [" & lblCodGrupo.Caption & "][" & lblDesGrupo.Caption & "] porque tiene ordenes de pago elaboradas." & Chr(13) & "Comuniquese con el administrador del sistema.", vbInformation, "Aviso"
        grdPrincipal.SetFocus
        fl_VerificaEliminaGrupo = False
        Exit Function
    End If
    
    Set rstTemp = Nothing
    
End Function

Private Sub pl_HabilitaUnaOpcion(Boton As String, swModo As Boolean, Proceso As String)
    'TITULO:                Procedimiento pl_HabilitaUnaOpcion
    'PROPOSITO:             Habilita o deshabilita el boton especificado del toolbar
    'EJEMPLO DE LLAMADA:    call pl_HabilitaUnaOpcion(NombreBoton, true/false, Proceso)
    
    Select Case Proceso
      Case "GrupoLiq"
        ' habilitamos o deshabilitamos las opciones del menu
        Select Case Boton
          Case "Tool_Nuevo"
            cmdNuevo.Enabled = swModo
          Case "Tool_Editar"
            cmdEditar.Enabled = swModo
          Case "Tool_Anular"
            cmdAnular.Enabled = swModo
          Case "Tool_Salir"
            cmdSalir.Enabled = swModo
        End Select
      
      Case "Liquida"
        ' habilitamos o deshabilitamos las opciones del menu
        Select Case Boton
          Case "Tool_Nuevo"
            cmdPagoNuevo.Enabled = swModo
          Case "Tool_Editar"
            cmdPagoEditar.Enabled = swModo
          Case "Tool_Anular"
            cmdPagoAnular.Enabled = swModo
          Case "Tool_Aprobar"
            cmdPagoAprob.Enabled = swModo
          Case "Tool_Desaprobar"
            cmdPagoDesaprob.Enabled = swModo
          Case "Tool_Devengar"
            cmdPagoDeven.Enabled = swModo
          Case "Tool_AnulaDevengar"
            cmdPagoAnulaDev.Enabled = swModo
        End Select
    
      Case "Beneficiario"
        ' habilitamos o deshabilitamos las opciones del menu
        Select Case Boton
          Case "Tool_Nuevo"
            cmdBenNuevo.Enabled = swModo
          Case "Tool_Monto"
            cmdBenMonto.Enabled = swModo
          Case "Tool_Anular"
            cmdBenAnular.Enabled = swModo
          Case "Tool_AnularTodo"
            cmdBenAnularTodo.Enabled = swModo
          Case "Tool_Conformidad"
            cmdBenConf.Enabled = swModo
          Case "Tool_Factura"
            cmdBenFact.Enabled = swModo
        End Select
    
    End Select
End Sub

Private Sub grdBeneficiario_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call pl_ControlaToolBar("Beneficiario")
End Sub

Private Sub grdPrincipal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'PROPOSITO:             Permite desplazarse sobre el browse actualizando los datos de las fichas
    
    On Error GoTo EtiqError

    ' datos cabecera
    If rs_grdPrincipal.RecordCount = 0 Then
        lblGestion.Caption = "" ' gestion
        lblFormulario.Caption = "" ' tipo de trammite
        lblCodUniSol.Caption = "" ' unidad solicitante
        lblDesUniSol.Caption = "" ' unidad solicitante
        lblCodGrupo.Caption = "" ' codigo grupo
        lblDesGrupo.Caption = "" ' descripcion grupo
        lblCodGrupo.Tag = "" ' para guardar el numero de consultoria
        lblDesGrupo.Tag = "" ' para guardar el tipo de liqudacion planialla o individual
      Else
        lblGestion.Caption = rs_grdPrincipal!ges_gestion & ""    ' gestion
        lblFormulario.Caption = rs_grdPrincipal!formulario & " - " & rs_grdPrincipal!Denominacion_Tipo    ' tipo de tramite
        lblCodUniSol.Caption = rs_grdPrincipal!codigo_unidad  ' unidad solicitante
        lblDesUniSol.Caption = rs_grdPrincipal!Uni_descripcion_larga   ' unidad solicitante
        lblCodGrupo.Caption = rs_grdPrincipal!codigo_grupo & "" ' codigo grupo
        lblDesGrupo.Caption = rs_grdPrincipal!descripcion_grupo & "" ' descripcion grupo
        lblCodGrupo.Tag = rs_grdPrincipal!numero_consultoria & "" ' numero conultoria
        lblDesGrupo.Tag = rs_grdPrincipal!modalidad_pago & "" ' modalidad de pago
        
    End If
    
    ' GRUPO LIQUIDACION
    Call pl_RefrescaLiquidacion
    Call pl_ControlaToolBar("GrupoLiq")
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub pl_RefrescaLiquidacion()
    'TITULO:                Procedimiento pl_RefrescaLiquidacion
    'PROPOSITO:             Actualiza los datos de la ficha
    'EJEMPLO DE LLAMADA:    call pl_RefrescaLiquidacion
    
    On Error GoTo EtiqError
        
    ' se actualiza el grid
    SQLs = "SELECT numero_pago, concepto, tipo_moneda, monto_us, monto_bs, estado_aprobado, estado_devengado, antecedente, codigo_orden, fecha_estimada_liq, correlativo_reg "
    SQLs = SQLs & "FROM ao_pagos_cronograma "
    SQLs = SQLs & "WHERE ges_gestion = '" & lblGestion.Caption & "' AND codigo_unidad ='" & lblCodUniSol.Caption & "' AND codigo_grupo =" & Val(lblCodGrupo.Caption)
    SQLs = SQLs & " and estado_devengado <> 'E'"

    Set rs_grdLiquida = New ADODB.Recordset
    rs_grdLiquida.Open SQLs, db, adOpenStatic, adLockReadOnly
   
    Set grdLiquida.DataSource = rs_grdLiquida
    grdLiquida.Caption = "Liquidación del grupo: [" & Val(lblCodGrupo.Caption) & " - " & IIf(rs_grdPrincipal.RecordCount = 0, "", rs_grdPrincipal!descripcion_grupo) & "]."
    lblNroLiquida.Caption = "Nro. de liquidaciones: " & rs_grdLiquida.RecordCount
    Call pl_PersonalizaGridLiquida

    ' calcula totales de liquidación
    SQLs = "SELECT 'total_us' = sum(monto_us), 'total_bs' = sum(monto_bs) "
    SQLs = SQLs & "FROM ao_pagos_cronograma "
    SQLs = SQLs & "WHERE ges_gestion = '" & lblGestion.Caption & "' AND codigo_unidad ='" & lblCodUniSol.Caption & "' AND codigo_grupo =" & Val(lblCodGrupo.Caption)
    SQLs = SQLs & " and estado_devengado <> 'E'"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    lblTotalUSLiq = Format(IIf(IsNull(rstTemp!total_us), 0, rstTemp!total_us), "##,##0.00") & " $US"
    lblTotalBSLiq = Format(IIf(IsNull(rstTemp!total_bs), 0, rstTemp!total_bs), "##,##0.00") & " Bs"
    
    Call pl_RefrescaBeneficiario
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub pl_PersonalizaGridLiquida()
    'TITULO:                Procedimiento pl_PersonalizaGridLiquida
    'PROPOSITO:             Personalizar los captions, anchos, etc. del grid
    'EJEMPLO DE LLAMADA:    call pl_PersonalizaGridLiquida
        
    Dim i As Integer
    
    ' define ancho de columnas y titulo de la cabecera
    grdLiquida.Columns(0).Width = 900 ' numero de liq.
    grdLiquida.Columns(0).Caption = "Nro. liquid."
    grdLiquida.Columns(1).Width = 2000 ' concepto
    grdLiquida.Columns(1).Caption = "Concepto"
    grdLiquida.Columns(2).Width = 1000 ' tipo moneda
    grdLiquida.Columns(2).Caption = "Tipo moneda"
    grdLiquida.Columns(3).Width = 1000 ' monto us
    grdLiquida.Columns(3).Caption = "Monto US"
    grdLiquida.Columns(4).Width = 1000 ' monto BS
    grdLiquida.Columns(4).Caption = "Monto BS"
    grdLiquida.Columns(5).Width = 900 ' estado aprobado
    grdLiquida.Columns(5).Caption = "Est. aprobado"
    grdLiquida.Columns(6).Width = 900 ' estado devengado
    grdLiquida.Columns(6).Caption = "Est. devengado"
    grdLiquida.Columns(7).Width = 900 ' antecedente
    grdLiquida.Columns(7).Caption = "Antecedente"
    grdLiquida.Columns(8).Width = 1000 ' codigo orden
    grdLiquida.Columns(8).Caption = "Cod. orden"
    grdLiquida.Columns(9).Width = 1000 ' fecha estimada liquidacion
    grdLiquida.Columns(9).Caption = "F. estimada liquidación"

End Sub

Private Sub grdLiquida_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'PROPOSITO:             Permite desplazarse sobre el browse actualizando los datos de las fichas

    On Error GoTo EtiqError

    If rs_grdLiquida.RecordCount = 0 Then
        lblEstadoBeneficiario.Tag = ""
        Call pl_RefrescaBeneficiario
        Call grdBeneficiario_RowColChange(0, 0)
        Call pl_ControlaToolBar("Liquida")
      Else
        lblEstadoBeneficiario.Tag = rs_grdLiquida!NUMERO_PAGO & ""
        Call pl_RefrescaBeneficiario
        Call grdBeneficiario_RowColChange(0, 0)
        Call pl_ControlaToolBar("Liquida")
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Function fl_VerificaEliminaLiquida() As Boolean
    'TITULO:                Función fl_VerificaEliminaLiquida
    'PROPOSITO:             Verifica los datos para procesar la elimnacion
    'EJEMPLO DE LLAMADA:    fl_VerificaEliminaLiquida
    Dim rstTemp As ADODB.Recordset ' usado para la carga de los combos de base
    
    fl_VerificaEliminaLiquida = True

    If rs_grdLiquida!estado_devengado = "S" Then
        MsgBox "No puede eliminar la liquidación por estar generado correspondiente devengado." & "", vbInformation, "Aviso"
        grdLiquida.SetFocus
        fl_VerificaEliminaLiquida = False
        Exit Function
    End If
    
    If rs_grdLiquida!ESTADO_APROBADO = "S" Then
        MsgBox "No puede eliminar la liquidación por estar aprobado." & "", vbInformation, "Aviso"
        grdLiquida.SetFocus
        fl_VerificaEliminaLiquida = False
        Exit Function
    End If
    
    ' verifica si existe pagos superiores
    SQLs = "SELECT * FROM ao_pagos_cronograma "
    SQLs = SQLs & "WHERE ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and codigo_unidad ='" & rs_grdPrincipal!codigo_unidad & "' and codigo_grupo =" & rs_grdPrincipal!codigo_grupo & " and numero_pago> " & rs_grdLiquida!NUMERO_PAGO & " and estado_aprobado <>'E'and estado_devengado <>'E'"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        MsgBox "No puede eliminar la liquidación [" & rs_grdLiquida!NUMERO_PAGO & "] del grupo [" & lblCodGrupo.Caption & " - " & lblDesGrupo.Caption & "] por tener registro de pagos superiores." & Chr(13) & "Corrija el error e intente eliminar nuevamente.", vbInformation, "Aviso"
        grdLiquida.SetFocus
        fl_VerificaEliminaLiquida = False
        Exit Function
    End If
    
    Set rstTemp = Nothing
    
End Function


''*******************************************************
''procesos de la ficha BENEFICIARIOS
''*******************************************************

Private Sub pl_RefrescaBeneficiario()
    'TITULO:                Procedimiento pl_RefrescaBeneficiario
    'PROPOSITO:             Actualiza los datos de la ficha beneficiario
    'EJEMPLO DE LLAMADA:    call pl_RefrescaBeneficiario
    
    On Error GoTo EtiqError
    
    ' obtiene datos de beneficiarios del pago
    ' dependiendo del tipo de proceso si es consultor por F05 ==> "producto - corto plazo"  o F10 ==> consultor por "tiempo - largo pazo"
    Select Case glProceso
      Case "F05"
        SQLs = "SELECT ao_pagos_cronograma_detalle.numero_pago, fc_beneficiario.paterno_beneficiario as paterno, fc_beneficiario.materno_beneficiario as materno, fc_beneficiario.nombres_beneficiario as nombre, ao_pagos_cronograma_detalle.codigo_beneficiario, ao_pagos_cronograma_detalle.monto_us, ao_pagos_cronograma_detalle.monto_bs, ao_pagos_cronograma_detalle.tc_us, ao_pagos_cronograma_detalle.tipo_moneda,"
        SQLs = SQLs & "ao_pagos_cronograma_detalle.emite_factura, ao_pagos_cronograma_detalle.estado_conformidad, ao_pagos_cronograma_detalle.estado_devengado, ao_pagos_cronograma_detalle.ncite_conformidad, ao_pagos_cronograma_detalle.fcite_conformidad, ao_pagos_cronograma_detalle.Numero_consultoriaHist, ao_pagos_cronograma_detalle.fte_financiamientoHist, ao_pagos_cronograma_detalle.correlativo_reg  "
        SQLs = SQLs & "FROM ao_pagos_cronograma_detalle INNER JOIN fc_beneficiario ON ao_pagos_cronograma_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario "
        SQLs = SQLs & "WHERE ao_pagos_cronograma_detalle.ges_gestion = '" & lblGestion.Caption & "' "
        SQLs = SQLs & " AND ao_pagos_cronograma_detalle.codigo_grupo = " & Val(lblCodGrupo.Caption)
        SQLs = SQLs & " AND ao_pagos_cronograma_detalle.codigo_unidad = '" & lblCodUniSol.Caption & "' "
        SQLs = SQLs & " AND ao_pagos_cronograma_detalle.numero_pago = " & IIf(rs_grdLiquida.RecordCount = 0, 9999, rs_grdLiquida!NUMERO_PAGO) & " AND ao_pagos_cronograma_detalle.correlativo_reg = " & IIf(rs_grdLiquida.RecordCount = 0, 9999, rs_grdLiquida!correlativo_reg)
        SQLs = SQLs & " and ao_pagos_cronograma_detalle.estado_devengado <> 'E' "
        SQLs = SQLs & " ORDER BY paterno, materno, nombre"
      
      Case "F10"
        SQLs = "SELECT ao_pagos_cronograma_detalle.numero_pago, RC_Personal.paterno as paterno, RC_Personal.materno as materno, RC_Personal.nombres as nombre, ao_pagos_cronograma_detalle.codigo_beneficiario, ao_pagos_cronograma_detalle.monto_us, ao_pagos_cronograma_detalle.monto_bs, ao_pagos_cronograma_detalle.tc_us, ao_pagos_cronograma_detalle.tipo_moneda,"
        SQLs = SQLs & "ao_pagos_cronograma_detalle.emite_factura, ao_pagos_cronograma_detalle.estado_conformidad, ao_pagos_cronograma_detalle.estado_devengado, ao_pagos_cronograma_detalle.ncite_conformidad, ao_pagos_cronograma_detalle.fcite_conformidad, ao_pagos_cronograma_detalle.Numero_consultoriaHist, ao_pagos_cronograma_detalle.fte_financiamientoHist, ao_pagos_cronograma_detalle.correlativo_reg "
        SQLs = SQLs & "FROM ao_pagos_cronograma_detalle INNER JOIN RC_Personal ON ao_pagos_cronograma_detalle.codigo_beneficiario = RC_Personal.ci "
        SQLs = SQLs & "WHERE ao_pagos_cronograma_detalle.ges_gestion = '" & lblGestion.Caption & "' "
        SQLs = SQLs & " AND ao_pagos_cronograma_detalle.codigo_grupo = " & Val(lblCodGrupo.Caption)
        SQLs = SQLs & " AND ao_pagos_cronograma_detalle.codigo_unidad = '" & lblCodUniSol.Caption & "' "
        SQLs = SQLs & " AND ao_pagos_cronograma_detalle.numero_pago = " & IIf(rs_grdLiquida.RecordCount = 0, 9999, rs_grdLiquida!NUMERO_PAGO) & " AND ao_pagos_cronograma_detalle.correlativo_reg = " & IIf(rs_grdLiquida.RecordCount = 0, 9999, rs_grdLiquida!correlativo_reg)
        SQLs = SQLs & " and ao_pagos_cronograma_detalle.estado_devengado <> 'E' "
        SQLs = SQLs & " ORDER BY paterno, materno, nombre"
      
    End Select
    Set rs_grdBeneficiario = New ADODB.Recordset
    rs_grdBeneficiario.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    Set grdBeneficiario.DataSource = rs_grdBeneficiario
    grdBeneficiario.Caption = "Beneficiarios del pago numero:[" & IIf(rs_grdLiquida.RecordCount = 0, 0, rs_grdLiquida!NUMERO_PAGO) & "]."
    
    Call pl_PersonalizaGridBeneficiario
    
    lblNroBeneficiario.Caption = "Nro. de beneficiarios: " & rs_grdBeneficiario.RecordCount
    
    ' calculamos montos totales por el numero de liquidacion
    
    SQLs = "SELECT 'total_US' = SUM(ao_pagos_cronograma_detalle.monto_us), 'total_BS' = SUM(ao_pagos_cronograma_detalle.monto_bs) "
    SQLs = SQLs & "FROM ao_pagos_cronograma_detalle "
    SQLs = SQLs & "WHERE ao_pagos_cronograma_detalle.ges_gestion = '" & lblGestion.Caption & "' "
    SQLs = SQLs & " AND ao_pagos_cronograma_detalle.codigo_grupo = " & Val(lblCodGrupo.Caption)
    SQLs = SQLs & " AND ao_pagos_cronograma_detalle.codigo_unidad = '" & lblCodUniSol.Caption & "' "
    SQLs = SQLs & " AND ao_pagos_cronograma_detalle.numero_pago = " & IIf(rs_grdLiquida.RecordCount = 0, 9999, rs_grdLiquida!NUMERO_PAGO) & " and ao_pagos_cronograma_detalle.correlativo_reg = " & IIf(rs_grdLiquida.RecordCount = 0, 999, rs_grdLiquida!correlativo_reg)
    SQLs = SQLs & " and ao_pagos_cronograma_detalle.estado_devengado <> 'E' "
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    lblTotalUS.Caption = Format(IIf(IsNull(rstTemp!total_us), 0, rstTemp!total_us), "######0.00") & " $US" ' total asignado pie de grid
    lblTotalBS.Caption = Format(IIf(IsNull(rstTemp!total_bs), 0, rstTemp!total_bs), "######0.00") & " Bs" ' total asignado pie de grid
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub pl_PersonalizaGridBeneficiario()
    'TITULO:                Procedimiento pl_PersonalizaGridBeneficiario
    'PROPOSITO:             Personalizar los captions, anchos, etc. del grid
    'EJEMPLO DE LLAMADA:    call pl_PersonalizaGridBeneficiario

    ' define ancho de columnas y titulo de la cabecera
    grdBeneficiario.Columns(0).Width = 400 ' Nro.
    grdBeneficiario.Columns(0).Caption = "No.Liq."
    grdBeneficiario.Columns(1).Width = 1000 ' paterno
    grdBeneficiario.Columns(1).Caption = "Paterno"
    grdBeneficiario.Columns(2).Width = 1000 ' materno
    grdBeneficiario.Columns(2).Caption = "Materno"
    grdBeneficiario.Columns(3).Width = 1200 ' nombres
    grdBeneficiario.Columns(3).Caption = "Nombre(s)"
    grdBeneficiario.Columns(4).Width = 900 ' codigo ben
    grdBeneficiario.Columns(4).Caption = "Cod. Benef."
    grdBeneficiario.Columns(5).Width = 1000 ' monto us
    grdBeneficiario.Columns(5).Caption = "Monto US"
    grdBeneficiario.Columns(6).Width = 1000 ' monto bs
    grdBeneficiario.Columns(6).Caption = "Monto BS"
    grdBeneficiario.Columns(7).Width = 500 ' tc us
    grdBeneficiario.Columns(7).Caption = "Tc US"
    grdBeneficiario.Columns(8).Width = 600 ' moneda
    grdBeneficiario.Columns(8).Caption = "Moneda"
    grdBeneficiario.Columns(9).Width = 500 ' emite factura
    grdBeneficiario.Columns(9).Caption = "Emite factura"
    grdBeneficiario.Columns(10).Width = 900 ' estado conformidad
    grdBeneficiario.Columns(10).Caption = "Est. conformidad"
    grdBeneficiario.Columns(11).Width = 900 ' estado devengado
    grdBeneficiario.Columns(11).Caption = "Est. devengado"
    grdBeneficiario.Columns(12).Width = 1200 ' nro cite
    grdBeneficiario.Columns(12).Caption = "Nro CITE"
    grdBeneficiario.Columns(13).Width = 1000 ' F. CITE
    grdBeneficiario.Columns(13).Caption = "F. CITE"
    grdBeneficiario.Columns(14).Width = 1000 ' Nro. con. hist
    grdBeneficiario.Columns(14).Caption = "Nro. consul. hist."
    grdBeneficiario.Columns(15).Width = 1000 ' fte. financ hist
    grdBeneficiario.Columns(15).Caption = "Fte. financ.hist."

End Sub

Private Sub pl_ControlaToolBar(Proceso As String)
    'TITULO:                Procedimiento ControlaBotones
    'PROPOSITO:             Permite controlar botones habilitando /deshabilitando para el tipo de proceso q siga

    On Error GoTo EtiqError
    
    If rs_grdPrincipal.RecordCount = 0 Then ' si no existen registros se cancelan todos los botones
        
        Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "GrupoLiq")
        Call pl_HabilitaUnaOpcion("Tool_Editar", False, "GrupoLiq")
        Call pl_HabilitaUnaOpcion("Tool_Anular", False, "GrupoLiq")
        
        Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "Liquida")
        Call pl_HabilitaUnaOpcion("Tool_Editar", False, "Liquida")
        Call pl_HabilitaUnaOpcion("Tool_Anular", False, "Liquida")
        Call pl_HabilitaUnaOpcion("Tool_Aprobar", False, "Liquida")
        Call pl_HabilitaUnaOpcion("Tool_Devengar", False, "Liquida")

        Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "Beneficiario")
        Call pl_HabilitaUnaOpcion("Tool_Monto", False, "Beneficiario")
        Call pl_HabilitaUnaOpcion("Tool_Anular", False, "Beneficiario")
        Call pl_HabilitaUnaOpcion("Tool_AnularTodo", False, "Beneficiario")
        Call pl_HabilitaUnaOpcion("Tool_Conformidad", False, "Beneficiario")
        Call pl_HabilitaUnaOpcion("Tool_Factura", False, "Beneficiario")

        Exit Sub
    End If
    
    Select Case Proceso
      
      ' ********************************************************
      ' controla botones de la: GRUPO DE PAGO
      ' ********************************************************
    
      Case "GrupoLiq" ' botones para procesos de la ficha GRUPO DE PAGO
        
        Select Case LTrim(RTrim(rs_grdPrincipal!ESTADO_APROBADO)) ' estado aprobación de consultoria
          Case "S", "A" ' estado de aprobación de la solicitud S=aprobado
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "GrupoLiq")
            Call pl_HabilitaUnaOpcion("Tool_Editar", False, "GrupoLiq")
            Call pl_HabilitaUnaOpcion("Tool_Anular", False, "GrupoLiq")
        
          Case Else ' "", Null = solo solicitado tramite no iniciado
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "GrupoLiq")
            Call pl_HabilitaUnaOpcion("Tool_Editar", True, "GrupoLiq")
            Call pl_HabilitaUnaOpcion("Tool_Anular", True, "GrupoLiq")
                    
        End Select
        
      ' ********************************************************
      ' controla botones de : LIQUIDACION
      ' ********************************************************
        
      Case "Liquida" ' botones para procesos de LIQUIDACION
        
        If rs_grdLiquida.RecordCount = 0 Then
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Editar", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Anular", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Aprobar", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Desaprobar", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Devengar", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_AnulaDevengar", False, "Liquida")
            cmdPagoDeven.Visible = True
            cmdPagoAnulaDev.Visible = False
            cmdPagoAprob.Visible = True
            cmdPagoDesaprob.Visible = False
            Exit Sub
        End If
        
        Select Case LTrim(RTrim(rs_grdLiquida!ESTADO_APROBADO)) ' estado aprobado
          Case "S"
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Editar", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Anular", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Aprobar", False, "Liquida")
            
            If LTrim(RTrim(rs_grdLiquida!estado_devengado)) <> "S" Then ' estado devengado
                Call pl_HabilitaUnaOpcion("Tool_Devengar", True, "Liquida")
                Call pl_HabilitaUnaOpcion("Tool_AnulaDevengar", False, "Liquida")
                Call pl_HabilitaUnaOpcion("Tool_Desaprobar", True, "Liquida")
                cmdPagoDeven.Visible = True
                cmdPagoAnulaDev.Visible = False
                cmdPagoAprob.Visible = False
                cmdPagoDesaprob.Visible = True
            Else
                Call pl_HabilitaUnaOpcion("Tool_Devengar", False, "Liquida")
                
                ' verifica si tiene devengado de pago aprobados
                SQLs = "select * from ac_ben_comprDeven where gp_ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and gp_codigo_unidad ='" & rs_grdPrincipal!codigo_unidad & "' and gp_codigo_grupo =" & rs_grdPrincipal!codigo_grupo & " and gp_numero_pago = " & rs_grdLiquida!NUMERO_PAGO & " and tipoComprobante ='DEV' and aprobotesoreria='S'"
                Set rstTemp = New ADODB.Recordset
                rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
                If rstTemp.RecordCount > 0 Then
                    Call pl_HabilitaUnaOpcion("Tool_AnulaDevengar", False, "Liquida")
                  Else
                    Call pl_HabilitaUnaOpcion("Tool_AnulaDevengar", True, "Liquida")
                End If
                
                Call pl_HabilitaUnaOpcion("Tool_Desaprobar", False, "Liquida")
                cmdPagoDeven.Visible = False
                cmdPagoAnulaDev.Visible = True
                cmdPagoAprob.Visible = True
                cmdPagoDesaprob.Visible = False
            End If
            
          Case Else
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Editar", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Anular", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Aprobar", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Desaprobar", False, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_Devengar", True, "Liquida")
            Call pl_HabilitaUnaOpcion("Tool_AnulaDevengar", False, "Liquida")
            cmdPagoDeven.Visible = True
            cmdPagoAnulaDev.Visible = False
            cmdPagoAprob.Visible = True
            cmdPagoDesaprob.Visible = False
        End Select
          
      ' ********************************************************
      ' controla botones de la ficha: BENEFICIARIO
      ' ********************************************************

      Case "Beneficiario" ' botones para procesos de la ficha BENEFICIARIO
        
        If rs_grdLiquida.RecordCount = 0 Then
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Monto", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Anular", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_AnularTodo", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Conformidad", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Factura", False, "Beneficiario")
            
            Exit Sub
        End If
        
        Select Case LTrim(RTrim(rs_grdLiquida!ESTADO_APROBADO)) ' estado
          Case "S"
            Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Monto", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Anular", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_AnularTodo", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Conformidad", False, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Factura", False, "Beneficiario")
            
          Case Else
            If rs_grdPrincipal!modalidad_pago = "I" And rs_grdBeneficiario.RecordCount > 0 Then
                Call pl_HabilitaUnaOpcion("Tool_Nuevo", False, "Beneficiario")
              Else
                Call pl_HabilitaUnaOpcion("Tool_Nuevo", True, "Beneficiario")
            End If
            Call pl_HabilitaUnaOpcion("Tool_Monto", True, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Anular", True, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_AnularTodo", True, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Conformidad", True, "Beneficiario")
            Call pl_HabilitaUnaOpcion("Tool_Factura", True, "Beneficiario")

        End Select
          
    End Select
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Function fl_VerificaConformidad() As Boolean
    'TITULO:                Función fl_VerificaConformidad
    'PROPOSITO:             Verifica los datos para registrar la conformidad
    'EJEMPLO DE LLAMADA:    fl_VerificaConformidad
    
    On Error GoTo EtiqError

    fl_VerificaConformidad = True ' asuminos que se cuenta con los datos mínimos para grabar

    
    ' verificamos si los montos son correstos
    rs_grdBeneficiario.MoveFirst
    While Not rs_grdBeneficiario.EOF
        If rs_grdBeneficiario!monto_us <= 0 Then
            MsgBox "El monto [" & rs_grdBeneficiario!monto_us & "] a liquidar correspondiente a [" & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & " " & rs_grdBeneficiario!Nombre & "] no es vàlido." & Chr(13) & "Corrija el error e intente registrar conformidad nuevamente.", vbInformation, "Aviso"
            Call grdBeneficiario_RowColChange(0, 0)
            cmdBenMonto.SetFocus ' se posiciona en el boton de editar
            fl_VerificaConformidad = False
            Exit Function
        End If
        rs_grdBeneficiario.MoveNext
    Wend

    rs_grdBeneficiario.MoveFirst
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Function

EtiqError:
    fl_VerificaConformidad = False
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Function

Private Function fl_VerificaAprobar() As Boolean
    'TITULO:                Función fl_VerificaAprobar
    'PROPOSITO:             Verifica los datos para aprobar la liquidación
    'EJEMPLO DE LLAMADA:    fl_VerificaAprobar
    
    On Error GoTo EtiqError
    
    fl_VerificaAprobar = True ' asuminos que se cuenta con los datos mnimos para grabar

    ' verificamos registro de conformidad
    If Not (fl_VerificaConformidad) Then
        fl_VerificaAprobar = False
            Exit Function
    End If
    
    ' verificamos si tiene registro de conformidad
    rs_grdBeneficiario.MoveFirst
    While Not rs_grdBeneficiario.EOF
        If rs_grdBeneficiario!estado_conformidad <> "S" Then
            MsgBox "La conformidad correspondiente a [" & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & " " & rs_grdBeneficiario!Nombre & "] no esta registrada." & Chr(13) & "Corrija el error e intente registrar aprobar nuevamente.", vbInformation, "Aviso"
            Call grdBeneficiario_RowColChange(0, 0)
            cmdBenConf.SetFocus ' se posiciona en el boton
            fl_VerificaAprobar = False
            Exit Function
        End If
        rs_grdBeneficiario.MoveNext
    Wend
    rs_grdBeneficiario.MoveFirst ' para no probar algun error de posicion

    ' verificamos si la modalidad de liquidación es coherente
    SQLs = "SELECT ges_gestion, codigo_unidad, codigo_grupo, numero_pago, count(codigo_beneficiario) as numbenf from ao_pagos_cronograma_detalle "
    SQLs = SQLs & "WHERE ges_gestion = '" & rs_grdPrincipal!ges_gestion & "' "
    SQLs = SQLs & " AND codigo_unidad = '" & rs_grdPrincipal!codigo_unidad & "' "
    SQLs = SQLs & " AND codigo_grupo = " & rs_grdPrincipal!codigo_grupo
    SQLs = SQLs & " AND estado_conformidad = 'S' "
    SQLs = SQLs & " AND estado_devengado = 'S' "
    SQLs = SQLs & " GROUP BY ges_gestion, codigo_unidad, codigo_grupo, numero_pago "
    SQLs = SQLs & " HAVING Count(codigo_beneficiario) > 1"
        
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 And rs_grdPrincipal!modalidad_pago = "I" Then
        MsgBox "La modalidad de liquidación [Planilla individual] correspondiente al grupo: [" & rs_grdPrincipal!codigo_grupo & "] [" & rs_grdPrincipal!descripcion_grupo & "], unidad: [" & rs_grdPrincipal!codigo_unidad & "] no es válida." & Chr(13) & "Corrija el error e intente aprobar nuevamente.", vbInformation, "Aviso"
        cmdEditar.SetFocus ' se posiciona en el boton de editar grupo
        fl_VerificaAprobar = False
        Exit Function
    End If
    
    ' verificamos si existe pendientes ordenes de liquidacion sin procesar aprobar
    SQLs = "SELECT 'MinPago' = MIN(numero_pago) FROM ao_pagos_cronograma "
    SQLs = SQLs & "WHERE (estado_aprobado ='N' or estado_devengado ='N') and ges_gestion = '" & rs_grdPrincipal!ges_gestion & "' "
    SQLs = SQLs & "AND codigo_unidad = '" & rs_grdPrincipal!codigo_unidad & "' "
    SQLs = SQLs & "AND codigo_grupo = " & rs_grdPrincipal!codigo_grupo
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        If rstTemp!MinPago < rs_grdLiquida!NUMERO_PAGO Then
            MsgBox "La Orden de Liquidación Nro:[" & rstTemp!MinPago & "] no fue procesada. Debe ser procesada antes de procesar una Liquidación posterior." & Chr(13) & "Corrija el error e intente procesar nuevamente nuevamente.", vbInformation, "Aviso"
            rs_grdLiquida.MoveFirst
            rs_grdLiquida.Find " numero_pago =" & rstTemp!MinPago
            Call grdLiquida_RowColChange(0, 0)
            grdLiquida.SetFocus ' se posiciona en el boton de conmformidad
            fl_VerificaAprobar = False
            Exit Function
        End If
      
    End If
    
    ' verifica si los beneficiarios cuentan con registro de emite o no factura
    rs_grdBeneficiario.MoveFirst
    While Not rs_grdBeneficiario.EOF
        If Len(Trim(rs_grdBeneficiario!emite_factura)) = 0 Then
            MsgBox "El beneficiario [" & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & " " & rs_grdBeneficiario!Nombre & "] NO cuanta con registro de CON/SIN RETENCION." & Chr(13) & "Corrija el error e intente aprobar nuevamente.", vbInformation, "Aviso"
            Call grdBeneficiario_RowColChange(0, 0)
            cmdBenFact.SetFocus ' se posiciona en el boton
            fl_VerificaAprobar = False
            Exit Function
        End If
        rs_grdBeneficiario.MoveNext
    Wend

    rs_grdBeneficiario.MoveFirst
    
    ' verificamos si existe registro de contrato si es asi verfiica si lasd fechas de inicio y fin estan registrados
    rs_grdBeneficiario.MoveFirst
    While Not rs_grdBeneficiario.EOF
        SQLs = "SELECT ao_contrato_c.fechas_confirmado FROM ao_adjudica_c LEFT OUTER JOIN ao_contrato_c ON ao_adjudica_c.ges_gestion = ao_contrato_c.ges_gestion AND ao_adjudica_c.codigo_unidad = ao_contrato_c.codigo_unidad AND "
        SQLs = SQLs & "ao_adjudica_c.codigo_solicitud = ao_contrato_c.codigo_solicitud AND ao_adjudica_c.numero_consultoria = ao_contrato_c.numero_consultoria AND ao_adjudica_c.codigo_beneficiario = ao_contrato_c.codigo_beneficiario "
        SQLs = SQLs & "WHERE ao_adjudica_c.gp_ges_gestion = '" & lblGestion.Caption & "' AND ao_adjudica_c.gp_codigo_unidad = '" & lblCodUniSol.Caption & "' AND ao_adjudica_c.gp_codigo_grupo = " & Val(lblCodGrupo.Caption) & " AND ao_adjudica_c.codigo_beneficiario = '" & rs_grdBeneficiario!codigo_beneficiario & "'"
        Set rstTemp = New ADODB.Recordset
        rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
        If rstTemp!fechas_confirmado <> "S" Then
            MsgBox "La fechas de inicio y fin de contrato de [" & rs_grdBeneficiario!paterno & " " & rs_grdBeneficiario!materno & " " & rs_grdBeneficiario!Nombre & "] no se encuentra confirmado." & Chr(13) & "Corrija el error e intente registrar procesar nuevamente.", vbInformation, "Aviso"
            Call grdBeneficiario_RowColChange(0, 0)
            cmdDatosContrato.SetFocus ' se posiciona en el boton
            fl_VerificaAprobar = False
            Exit Function
        End If
        rs_grdBeneficiario.MoveNext
    Wend
    rs_grdBeneficiario.MoveFirst ' para no probar algun error de posicion
  
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Function

EtiqError:
    fl_VerificaAprobar = False
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Function

Private Function fl_VerificaDevengar() As Boolean
    'TITULO:                Función fl_VerificaDevengar
    'PROPOSITO:             Verifica los datos para devengar la liquidación
    'EJEMPLO DE LLAMADA:    fl_VerificaDevengar
    
    On Error GoTo EtiqError
    
    fl_VerificaDevengar = True ' asuminos que se cuenta con los datos mnimos para grabar

    ' verificamos registro de conformidad
    If Not (fl_VerificaConformidad) Then
        fl_VerificaDevengar = False
        Exit Function
    End If
    
    ' verificamos aprobación de liquidación
    If rs_grdLiquida!ESTADO_APROBADO <> "S" Then
        MsgBox "La liquidación Nro.: [" & rs_grdLiquida!NUMERO_PAGO & "] [" & rs_grdLiquida!Concepto & "] no se encuentra aprobado." & Chr(13) & "Corrija el error e intente devengar nuevamente.", vbInformation, "Aviso"
        cmdPagoAprob.SetFocus ' se posiciona en el boton
        fl_VerificaDevengar = False
        Exit Function
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Function

EtiqError:
    fl_VerificaDevengar = False
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Function

Private Sub pl_GeneraDevengado()
    'TITULO:                Procedimiento pl_GeneraDevengado
    'PROPOSITO:             Genera el devengado por beneficiario pasando por la tabla temporal
    'EJEMPLO DE LLAMADA:    call pl_GeneraDevengado
    
Dim Respuesta As String
Dim montocontrol As Double
Dim sesion$
Dim Error As Integer
Dim SeAsigno As Boolean

Dim rstOrden As New ADODB.Recordset
Dim rsTMP As New ADODB.Recordset
Dim rsc As New ADODB.Recordset


    On Error GoTo EtiqError ' activamos el manejador de errores
    Screen.MousePointer = vbHourglass
Error = 0

sesion = Left("S" & CStr(Rnd()), 10)

Set rsTMP = New ADODB.Recordset
rsTMP.Open "select * from ac_Ben_Devengado_TMP where sesion='" & sesion & "'", db, adOpenDynamic, adLockOptimistic

Do While rsTMP.RecordCount > 0
    sesion = Left("S" & CStr(Rnd()), 10)
    Set rsTMP = New ADODB.Recordset
    rsTMP.Open "select * from ac_Ben_Devengado_TMP where sesion='" & sesion & "'", db, adOpenDynamic, adLockOptimistic
Loop

'recorre todos los beneficiarios para generar su devengado en tabla temporal
rs_grdBeneficiario.MoveFirst
Do While Not rs_grdBeneficiario.EOF
    If rs_grdBeneficiario!estado_conformidad = "S" Then
        'JQ QR
        'DE.dbo_apGeneralSearching "update ac_ben_comprdeven set monto_dolares_acum = 0 where gp_ges_gestion = '" & lblGestion.Caption & "' and gp_codigo_unidad ='" & lblCodUniSol.Caption & "' and gp_codigo_grupo = " & Val(lblCodGrupo.Caption) & " and codigo_beneficiario = '" & rs_grdBeneficiario!codigo_beneficiario & "'"
        
        SQLs = "select MIN(ordencomprobante) AS ordencomprobante From AC_BEN_COMPRDEVEN WHERE MONTO_BOLIVIANOS >0 AND gp_ges_gestion = '" & lblGestion.Caption & "' and gp_codigo_unidad ='" & lblCodUniSol.Caption & "' and gp_codigo_grupo = " & Val(lblCodGrupo.Caption) & " AND codigo_beneficiario = '" & rs_grdBeneficiario!codigo_beneficiario & "' and aprobotesoreria   = 'S' and tipocomprobante = 'COM' order by ordencomprobante"
''        SQLs = "select distinct ordencomprobante From AC_BEN_COMPRDEVEN WHERE gp_ges_gestion = '" & lblGestion.Caption & "' and gp_codigo_unidad ='" & lblCodUniSol.Caption & "' and gp_codigo_grupo = " & Val(lblCodGrupo.Caption) & " AND codigo_beneficiario = '" & rs_grdBeneficiario!codigo_beneficiario & "' and aprobotesoreria   = 'S' and tipocomprobante = 'COM' order by ordencomprobante"
        Set rstOrden = New ADODB.Recordset
        rstOrden.Open SQLs, db, adOpenStatic, adLockReadOnly
        
        If rstOrden.RecordCount = 0 Then
            Error = 1
        Else
             SeAsigno = False
             Do While Not rstOrden.EOF
                 If fl_ExisteSaldoParaDevengar(rs_grdBeneficiario!codigo_beneficiario, rstOrden!ordenComprobante) Then
                     If fl_ExisteEspacioEnComprometido(sesion, rstOrden!ordenComprobante) Then
                         'devenga lo que se pueda del comprometido --> TMP
                         ' comprueba si es solo porcentaje 100%
                         'JQ QR
                         'DE.dbo_ap_ComSolo100y258o222 rs_grdPrincipal!ges_gestion, rs_grdPrincipal!codigo_unidad, rs_grdPrincipal!codigo_grupo, rs_grdBeneficiario!codigo_beneficiario, rstOrden!ordenComprobante, Respuesta
                         If Respuesta = "S" Then
                            'JQ QR
                            'DE.dbo_ap_GeneraDevEnTmp100 sesion, rs_grdPrincipal!ges_gestion, rs_grdBeneficiario!codigo_beneficiario, rstOrden!ordenComprobante, rs_grdPrincipal!codigo_unidad, rs_grdPrincipal!codigo_grupo, rs_grdLiquida!NUMERO_PAGO, rs_grdBeneficiario!emite_factura
                         Else
                            'JQ QR
                            'DE.dbo_ap_GeneraDevengadoEnTmp sesion, rs_grdPrincipal!ges_gestion, rs_grdBeneficiario!codigo_beneficiario, rstOrden!ordenComprobante, rs_grdPrincipal!codigo_unidad, rs_grdPrincipal!codigo_grupo, rs_grdLiquida!NUMERO_PAGO, rs_grdBeneficiario!emite_factura
                         End If
                         
                         SeAsigno = True
                     End If
                 End If
                 rstOrden.MoveNext
             Loop
             
             If SeAsigno = False Then
                 Error = 2
             End If
         End If
    
    End If
    'toma siguiente beneficiario
    rs_grdBeneficiario.MoveNext
    If Error > 0 Then Exit Do
Loop

If Error = 0 Then ' sin errores
    'este SP genera el devengado en base a la tabla ac_ben_devengado_tmp teniendo como agrupador la sesión
    'JQ QR
    'DE.dbo_ap_GeneraDevengado sesion, rs_grdPrincipal!ges_gestion, rs_grdPrincipal!codigo_unidad, rs_grdPrincipal!codigo_grupo, rs_grdLiquida!NUMERO_PAGO, GlUsuario

ElseIf Error = 1 Then
    MsgBox "Existe conformidad de parte de la Unidad, pero el Compromiso de Pago no está aprobado", vbCritical, "saf2002"
ElseIf Error = 2 Then
    MsgBox "No se genero Devengado por que existe error en los saldos del Compromiso de Pago, revise por favor", vbCritical, "SAF"
End If

'elimina registros de la sesion en la tabla temporal
'JQ QR
'DE.dbo_apGeneralSearching "delete ac_ben_devengado_tmp where sesion='" & sesion & "'"
    
    Screen.MousePointer = vbDefault
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Function fl_ExisteSaldoParaDevengar(codBen As String, ordenComprobante As Integer) As Boolean

    'verifica si existe saldo del comprometido para devengar
    fl_ExisteSaldoParaDevengar = False
    
    SQLs = "select saldo_US = SUM(monto_dolares), saldo_BS = SUM(monto_bolivianos) from ac_ben_comprdeven where codigo_beneficiario = '" & codBen & "' and tipocomprobante = 'COM' and aprobotesoreria='S' AND GP_GES_GESTION='" & rs_grdPrincipal!ges_gestion & "' AND gp_codigo_unidad='" & rs_grdPrincipal!codigo_unidad & "' and  GP_CODIGO_GRUPO=" & rs_grdPrincipal!codigo_grupo
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    If Not rstTemp.EOF Then
        Do While Not rstTemp.EOF
            If rs_grdBeneficiario!tipo_moneda = "$US" And rstTemp!saldo_US - rs_grdBeneficiario!monto_us >= 0 Then
                fl_ExisteSaldoParaDevengar = True
            End If
            If rs_grdBeneficiario!tipo_moneda = "Bs" And rstTemp!saldo_BS - rs_grdBeneficiario!monto_bs >= 0 Then
                fl_ExisteSaldoParaDevengar = True
            End If
            rstTemp.MoveNext
        Loop
    End If
    rstTemp.Close
    
    
''''    'verifica si existe saldo del comprometido para devengar
''''    fl_ExisteSaldoParaDevengar = False
''''
''''''    SQLs = "select saldo = monto_dolares - monto_dolares_acum from ac_ben_comprdeven where codigo_beneficiario = '" & codBen & "' and ordencomprobante=" & ordenComprobante & " and tipocomprobante = 'COM' and aprobotesoreria='S' AND GP_GES_GESTION='" & rs_grdPrincipal!ges_gestion & "' AND gp_codigo_unidad='" & rs_grdPrincipal!codigo_unidad & "' and  GP_CODIGO_GRUPO=" & rs_grdPrincipal!codigo_grupo
''''    SQLs = "select saldo_US = SUM(monto_dolares), saldo_BS = SUM(monto_bolivianos) from ac_ben_comprdeven where codigo_beneficiario = '" & codBen & "' and tipocomprobante = 'COM' and aprobotesoreria='S' AND GP_GES_GESTION='" & rs_grdPrincipal!ges_gestion & "' AND gp_codigo_unidad='" & rs_grdPrincipal!codigo_unidad & "' and  GP_CODIGO_GRUPO=" & rs_grdPrincipal!codigo_grupo
''''    Set rstTemp = New ADODB.Recordset
''''    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
''''
''''    If Not rstTemp.EOF Then
''''        Do While Not rstTemp.EOF
''''            If rs_grdBeneficiario!tipo_moneda = "$US" And rstTemp!saldo_US - rs_grdBeneficiario!monto_US > 0 Then
''''            If rstTemp!saldo > 0 Then fl_ExisteSaldoParaDevengar = True
''''            rstTemp.MoveNext
''''        Loop
''''    End If
''''    rstTemp.Close
    
End Function

Function fl_ExisteEspacioEnComprometido(sesion As String, ordenComprobante As Integer) As Boolean
    'determina si existe espacio en el comprometido considerando los devengados benerados en la tabla temporal
    
    Dim rsc As New ADODB.Recordset
    Dim rsR As New ADODB.Recordset
    Dim rsT As New ADODB.Recordset
    Dim montoTMP As Double
    
    On Error GoTo EtiqError ' activamos el manejador de errores

    fl_ExisteEspacioEnComprometido = True
    
    rsR.Open "SELECT ges_gestion, org_codigo, codigo_pago " & _
             "From AC_BEN_COMPRDEVEN WHERE   codigo_beneficiario = '" & rs_grdBeneficiario!codigo_beneficiario & "' and " & _
                                            "aprobotesoreria   = 'S'         and " & _
                                            "tipocomprobante   = 'COM'       and " & _
                                            "ordencomprobante  = " & ordenComprobante & " and " & _
                                            "gp_ges_Gestion    = '" & rs_grdPrincipal!ges_gestion & "' and " & _
                                            "gp_codigo_unidad  = '" & rs_grdPrincipal!codigo_unidad & "' and " & _
                                            "gp_codigo_grupo   = " & rs_grdPrincipal!codigo_grupo & " " & _
             "order by ges_gestion, org_codigo desc, codigo_pago", db, adOpenStatic, adLockReadOnly
    
    If rsR.RecordCount > 0 Then
        rsc.Open "SELECT monto_Dolares " & _
             "From pagos  WHERE     ges_Gestion = '" & rsR!ges_gestion & "' and " & _
                                     "org_codigo = '" & rsR!org_codigo & "' and " & _
                                     "codigo_pago = " & rsR!codigo_pago & " and " & _
                                     "tipo_formulario = 'COM' ", db, adOpenStatic, adLockReadOnly
        If rsc.EOF Then
            fl_ExisteEspacioEnComprometido = False
        Else
            rsT.Open "SELECT monto_dolares From ac_ben_devengado_TMP " & _
                     "WHERE sesion       = '" & sesion & "' and " & _
                           "Cges_gestion = '" & rsR!ges_gestion & "' and " & _
                           "Corg_codigo  = '" & rsR!org_codigo & "' and " & _
                           "Ccodigo_pago = " & rsR!codigo_pago, db, adOpenStatic, adLockReadOnly
            If rsT.RecordCount > 0 Then
                montoTMP = IIf(IsNull(rsT!monto_dolares), 0, rsT!monto_dolares)
            Else
                montoTMP = 0
            End If
            If rsc!monto_dolares - montoTMP > 0 Then
                fl_ExisteEspacioEnComprometido = True
            Else
                fl_ExisteEspacioEnComprometido = False
            End If
        End If
    Else
        fl_ExisteEspacioEnComprometido = False
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Function

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"
    
End Function

Private Function fl_VerificaAnulaOrdLiq() As Boolean
    'TITULO:                Función fl_VerificaAnulaOrdLiq
    'PROPOSITO:             Verifica los datos para procesar la elimnacion
    'EJEMPLO DE LLAMADA:    fl_VerificaAnulaOrdLiq
    Dim rstTemp As ADODB.Recordset ' usado para la carga de los combos de base
    
    fl_VerificaAnulaOrdLiq = True
    
    ' verificamos si tiene algun devengado generado
    SQLs = "select * from ao_pagos_cronograma where ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and codigo_unidad='" & rs_grdPrincipal!codigo_unidad & "' and codigo_grupo=" & rs_grdPrincipal!codigo_grupo & " and numero_pago = " & rs_grdLiquida!NUMERO_PAGO & " and estado_devengado ='S'"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    If rstTemp.RecordCount = 0 Then
        MsgBox "No tiene Orden de Liquidación generada para el Nro. de liquidación [" & rs_grdLiquida!NUMERO_PAGO & "].", vbInformation, "Aviso"
        grdLiquida.SetFocus
        fl_VerificaAnulaOrdLiq = False
        Exit Function
    End If
    
    ' verificamos si existe un numero de liquidación mayor con orden de liquidación
    SQLs = "select * from ao_pagos_cronograma where ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and codigo_unidad='" & rs_grdPrincipal!codigo_unidad & "' and codigo_grupo=" & rs_grdPrincipal!codigo_grupo & " and numero_pago > " & rs_grdLiquida!NUMERO_PAGO & " and estado_aprobado in('N','S')"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    
    If rstTemp.RecordCount > 0 Then
        MsgBox "No puede anular la Orden de Liquidación Nro.[" & rs_grdLiquida!NUMERO_PAGO & "] por existir registro de liquidaciones posteriores.", vbInformation, "Aviso"
        grdLiquida.SetFocus
        fl_VerificaAnulaOrdLiq = False
        Exit Function
    End If
    
    ' verifica si tiene devengado de pago aprobados
    SQLs = "select * from ac_ben_comprDeven where gp_ges_gestion='" & rs_grdPrincipal!ges_gestion & "' and gp_codigo_unidad ='" & rs_grdPrincipal!codigo_unidad & "' and gp_codigo_grupo =" & rs_grdPrincipal!codigo_grupo & " and gp_numero_pago = " & rs_grdLiquida!NUMERO_PAGO & " and tipoComprobante ='DEV' and aprobotesoreria='S'"
    Set rstTemp = New ADODB.Recordset
    rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
        MsgBox "No puede anular la Orden de Liquidación correspondiente a [" & lblCodGrupo.Caption & "][" & lblDesGrupo.Caption & "] Nro. Liquidación: [" & rs_grdLiquida!NUMERO_PAGO & "] de liquidación por tener devengado APROBADO." & Chr(13) & "Comuniquese con el administrador del sistema.", vbInformation, "Aviso"
        grdLiquida.SetFocus
        fl_VerificaAnulaOrdLiq = False
        Exit Function
    End If
    
    Set rstTemp = Nothing
    
End Function


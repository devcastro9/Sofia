VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_ao_solicitud_rrhh 
   BackColor       =   &H00000000&
   Caption         =   "Administración de Personal - Contratación - Solicitud de Contratación"
   ClientHeight    =   10260
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "frm_ao_solicitud_rrhh.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10260
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.CommandButton BtnImprimir2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Solicitud"
      Height          =   640
      Left            =   1120
      Picture         =   "frm_ao_solicitud_rrhh.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Imprime Solicitud de Contratación"
      Top             =   6210
      Width           =   765
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1490
      Left            =   120
      Picture         =   "frm_ao_solicitud_rrhh.frx":2184
      ScaleHeight     =   1425
      ScaleWidth      =   1875
      TabIndex        =   70
      Top             =   7060
      Width           =   1935
      Begin VB.CommandButton BtnAnlDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   120
         Picture         =   "frm_ao_solicitud_rrhh.frx":6E1B6
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Elimina Detalle Elegido"
         Top             =   740
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   980
         Picture         =   "frm_ao_solicitud_rrhh.frx":6E5F8
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   120
         Picture         =   "frm_ao_solicitud_rrhh.frx":6EA3A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Adiciona Detalle"
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir1 
         BackColor       =   &H80000018&
         Caption         =   "Bitácora"
         Height          =   640
         Left            =   980
         Picture         =   "frm_ao_solicitud_rrhh.frx":6EE7C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Imprime Bitácora de RR.HH."
         Top             =   740
         Width           =   765
      End
   End
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1490
      Left            =   120
      Picture         =   "frm_ao_solicitud_rrhh.frx":705FE
      ScaleHeight     =   1425
      ScaleWidth      =   1875
      TabIndex        =   69
      Top             =   5445
      Width           =   1935
      Begin VB.CommandButton BtnAddDetalle2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   120
         Picture         =   "frm_ao_solicitud_rrhh.frx":DC630
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Adiciona Detalle"
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   980
         Picture         =   "frm_ao_solicitud_rrhh.frx":DCA72
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton BtnAnlDetalle2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   120
         Picture         =   "frm_ao_solicitud_rrhh.frx":DCEB4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Elimina Detalle Elegido"
         Top             =   730
         Width           =   765
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "frm_ao_solicitud_rrhh.frx":DD2F6
      ScaleHeight     =   960
      ScaleWidth      =   14880
      TabIndex        =   65
      Top             =   120
      Width           =   14940
      Begin VB.CommandButton CmdFoto 
         BackColor       =   &H00808000&
         Caption         =   "&Digitaliza"
         Height          =   720
         Left            =   5160
         Picture         =   "frm_ao_solicitud_rrhh.frx":149328
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Carga Foto de la Persona"
         Top             =   120
         Visible         =   0   'False
         Width           =   740
      End
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "frm_ao_solicitud_rrhh.frx":1498B2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "frm_ao_solicitud_rrhh.frx":149ABC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "frm_ao_solicitud_rrhh.frx":14A0E0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "frm_ao_solicitud_rrhh.frx":14A6C0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   6000
         Picture         =   "frm_ao_solicitud_rrhh.frx":14B38A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4320
         Picture         =   "frm_ao_solicitud_rrhh.frx":14B594
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprime Lista de Solicitudes"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "frm_ao_solicitud_rrhh.frx":14BB51
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdDesapr 
         BackColor       =   &H0080C0FF&
         Caption         =   "Desapr"
         Height          =   720
         Left            =   2640
         Picture         =   "frm_ao_solicitud_rrhh.frx":14C109
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   740
      End
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Docs."
         Enabled         =   0   'False
         Height          =   720
         Left            =   5160
         Picture         =   "frm_ao_solicitud_rrhh.frx":14C313
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   120
         Width           =   740
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   10530
         TabIndex        =   68
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_ao_solicitud_rrhh.frx":14C69B
      ScaleHeight     =   915
      ScaleWidth      =   14880
      TabIndex        =   61
      Top             =   120
      Width           =   14940
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "frm_ao_solicitud_rrhh.frx":1B86CD
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "frm_ao_solicitud_rrhh.frx":1B88D7
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   9900
         TabIndex        =   64
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00000000&
      Caption         =   "BITÁCORA DE "
      ForeColor       =   &H00FFFFC0&
      Height          =   1575
      Left            =   2125
      TabIndex        =   58
      Top             =   6980
      Width           =   12900
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8AE1
         Height          =   1215
         Left            =   195
         TabIndex        =   59
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "bitacora_codigo"
            Caption         =   "Correl"
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
            DataField       =   "negocia_forma"
            Caption         =   "Tipo.Evento"
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
            DataField       =   "negocia_fecha_real"
            Caption         =   "Fecha Evento"
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
            DataField       =   "negocia_hora_real"
            Caption         =   "Hora Evento"
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
            DataField       =   "negocia_gasto_estimado"
            Caption         =   "Gasto Estimado"
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
         BeginProperty Column05 
            DataField       =   "beneficiario_codigo"
            Caption         =   "Contacto con:"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Personal CGI"
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
            DataField       =   "negocia_tarea_realizada"
            Caption         =   "Tema Tratado"
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
            DataField       =   "negocia_observaciones"
            Caption         =   "Conclusiones u Observaciones"
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
               Locked          =   -1  'True
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   3644.788
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               ColumnWidth     =   2294.929
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00000000&
      Caption         =   "PUESTOS CONVOCADOS PARA POSTULANTES"
      ForeColor       =   &H00FFFFC0&
      Height          =   1575
      Left            =   2145
      TabIndex        =   38
      Top             =   5350
      Width           =   12900
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8AFC
         Height          =   1215
         Left            =   195
         TabIndex        =   39
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "benef_id"
            Caption         =   "Correl"
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
            DataField       =   "puesto_codigo"
            Caption         =   "Puesto"
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
            DataField       =   "observaciones"
            Caption         =   "Descripcion del Puesto"
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
            DataField       =   "ocup_descripcion"
            Caption         =   "Perfil Profesional"
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
            DataField       =   "nivel_educ_descripcion"
            Caption         =   "Nivel Educacional"
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
         BeginProperty Column05 
            DataField       =   "munic_descripcion"
            Caption         =   "Lugar.Convocatoria"
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
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   3119.811
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   2145.26
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   2835.213
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1934.929
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   599.811
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFC0&
      Height          =   4080
      Left            =   120
      TabIndex        =   32
      Top             =   1200
      Width           =   5960
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   3330
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   5874
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "solicitud_codigo"
            Caption         =   "Nro.Trámite"
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
            Caption         =   "Unidad Ejec."
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "CI.Solicitante"
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
            DataField       =   "solicitud_fecha_solicitud"
            Caption         =   "Fecha.Solicitud"
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
         BeginProperty Column05 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFC0&
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
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   1320
         TabIndex        =   53
         Top             =   3700
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFC0&
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
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   3600
         TabIndex        =   54
         Top             =   3700
         Width           =   915
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   3640
         Width           =   5760
         _ExtentX        =   10160
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
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00000000&
      Height          =   4080
      Left            =   6160
      TabIndex        =   29
      Top             =   1200
      Width           =   8895
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   8410
         TabIndex        =   75
         Top             =   3250
         Width           =   290
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   5520
         TabIndex        =   57
         Top             =   490
         Width           =   290
      End
      Begin VB.TextBox txt_obs 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_observaciones"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   2520
         Visible         =   0   'False
         Width           =   1605
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8B17
         DataField       =   "unidad_codigo_sol"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   1860
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo11 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8B30
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7080
         TabIndex        =   45
         Top             =   960
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux1 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8B4A
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6240
         TabIndex        =   44
         Top             =   960
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8B64
         DataField       =   "solicitud_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3120
         TabIndex        =   35
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "solicitud_tipo"
         BoundColumn     =   "solicitud_tipo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8B7D
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3060
         TabIndex        =   5
         Top             =   3240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8B97
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7200
         TabIndex        =   33
         Top             =   1560
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox Txt_descripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_justificacion"
         DataSource      =   "Ado_datos"
         Height          =   555
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2475
         Width           =   6930
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8BB0
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4720
         TabIndex        =   3
         Top             =   1860
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8BC9
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4680
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8BE2
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1575
         TabIndex        =   21
         Top             =   480
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8BFB
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   41
         Top             =   3240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc11 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8C15
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4720
         TabIndex        =   1
         Top             =   1230
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8C2F
         DataField       =   "solicitud_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   1230
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "solicitud_tipo_descripcion"
         BoundColumn     =   "solicitud_tipo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "frm_ao_solicitud_rrhh.frx":1B8C48
         DataField       =   "unidad_codigo_sol"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3000
         TabIndex        =   72
         Top             =   1560
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   7200
         TabIndex        =   73
         Top             =   2720
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         Format          =   109510657
         CurrentDate     =   42255
         MaxDate         =   55153
         MinDate         =   2
      End
      Begin VB.Label lbl_campo2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tipo de Solicitud"
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
         Left            =   180
         TabIndex        =   76
         Top             =   975
         Width           =   1515
      End
      Begin VB.Label lbl_campo3 
         BackColor       =   &H00000000&
         Caption         =   "Unidad Solicitante"
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
         Left            =   180
         TabIndex        =   71
         Top             =   1600
         Width           =   1815
      End
      Begin VB.Label dtc_codigo9 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2040
         TabIndex        =   60
         Top             =   3645
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cite del Trámite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   6
         Left            =   6150
         TabIndex        =   56
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label Txt_campo2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "RRHH"
         DataField       =   "unidad_codigo_ant"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   6020
         TabIndex        =   55
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lbl_campo4 
         BackColor       =   &H00000000&
         Caption         =   "Responsable de la Solicitud"
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
         Left            =   4740
         TabIndex        =   52
         Top             =   1605
         Width           =   2895
      End
      Begin VB.Label lbl_descripcion 
         BackColor       =   &H00000000&
         Caption         =   "Justificación de la Solicitud:"
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
         Left            =   180
         TabIndex        =   51
         Top             =   2230
         Width           =   2565
      End
      Begin VB.Label lbl_campo10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Actividad del POA"
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
         Left            =   180
         TabIndex        =   50
         Top             =   3250
         Width           =   1635
      End
      Begin VB.Label lbl_campo9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Código de Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Left            =   180
         TabIndex        =   49
         Top             =   3660
         Width           =   1755
      End
      Begin VB.Label lbl_campo11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Dirigido a (RR.HH.):"
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
         Left            =   4740
         TabIndex        =   48
         Top             =   975
         Width           =   1770
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Unidad Ejecutora"
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
         Left            =   1605
         TabIndex        =   47
         Top             =   225
         Width           =   1560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   8880
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   8880
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   180
         TabIndex        =   43
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   7120
         TabIndex        =   42
         Top             =   3640
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro. de Documento Respaldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   13
         Left            =   4320
         TabIndex        =   37
         Top             =   3660
         Width           =   2730
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Fecha Solicitud"
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
         Index           =   12
         Left            =   7170
         TabIndex        =   36
         Top             =   2400
         Width           =   1545
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "REG"
         DataField       =   "estado_codigo"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   7905
         TabIndex        =   22
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro. Trámite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   31
         Top             =   225
         Width           =   1110
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   2
         Left            =   7995
         TabIndex        =   30
         Top             =   225
         Width           =   645
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   11280
      TabIndex        =   23
      Top             =   10260
      Width           =   11280
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   28
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   8640
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
   Begin Crystal.CrystalReport CR01 
      Left            =   7200
      Top             =   11040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2280
      Top             =   8640
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
   Begin MSAdodcLib.Adodc Ado_datos3 
      Height          =   330
      Left            =   4440
      Top             =   8640
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6720
      Top             =   8640
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   9000
      Top             =   8640
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   11280
      Top             =   8640
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
      Left            =   13560
      Top             =   8640
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
      Left            =   120
      Top             =   9000
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Left            =   2280
      Top             =   9000
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   4440
      Top             =   9000
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
      Caption         =   "Ado_datos10"
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
   Begin MSAdodcLib.Adodc Ado_detalle1 
      Height          =   330
      Left            =   120
      Top             =   11040
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
      Caption         =   "Ado_detalle1"
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
   Begin MSAdodcLib.Adodc Ado_detalle2 
      Height          =   330
      Left            =   2400
      Top             =   11040
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
      Caption         =   "Ado_detalle2"
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
      Left            =   6720
      Top             =   9000
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
   Begin MSAdodcLib.Adodc Ado_datos12 
      Height          =   330
      Left            =   9000
      Top             =   9000
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
      Caption         =   "Ado_datos12"
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
   Begin Crystal.CrystalReport CR02 
      Left            =   11400
      Top             =   8880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CR03 
      Left            =   11880
      Top             =   8880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm_ao_solicitud_rrhh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim queryinicial As String

Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim CodBien As String
Dim VAR_UNI, VAR_UNI2 As String
Dim sino As String
Dim parametro As String
Dim VAR_AUX, VAR_CONT2 As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAddDetalle_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  'If rs_datos!estado_codigo = "REG" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False
    'Call ABRIR_TABLA_DET
    frm_ao_solicitud_bitacora.txt_codigo.Caption = Me.txt_codigo.Caption
    frm_ao_solicitud_bitacora.Txt_campo1.Caption = Me.dtc_codigo1.Text
    frm_ao_solicitud_bitacora.Txt_descripcion.Caption = Me.dtc_desc1.Text
    frm_ao_solicitud_bitacora.Txt_Correl.Caption = 0    'rs_datos!correl_bitacora + 1
    frm_ao_solicitud_bitacora.DTPfecha1.Value = Date
    frm_ao_solicitud_bitacora.Txt_estado.Caption = "REG"
    frm_ao_solicitud_bitacora.lbl_bitacora.Caption = Me.FraDet1.Caption
    Ado_detalle1.Recordset.AddNew
    frm_ao_solicitud_bitacora.Show vbModal
    
    'Call ABRIR_TABLA_DET
    
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    'Fra_datos.Enabled = True
  'Else
  '  MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  'End If
End Sub

Private Sub BtnAddDetalle2_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos!estado_codigo = "REG" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
'    FraDet3.Enabled = False
'    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False
'    Select Case dtc_codigo2.Text
'        Case "1"    'SOLO COMPRAS BB y SS
'        Case "2"    'SOLO VENTA DE BIENES
'        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
'
'
'        Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'            Call ABRIR_TABLA_DET
'            Ado_detalle1.Recordset.AddNew
'            frm_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes.txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
''            'aw_p_ao_solicitud_edificacion.Txt_Correl.Caption = 0
''            aw_p_ao_solicitud_edificacion.dtc_codigo1.Text = Me.dtc_codigo3.Text
''            aw_p_ao_solicitud_edificacion.dtc_desc1.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
''            aw_p_ao_solicitud_edificacion.dtc_aux1.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
''            aw_p_ao_solicitud_edificacion.dtc_aux2.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
''            aw_p_ao_solicitud_edificacion.dtc_aux3.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
'            frm_solicitud_bienes.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes.Show vbModal
'        Case "5"    ' SERVICIO MODERNIZACION
'
'        Case "11"    ' CONTRATACION DE PERSONAL - RRHH
            'Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            frm_ao_solicitud_persona.txtFecha = Date
            frm_ao_solicitud_persona.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_ao_solicitud_persona.Txt_campo1.Text = Me.dtc_codigo1.Text
            frm_ao_solicitud_persona.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_ao_solicitud_persona.txtFecha.Value = DTPfecha1.Value
            frm_ao_solicitud_persona.txtEstado.Text = "REG"
            frm_ao_solicitud_persona.Show vbModal
'    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
'    FraDet3.Enabled = True
'    FrmABMDet3.Enabled = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnAnlDetalle_Click()
  If Ado_detalle1.Recordset.RecordCount > 0 Then
   sino = MsgBox("Está Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_detalle1.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_detalle1.Recordset.Delete 'adAffectAll
'        Ado_detalle1.Recordset("estado_codigo") = "ERR"
'        Ado_detalle1.Recordset("fecha_registro") = Date
'        Ado_detalle1.Recordset("usr_codigo") = GlUsuario
'        Ado_detalle1.Recordset("campo1") = "REG. ANULADO"
'        Ado_detalle1.Recordset.Update  'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR, un registro Aprobado o Anulado ...", vbExclamation, "Validación de Registro"
   End If
 Else
     MsgBox "No se puede ANULAR, el registro no fue identificado correctamente ...", vbExclamation, "Validación de Registro"
 End If
End Sub

Private Sub BtnAnlDetalle2_Click()
  If Ado_detalle2.Recordset.RecordCount > 0 Then
    If ExisteReg(Ado_datos.Recordset!unidad_codigo, Ado_datos.Recordset!solicitud_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ...", vbInformation + vbOKOnly, "Atención": Exit Sub
    sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + Str(Ado_detalle2.Recordset!benef_id), vbYesNo + vbQuestion, "Atención")
    If Ado_detalle2.Recordset("estado_codigo") = "REG" Then
       If sino = vbYes Then
         db.Execute "delete ao_solicitud_persona Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and benef_id = " & Ado_detalle2.Recordset!benef_id & " "
         Call ABRIR_TABLA_DET
       End If
    Else
         MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
    End If
  Else
     MsgBox "No se puede ANULAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
  End If
End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
   If Ado_datos.Recordset!unidad_codigo = "" Or Ado_datos.Recordset!beneficiario_codigo = "" Then
        MsgBox "No se puede APROBAR, debe registrar al " + lbl_campo4.Caption, vbExclamation, "Validación de Registro"
        Exit Sub
   End If
   'If rs_datos!estado_codigo = "REG" And Ado_datos.Recordset!correl_edificacion > 0 Then
   If rs_datos!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        Select Case dtc_codigo2.Text
            Case "1"    'SOLO COMPRAS BB y SS
            Case "2"    'SOLO VENTA DE BIENES
            Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
                Set rs_aux1 = New ADODB.Recordset
                'SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_detalle1.Recordset!edif_codigo & "'  "
                SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   "
                rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                'If rs_aux1.RecordCount > 0 Then
                '    MsgBox "El código ya existe, consulte con el administrador del Sistema..."
                '    var_cod = 0
                '    Exit Sub
                'Else
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    'rs_aux2.Open "Select max(trafico_codigo) as Codigo from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    rs_aux2.Open "Select max(trafico_codigo) as Codigo from ao_solicitud_calculo_trafico where unidad_codigo = '" & parametro & "' ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    End If
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    rs_aux2.Open "Select edif_capacidad_min_trafico as Codigo from ao_solicitud_edificacion where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        VAR_AUX = rs_aux2!Codigo
                    End If
                    rs_aux1.AddNew
                    'var_cod = rs_aux1.RecordCount + 1
                    rs_aux1!ges_gestion = glGestion         'Year(Date)
                    rs_aux1!unidad_codigo = parametro
                    rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
                    rs_aux1!edif_codigo = Ado_detalle1.Recordset!edif_codigo
                    rs_aux1!trafico_codigo = var_cod
                    rs_aux1!trafico_h_capacidad_trafico_parametro = Round(VAR_AUX, 2)
                    rs_aux1!estado_codigo = "REG"
                    rs_aux1!fecha_registro = Date
                    rs_aux1!usr_codigo = glusuario
                    rs_aux1.Update
                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
                'End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
            
            Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
            Case "5"    ' SERVICIO MODERNIZACION
            
            Case "11"    ' CONTRATACION DE PERSONAL
                Set rs_aux1 = New ADODB.Recordset
                SQL_FOR = "select * from ro_rrhh_cabecera where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   "
                rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                
                Set rs_aux2 = New ADODB.Recordset
                If rs_aux2.State = 1 Then rs_aux2.Close
                rs_aux2.Open "Select max(rrhh_codigo) as Codigo from ro_rrhh_cabecera where unidad_codigo = '" & parametro & "' ", db, adOpenStatic
                If Not rs_aux2.EOF Then
                    var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                End If
                
                rs_aux1.AddNew
                'var_cod = rs_aux1.RecordCount + 1
                rs_aux1!ges_gestion = glGestion         'Year(Date)
                rs_aux1!unidad_codigo = parametro       'Ado_datos.Recordset!unidad_codigo
                rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
                rs_aux1!rrhh_fecha = Date
                rs_aux1!rrhh_codigo = var_cod
                rs_aux1!beneficiario_codigo_resp = Ado_datos.Recordset!beneficiario_codigo_resp
                rs_aux1!rrhh_descripcion = Ado_datos.Recordset!solicitud_justificacion
                rs_aux1!solicitud_tipo = Ado_datos.Recordset!solicitud_tipo
                rs_aux1!cite_tramite = parametro + "-" + Trim(txt_codigo)
                rs_aux1!estado_codigo = "REG"
                rs_aux1!fecha_registro = Date
                rs_aux1!usr_codigo = glusuario
                rs_aux1.Update
        End Select
        Set rs_aux2 = New ADODB.Recordset
        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9 & "'  "
        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
            Txt_campo1.Caption = rs_aux2!correl_doc
            rs_aux2.Update
        End If
        Ado_datos.Recordset!doc_numero = Txt_campo1.Caption
        'REVISAR !!! JQA 2014_07_08
        'VAR_ARCH = RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
        VAR_ARCH = "ADM_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(Txt_campo1.Caption)))
        Ado_datos.Recordset!archivo_respaldo = VAR_ARCH + ".PDF"
        Ado_datos.Recordset!archivo_respaldo_cargado = "N"
        Ado_datos.Recordset!estado_codigo = "APR"
        Ado_datos.Recordset!fecha_aprueba = Date
        Ado_datos.Recordset!usr_codigo_aprueba = glusuario
        Ado_datos.Recordset.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene DETALLE ...", vbExclamation, "Validación de Registro"
   End If
  Else
      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnBuscar_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Set ClBuscaGrid = New ClBuscaEnGridExterno
        Set ClBuscaGrid.Conexión = db
        ClBuscaGrid.EsTdbGrid = False
        Set ClBuscaGrid.GridTrabajo = dg_datos
        ClBuscaGrid.QueryUtilizado = queryinicial
        Set ClBuscaGrid.RecordsetTrabajo = rs_datos
        'ClBuscaGrid.CamposVisibles = "11010011"
        ClBuscaGrid.Ejecutar
    Else
      MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
    End If
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelUpdate
'        If mvBookMark > 0 Then
'          rs_datos.BookMark = mvBookMark
'        Else
'          rs_datos.MoveFirst
'        End If
        If Ado_datos.Recordset!estado_codigo = "REG" Then
            Call OptFilGral1_Click
        Else
            Call OptFilGral2_Click
        End If
        rs_datos.MoveFirst
        mbDataChanged = False
        Fra_datos.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        FrmABMDet2.Visible = True
        FrmABMDet.Visible = True
'        'FrmABMDet3.Visible = False
        FraDet2.Visible = True
        FraDet1.Visible = True
        dtc_desc1.Enabled = True
        VAR_SW = ""
'        dtc_codigo9.Enabled = True
    End If
'    dtc_desc1.Visible = True
'    lbl_aux1.Visible = False
End Sub

Private Sub btnEliminar_Click()
  On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
    If ExisteReg(parametro, Ado_datos.Recordset!solicitud_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ...", vbInformation + vbOKOnly, "Atención": Exit Sub
    If rs_datos!estado_codigo = "APR" Then
       sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If sino = vbYes Then
          rs_datos!estado_codigo = "ERR"
          rs_datos!fecha_registro = Date
          rs_datos!usr_codigo = glusuario
          rs_datos.UpdateBatch adAffectAll
       End If
    Else
       MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
    End If
  Else
      MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  Exit Sub
  
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnDesAprobar_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_datos!estado_codigo = "APR" Then
      If sino = vbYes Then
         rs_datos!estado_codigo = "REG"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If VAR_SW = "ADD" Then
        VAR_UNI = parametro     'dtc_codigo1.Text
        var_cod = IIf(txt_codigo.Caption = "", 0, txt_codigo.Caption)
        Set rs_aux1 = New ADODB.Recordset
        'SQL_FOR = "select * from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & var_cod & "  "
        SQL_FOR = "select * from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            var_cod = rs_aux1.RecordCount + 1
            'MsgBox "El código ya existe, consulte con el administrador del Sistema..."
            'var_cod = 0
            'Exit Sub
        Else
            'var_cod = rs_datos.RecordCount '+ 1
            var_cod = 1
        End If
        'var_cod = RTrim(RTrim(dtc_codigo2.Text) + "-") + LTrim(Str(Val(dtc_aux2) + 1))
        txt_codigo.Caption = var_cod
        rs_datos!solicitud_codigo = var_cod
        rs_datos!estado_codigo = "REG"      'no cambia
        'rs_datos!ges_gestion = Year(DTPfecha1.Value) 'glGestion    'Year(Date)   'no cambia
        rs_datos!ges_gestion = Year(Date)
        rs_datos!unidad_codigo = VAR_UNI
        'Actualiza correaltivo ...
        db.Execute "Update gc_unidad_ejecutora Set correl_solicitud = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "'   "
        rs_datos!doc_numero = "0"    'txt_campo1.Caption
        'rs_datos!correl_edificacion = 0
        rs_datos!archivo_respaldo = "sin_nombre"
        rs_datos!archivo_respaldo_cargado = "N"
        rs_datos!correl_bitacora = 0
     End If
     rs_datos!beneficiario_codigo_resp2 = "0"       'dtc_aux4.Text
     rs_datos!unidad_codigo_sol = dtc_codigo3.Text      'UNIDAD SOLICITANTE
     rs_datos!solicitud_fecha_solicitud = DTPfecha1.Value
     rs_datos!solicitud_tipo = dtc_codigo2.Text
     rs_datos!edif_codigo = "20101-0"       'IIf(dtc_codigo3.Text = "", "10101-0", dtc_codigo3.Text)
     'If dtc_codigo4.Text = "" Or dtc_codigo4.Text = "0" Then
     '   rs_datos!beneficiario_codigo = dtc_aux3.Text
     'Else
        rs_datos!beneficiario_codigo = dtc_codigo4.Text
     'End If
     rs_datos!solicitud_justificacion = Txt_descripcion.Text
     
     Select Case dtc_codigo2.Text
        Case "1"    'SOLO COMPRAS BB y SS
        Case "2"    'SOLO VENTA DE BIENES
        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
  
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-01-02"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "COM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-234"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
        Case "5"    ' SERVICIO MODERNIZACION
        
        Case "11"    ' CONTRATACION DE PERSONAL DE PLANTA
                rs_datos!proceso_codigo = "ADM"
                rs_datos!subproceso_codigo = "ADM-01"
                rs_datos!etapa_codigo = "ADM-01-01"
                rs_datos!clasif_codigo = "ADM"
                rs_datos!doc_codigo = "R-170"
        Case "12"    ' CONTRATACION DE PERSONAL POR PRODUCTO
                rs_datos!proceso_codigo = "ADM"
                rs_datos!subproceso_codigo = "ADM-01"
                rs_datos!etapa_codigo = "ADM-01-01"
                rs_datos!clasif_codigo = "ADM"
                rs_datos!doc_codigo = "R-170"
     End Select
     rs_datos!poa_codigo = IIf(dtc_codigo10.Text = "", "7.1.1", dtc_codigo10.Text)
     rs_datos!solicitud_observaciones = txt_obs.Text
     rs_datos!solicitud_fecha_recepción = DTPfecha1.Value
     rs_datos!beneficiario_codigo_resp = dtc_codigo11.Text

     rs_datos!ges_gestion_ant = glGestion       'Year(Date)
     var_cod = rs_datos!solicitud_codigo
     If var_cod < 10 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-00000" + Trim(txt_codigo)
     End If
     If var_cod > 9 And var_cod < 100 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-0000" + Trim(txt_codigo)
     End If
     If var_cod > 99 And var_cod < 1000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-000" + Trim(txt_codigo)
     End If
     If var_cod > 999 And var_cod < 10000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-00" + Trim(txt_codigo)
     End If
     If var_cod > 9999 And var_cod < 100000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-0" + Trim(txt_codigo)
     End If
     If var_cod > 99999 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-" + Trim(txt_codigo)
     End If
     rs_datos!solicitud_codigo_ant = 0
     rs_datos!usr_codigo_aprueba = ""
     rs_datos!fecha_aprueba = Date
     rs_datos!hora_aprueba = ""
     'rs_datos!Foto = Date
     'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
     'rs_datos!archivo_foto_cargado = "N"
     'hora_registro
     rs_datos!fecha_registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     rs_datos.Update    'Batch 'adAffectAll
     If Ado_datos.Recordset!estado_codigo = "REG" Then
        Call OptFilGral1_Click
     Else
        Call OptFilGral2_Click
     End If
     rs_datos.MoveLast
     mbDataChanged = False
      
     Fra_datos.Enabled = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
     FrmABMDet2.Visible = True
     FrmABMDet.Visible = True
     FraDet2.Visible = True
     FraDet1.Visible = True
     dtc_desc1.Enabled = True
'     dtc_desc1.BackColor = &HFFFFC0
     VAR_SW = ""
      
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  If (dtc_codigo1.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo2.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo11.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo3.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo4.Text = "") Or (dtc_codigo4.Text = "0") Then
    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If (dtc_codigo10.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo10.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle2.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR01.ReportFileName = App.Path & "\Reportes\RRHH\rr_lista_solicitud_rrhh.rpt"
        
        CR01.WindowShowPrintSetupBtn = True
        CR01.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
        CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
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
          Case "DRRHH"
              var_titulo = "Módulo Recursos Humanos"
      End Select
      CR01.Formulas(3) = "titulo = '" & var_titulo & "' "
      CR01.Formulas(4) = "subtitulo = '" & lbl_titulo.Caption & "' "

        iResult = CR01.PrintReport
        If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
        CR01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
End Sub

Private Sub BtnImprimir1_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR03.ReportFileName = App.Path & "\Reportes\RRHH\rr_bitacora_rrhh.rpt"
        CR03.WindowShowPrintSetupBtn = True
        CR03.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          CR03.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
          CR03.Formulas(1) = "Subtitulo = '" & FraDet1.Caption & "' "
        'Call CREAVISTAF11          'JQA JUN-2008
        CR03.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR03.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR03.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR01.PrintReport
        If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
        CR03.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
End Sub

Private Sub BtnImprimir2_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If (Ado_detalle2.Recordset.RecordCount > 0) Then
      Dim iResult As Variant  ', i%, y%
      CR02.ReportFileName = App.Path & "\reportes\RRHH\rr_solicitud_contratacion.rpt"
      CR02.WindowShowPrintSetupBtn = True
      CR02.WindowShowRefreshBtn = True
      CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
      CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
      CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
      var_titulo = "Módulo Recursos Humanos"
      CR02.Formulas(3) = "titulo = '" & var_titulo & "' "
      CR02.Formulas(4) = "subtitulo = '" & lbl_titulo.Caption & "' "
      iResult = CR02.PrintReport
      If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
    Else
      MsgBox "No se puede IMPRIMIR el registro, Debe registrar el Detalle, vuelva a intentar ...", , "Atención"
    End If
  Else
    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If

End Sub

Private Sub BtnModDetalle_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  'If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" And Ado_detalle1.Recordset.RecordCount > 0 Then
  If rs_datos.RecordCount > 0 And Ado_detalle1.Recordset.RecordCount > 0 Then
    swnuevo = 2
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False
        
    frm_ao_solicitud_bitacora.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
    frm_ao_solicitud_bitacora.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
    frm_ao_solicitud_bitacora.Txt_descripcion.Caption = Me.dtc_desc1.Text
    frm_ao_solicitud_bitacora.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
    'frm_ao_solicitud_bitacora.Txt_estado.Caption = "REG"
    'Ado_detalle1.Recordset.AddNew
     
    frm_ao_solicitud_bitacora.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("negocia_forma")
    frm_ao_solicitud_bitacora.DTPfecha1.Value = Me.Ado_detalle1.Recordset("negocia_fecha_real")
    frm_ao_solicitud_bitacora.Txt_campo2.Value = Me.Ado_detalle1.Recordset("negocia_hora_real")
    frm_ao_solicitud_bitacora.Txt_monto1.Text = Me.Ado_detalle1.Recordset("negocia_gasto_estimado")
    frm_ao_solicitud_bitacora.dtc_codigo2.Text = Me.Ado_detalle1.Recordset("beneficiario_codigo")
    frm_ao_solicitud_bitacora.dtc_codigo3.Text = Me.Ado_detalle1.Recordset("beneficiario_codigo_resp")
    frm_ao_solicitud_bitacora.Txt_campo3.Text = Me.Ado_detalle1.Recordset("negocia_tarea_realizada")
    frm_ao_solicitud_bitacora.Txt_campo4.Text = Me.Ado_detalle1.Recordset("negocia_observaciones")
    frm_ao_solicitud_bitacora.Txt_campo5.Text = Me.Ado_detalle1.Recordset("bitacora_cite")
    If swnuevo = 2 Then
        frm_ao_solicitud_bitacora.dtc_desc1.BoundText = frm_ao_solicitud_bitacora.dtc_codigo1.BoundText
        frm_ao_solicitud_bitacora.dtc_desc2.BoundText = frm_ao_solicitud_bitacora.dtc_codigo2.BoundText
        frm_ao_solicitud_bitacora.dtc_desc3.BoundText = frm_ao_solicitud_bitacora.dtc_codigo3.BoundText
'        frm_ao_solicitud_bitacora.HH = Left(frm_ao_solicitud_bitacora.Txt_campo2.Text, 2)
'        frm_ao_solicitud_bitacora.MM = Right(frm_ao_solicitud_bitacora.Txt_campo2.Text, 2)
    End If
    
    frm_ao_solicitud_bitacora.Show vbModal
    
    Call ABRIR_TABLA_DET
    
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    'Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Modificar, verifique si fue correctamente identificado !! ", vbExclamation
  End If
End Sub

Private Sub BtnModDetalle2_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" And Ado_detalle2.Recordset.RecordCount > 0 Then
    swnuevo = 2
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
'    FraDet3.Enabled = False
'    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False

    Select Case dtc_codigo2.Text
        Case "1"    'SOLO COMPRAS BB y SS
        Case "2"    'SOLO VENTA DE BIENES
        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
            Call ABRIR_TABLA_DET
                   
            aw_p_ao_solicitud_edificacion.Show vbModal
        Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
        Case "5"    ' SERVICIO MODERNIZACION
        Case "11"    ' CONTRATACION DE PERSONAL
            Call ABRIR_TABLA_DET
            frm_ao_solicitud_persona.txt_codigo.Caption = Ado_detalle2.Recordset("solicitud_codigo")
            frm_ao_solicitud_persona.Txt_campo1.Text = Ado_detalle2.Recordset("unidad_codigo")
            frm_ao_solicitud_persona.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_ao_solicitud_persona.txtBenef = Ado_detalle2.Recordset("benef_id")
            frm_ao_solicitud_persona.txtEstado.Text = "REG"
            
'            frm_ao_solicitud_persona.Frame4.Visible = True
            
            frm_ao_solicitud_persona.txtPat.Text = Ado_detalle2.Recordset("benef_primer_apellido")
            frm_ao_solicitud_persona.txtMat.Text = Ado_detalle2.Recordset("benef_segundo_apellido")
            frm_ao_solicitud_persona.txtNom.Text = Ado_detalle2.Recordset("benef_nombres")
            frm_ao_solicitud_persona.TxtDireccion.Text = Ado_detalle2.Recordset("benef_direccion_domicilio")
            frm_ao_solicitud_persona.txtTelefono.Text = Ado_detalle2.Recordset("benef_telefonos_ref")
            frm_ao_solicitud_persona.txtCI.Text = Ado_detalle2.Recordset("benef_codigo")
            frm_ao_solicitud_persona.dtc_codigo1.Text = Ado_detalle2.Recordset("puesto_codigo")
            frm_ao_solicitud_persona.dtc_desc1.BoundText = frm_ao_solicitud_persona.dtc_codigo1.BoundText
            frm_ao_solicitud_persona.dtc_codigo2.Text = Ado_detalle2.Recordset("ocup_codigo")
            frm_ao_solicitud_persona.dtc_desc2.BoundText = frm_ao_solicitud_persona.dtc_codigo2.BoundText
            frm_ao_solicitud_persona.dtc_codigo4.Text = Ado_detalle2.Recordset("munic_codigo")
            frm_ao_solicitud_persona.dtc_desc4.BoundText = frm_ao_solicitud_persona.dtc_codigo4.BoundText
            frm_ao_solicitud_persona.dtc_codigo3.Text = Ado_detalle2.Recordset("nivel_educ_codigo")
            frm_ao_solicitud_persona.dtc_desc3.BoundText = frm_ao_solicitud_persona.dtc_codigo3.BoundText
            frm_ao_solicitud_persona.dtc_desc1.Text = Ado_detalle2.Recordset("observaciones")
            frm_ao_solicitud_persona.txtFecha.Value = Ado_detalle2.Recordset("benef_fecha")
            
            frm_ao_solicitud_persona.Show vbModal
    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
'    FraDet3.Enabled = True
'    FrmABMDet3.Enabled = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede MODIFICAR, porque el registro ya está Aprobado o No fue identificado!! ", vbExclamation
  End If

End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
  If Ado_datos.Recordset.RecordCount > 0 Then
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        Fra_datos.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        FrmABMDet2.Visible = False
        FrmABMDet.Visible = False
'        FrmABMDet3.Visible = False
        FraDet2.Visible = False
        FraDet1.Visible = False
'        FraDet3.Visible = False
        VAR_SW = "MOD"
        Text7.Visible = True
        VAR_UNI = parametro
        dtc_desc1.Enabled = False
        dtc_codigo1.BoundText = dtc_desc1.BoundText
        DTPfecha1.Value = Ado_datos.Recordset!solicitud_fecha_solicitud
'        dtc_aux1.BoundText = dtc_desc1.BoundText
        Call pnivel1(dtc_codigo1.BoundText)
        dtc_desc10.Enabled = True
        dtc_desc4.Enabled = False
'        Call pnivel11(dtc_codigo1.BoundText)
'        dtc_desc11.Enabled = True
    '    lbl_aux1.Visible = True
    '    lbl_aux1.Caption = dtc_desc1.Text
        'dtc_desc4.Enabled = False
        Txt_descripcion.SetFocus
    '    BtnVer.Visible = True
    Else
      MsgBox "No se puede MODIFICAR un registro APROBADO o Anulado ...", vbExclamation, "Validación de Registro"
    End If
  Else
        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnSalir_Click()
'  If glPersOtro = "O" Then
'    frmmo_pacientes.Dtc_ocupac = rs_datos!ocup_codigo
'    frmmo_pacientes.Dtc_OcupacDes = rs_datos!ocup_descripcion
'  End If
'  glPersOtro = "N"
  Unload Me
End Sub

Private Sub BtnVer_Click()
  On Error GoTo QError
  If rs_datos!estado_codigo = "APR" Then
    Dim ARCH_FOTO As String
    Dim SW0 As String
    Select Case Left(Trim(Ado_datos.Recordset("edif_codigo")), 1)
        Case "1"    'CHQ
            VAR_DPTO = "CHQ"
        Case "2"    'LPZ
            VAR_DPTO = "LPZ"
        Case "3"    'CBB
            VAR_DPTO = "CBB"
        Case "4"    'SCZ
            VAR_DPTO = "SCZ"
        Case "5"    'PTS
            VAR_DPTO = "PTS"
        Case "6"    'ORU
            VAR_DPTO = "ORU"
        Case "7"    'TJA
            VAR_DPTO = "TJA"
        Case "8"    'BEN
            VAR_DPTO = "BEN"
        Case "9"    'PDO
            VAR_DPTO = "PDO"
    End Select
    If Ado_datos.Recordset!archivo_respaldo_cargado = "N" Then
      'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!edif_tipo) & "\" & Trim(Ado_datos.Recordset!negocia_codigo) & "\"
      NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "DED2"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'      Else
         e = NombreCarpeta
'      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
      SW0 = 1
    Else
      'MsgBox ""
      'negocia_codigo, unidad_codigo, negocia_fecha_inicio as fecha1, negocia_descripcion, estado_codigo, fecha_registro, usr_codigo, solicitud_tipo as codigo2, edif_codigo as codigo3, beneficiario_codigo as codigo4, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero As campo1, poa_codigo As codigo10, hora_registro, ges_gestion, archivo_respaldo, archivo_respaldo_cargado
      sino = MsgBox("El archivo ya existe, elija: <SI> para Volver a Cargarlo. <NO> para Visualizarlo. ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!edif_tipo) & "\" & Trim(Ado_datos.Recordset!negocia_codigo) & "\"
          NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "DED2"
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'          Else
            e = NombreCarpeta
'          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
          SW0 = 1
      Else
        SW0 = 0
        'e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!codigo_beneficiario) & "\LICENCIAS\" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
        e = ShellExecute(0, vbNullString, App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\" & Trim(Ado_datos.Recordset("archivo_respaldo")), vbNullString, vbNullString, vbNormalFocus)
      End If
    End If
    '    If SW0 = 1 Then
    '    '    If GlServidor = "SRVPRO" Then
    '    '        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
    '    '    Else
    '            'ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset!edif_tipo) + "\" + Trim(Ado_datos.Recordset!edif_codigo)
    '            ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset!edif_tipo) + "\" + Trim(Ado_datos.Recordset!edif_codigo) + ".JPG"
    '    '    End If
    '        'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + Ado_datos.Recordset!codigo_beneficiario + "\" + Ado_datos.Recordset("codigo_beneficiario") + "-FOTO.JPG"
    '        CodBien = Ado_datos.Recordset!edif_codigo
    '        If Guardar_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo= '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
    '            MsgBox "Se cargo la Imagen Correctamente !!"
    '        Else
    '            MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    '        End If
    '    Else
    '        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
    '        Image2 = Img_Foto
    '    End If
  Else
       MsgBox "No se puede Guardar el documento PDF, debe APROBAR previamente el registro ...", vbExclamation, "Validación de Registro"
  End If
QError:
    ' Manejo de errores
    If Err.Number > 0 Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
    '    db.RollbackTrans
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub dtc_aux1_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_aux1.BoundText
    dtc_codigo11.BoundText = dtc_aux1.BoundText
End Sub

'Private Sub dtc_aux3_Click(Area As Integer)
'    dtc_codigo3.BoundText = dtc_aux3.BoundText
'    dtc_desc3.BoundText = dtc_aux3.BoundText
'End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
    dtc_desc10.BoundText = dtc_codigo10.BoundText
End Sub

Private Sub dtc_codigo11_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_codigo11.BoundText
    dtc_aux1.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

'Private Sub dtc_codigo5_Click(Area As Integer)
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'End Sub

'Private Sub dtc_codigo6_Click(Area As Integer)
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'End Sub

'Private Sub dtc_codigo7_Click(Area As Integer)
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'End Sub

'Private Sub dtc_codigo8_Click(Area As Integer)
'    dtc_desc8.BoundText = dtc_codigo8.BoundText
'End Sub

'Private Sub dtc_codigo9_Click(Area As Integer)
'    dtc_desc9.BoundText = dtc_codigo9.BoundText
'End Sub

'Private Sub dtc_codigo9_LostFocus()
''  If VAR_SW = "ADD" Then
''    Set rs_aux2 = New ADODB.Recordset
''    SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9.Text & "'  "
''    rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
''    If rs_aux2.RecordCount > 0 Then
''        rs_aux2!correl_doc = rs_aux2!correl_doc + 1
''        txt_campo1.Caption = rs_aux2!correl_doc
''        rs_aux2.Update
''    End If
''  End If
'  txt_aux9.Text = dtc_desc9.Text
'End Sub

'Private Sub dtc_desc5_Click(Area As Integer)
'    dtc_codigo5.BoundText = dtc_desc5.BoundText
''    Call pnivel5(dtc_codigo5.BoundText)
''    dtc_desc6.Enabled = True
'End Sub
   
'Private Sub pnivel5(codigo5 As String)
'   'Dim strConsultaF As String
'   'strConsultaF = "select * from gc_proceso_nivel2 where proceso_codigo = '" & codigo5 & "'"
'
'   Set dtc_codigo6.RowSource = Nothing
'   'Set dtc_codigo6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_codigo6.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel2 '" & codigo5 & "' ")
'   dtc_codigo6.ReFill
'   dtc_codigo6.BoundText = Empty
'
'   Set dtc_desc6.RowSource = Nothing
'   'Set dtc_desc6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_desc6.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel2 '" & codigo5 & "' ")
'   dtc_desc6.ReFill
'   dtc_desc6.BoundText = Empty
'End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    Call pnivel1(dtc_codigo1.BoundText)
    dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc4.Enabled = True
End Sub
   
Private Sub pnivel1(codigo1 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
   
   Set dtc_codigo10.RowSource = Nothing
'   Set dtc_codigo10.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_codigo10.ReFill
   dtc_codigo10.BoundText = Empty
   
   Set dtc_desc10.RowSource = Nothing
   'Set dtc_desc10.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_desc10.ReFill
   dtc_desc10.BoundText = Empty
End Sub
  
Private Sub pnivel11(codigo1 As String)
   Dim strConsultaF As String
   strConsultaF = "Select * from rv_unidad_vs_responsable where unidad_codigo = '" & codigo1 & "' order by beneficiario_denominacion"

   Set dtc_codigo4.RowSource = Nothing
   Set dtc_codigo4.RowSource = db.Execute(strConsultaF, , adCmdText)

   dtc_codigo4.ReFill
   dtc_codigo4.BoundText = Empty

   Set dtc_desc4.RowSource = Nothing
   Set dtc_desc4.RowSource = db.Execute(strConsultaF, , adCmdText)

   dtc_desc4.ReFill
   dtc_desc4.BoundText = Empty
End Sub

'Private Sub dtc_desc1_LostFocus()
''    dtc_codigo5.Text = dtc_aux1.Text
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
'    Call pnivel5(dtc_codigo5.BoundText)
'    dtc_desc6.Enabled = True
'End Sub

Private Sub dtc_desc10_Click(Area As Integer)
    dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub

Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
    dtc_aux1.BoundText = dtc_desc11.BoundText
    VAR_UNI2 = IIf(dtc_aux1.Text = "", "DRRHH", dtc_aux1.Text)
    Call pnivel1(VAR_UNI2)
    dtc_desc10.Enabled = True
End Sub

Private Sub dtc_desc11_LostFocus()
    'Txt_descripcion.Text = lbl_titulo + " - " + dtc_desc2.Text
    Txt_descripcion.Text = "SOLICITUD - " + dtc_desc2.Text
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc2_LostFocus()
    'Txt_descripcion.Text = lbl_titulo + " - " + dtc_desc2.Text
    Txt_descripcion.Text = "SOLICITUD - " + dtc_desc2.Text
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    Call pnivel3(dtc_codigo3.BoundText)
    dtc_desc4.Enabled = True
End Sub
 
Private Sub dtc_desc3_LostFocus()
'    dtc_codigo3.BoundText = dtc_desc3.BoundText
'    Call pnivel3(dtc_codigo3.BoundText)
'    dtc_desc4.Enabled = True
End Sub
  
Private Sub pnivel3(codigo3 As String)
   Dim strConsultaF As String
   strConsultaF = "Select * from rv_unidad_vs_responsable where unidad_codigo = '" & codigo3 & "' order by beneficiario_denominacion"

   Set dtc_codigo4.RowSource = Nothing
   Set dtc_codigo4.RowSource = db.Execute(strConsultaF, , adCmdText)

   dtc_codigo4.ReFill
   dtc_codigo4.BoundText = Empty

   Set dtc_desc4.RowSource = Nothing
   Set dtc_desc4.RowSource = db.Execute(strConsultaF, , adCmdText)

   dtc_desc4.ReFill
   dtc_desc4.BoundText = Empty
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    If Aux = "" Then
        parametro = "DRRHH"
    Else
        parametro = Aux
    End If
    'parametro = "estado_codigo" + " = " + "'REG'"
    '
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    'JQA 2014-JUL-14
    'db.Execute (" EXEC gp_actualiza_beneficiario_edif ")
'    lbl_aux1.Visible = False
    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_unidad_ejecutora
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
'    dtc_aux1.BoundText = dtc_codigo1.BoundText
    
    'gc_tipo_solicitud
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from gc_tipo_solicitud WHERE solicitud_tipo = '11' OR solicitud_tipo = '12' ", db, adOpenStatic
    'rs_datos2.Open "gp_listar_apr_gc_tipo_solicitud", DB, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    'gc_unidad_ejecutora       Unidad solicitante
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'gc_beneficiario (Personas Nat. y Juridicas / Clientes, Proveedores, etc.)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "gp_listar_gc_beneficiario_personas", DB, adOpenStatic
    rs_datos4.Open "select * from rv_unidad_vs_responsable ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
'    Set rs_datos5 = New ADODB.Recordset
'    If rs_datos5.State = 1 Then rs_datos5.Close
'    'rs_datos5.Open "Select * from gc_proceso_nivel1 order by proceso_descripcion", db, adOpenStatic
'    rs_datos5.Open "gp_listar_apr_gc_proceso_nivel1", db, adOpenStatic
'    Set Ado_datos5.Recordset = rs_datos5
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
'
'    Set rs_datos6 = New ADODB.Recordset
'    If rs_datos6.State = 1 Then rs_datos6.Close
'    'rs_datos6.Open "Select * from gc_proceso_nivel2 order by subproceso_descripcion", db, adOpenStatic
'    rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
'    Set Ado_datos6.Recordset = rs_datos6
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'
'    Set rs_datos7 = New ADODB.Recordset
'    If rs_datos7.State = 1 Then rs_datos7.Close
'    'rs_datos7.Open "Select * from gc_proceso_nivel3 order by etapa_descripcion", db, adOpenStatic
'    rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
'    Set Ado_datos7.Recordset = rs_datos7
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'
'    Set rs_datos8 = New ADODB.Recordset
'    If rs_datos8.State = 1 Then rs_datos8.Close
'    'rs_datos8.Open "Select * from gc_documentos_clasificacion order by clasif_codigo", db, adOpenStatic
'    rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
'    Set Ado_datos8.Recordset = rs_datos8
''    dtc_desc8.BoundText = dtc_codigo8.BoundText
    
'    'gc_documentos_respaldo
'    Set rs_datos9 = New ADODB.Recordset
'    If rs_datos9.State = 1 Then rs_datos9.Close
'    'rs_datos9.Open "Select * from gc_documentos_respaldo order by doc_codigo", db, adOpenStatic
'    rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
'    Set Ado_datos9.Recordset = rs_datos9
'    dtc_desc9.BoundText = dtc_codigo9.BoundText
    
    'pc_poa_actividad
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    'rs_datos10.Open "Select * from pc_poa_actividad order by poa_codigo", db, adOpenStatic
    rs_datos10.Open "pp_listar_apr_pc_poa_actividad", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText
    
    'gc_beneficiario (Personal CGI)
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

'Private Sub ABRIR_TABLA()
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    'queryinicial = "select solicitud_codigo, unidad_codigo, solicitud_justificacion, solicitud_observaciones, estado_codigo, fecha_registro, usr_codigo, hora_registro, ges_gestion, solicitud_fecha_solicitud as fecha1,  solicitud_fecha_recepción as fecha2, solicitud_tipo as codigo2, beneficiario_codigo as codigo4, beneficiario_codigo_resp as codigo11, edif_codigo as codigo3, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero As campo1, poa_codigo As codigo10, archivo_respaldo, archivo_respaldo_cargado, ges_gestion_ant, unidad_codigo_ant, solicitud_codigo_ant, usr_codigo_aprueba, fecha_aprueba, hora_aprueba From ao_solicitud WHERE estado_codigo = 'REG' "
'    queryinicial = "Select * from ao_solicitud where " + parametro
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    Set dg_datos.DataSource = Ado_datos.Recordset
'End Sub

Public Sub ABRIR_TABLA_DET()
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    'rs_det1.Open "select * from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
    rs_det1.Open "select * from ao_solicitud_bitacora where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    Set dg_det1.DataSource = Ado_detalle1.Recordset
    
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    'rs_aux2.Open "select * from ao_solicitud_persona where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", DB, adOpenKeyset, adLockOptimistic, adCmdText
    rs_aux2.Open "select * from av_solicitud_persona where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle2.Recordset = rs_aux2
    Set dg_det2.DataSource = Ado_detalle2.Recordset
End Sub

Private Sub ABRIR_TABLA_AUX2()
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
    
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If Ado_datos.Recordset.RecordCount > 0 Then
    'Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
    ' <-- Inicio                Identificación del Cliente                Fin -->   'esto es de Caption
    'Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
    'Image2 = Img_Foto
'    If Ado_datos.Recordset!archivo_foto_cargado = "S" Then
'        'BtnVer.Visible = True
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
'        Image2 = Img_Foto
'    Else
'        'BtnVer.Visible = False
'        'chkEstado.Value = vbUnchecked
'    End If
    If VAR_SW <> "ADD" Then
        Select Case rs_datos!solicitud_tipo     'dtc_codigo2.Text
            Case "1"    'SOLO COMPRAS BB y SS
            Case "2"    'SOLO VENTA DE BIENES
            Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
                Call ABRIR_TABLA_DET
            Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
                Call ABRIR_TABLA_DET
            Case "5"    ' SERVICIO MODERNIZACION
            
            Case "7"    'SOLO COMPRAS BB y SS
            Case "8"    'SOLO VENTA DE BIENES
            Case "9"    ' COMPRA-VENTA BB Y SS - COMERCIAL
                Call ABRIR_TABLA_DET
            Case "10"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
                Call ABRIR_TABLA_DET
            Case "11"    ' RRHH
                Call ABRIR_TABLA_DET
                DTPfecha1.Value = Ado_datos.Recordset!solicitud_fecha_solicitud
                dtc_codigo4.Text = Ado_datos.Recordset!beneficiario_codigo_resp2
        End Select
'        Call ABRIR_TABLA_AUX2
    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det2.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
    FraDet1.Caption = "BITÁCORA DE " + parametro
'    txt_aux9.Text = dtc_desc9.Text
    If Ado_datos.Recordset!estado_codigo = "REG" Then
            FrmABMDet2.Visible = True
    Else
    
            FrmABMDet2.Visible = False
    End If
  End If
End Sub

Private Sub Ado_datos_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub BtnAñadir_Click()
  On Error GoTo AddErr
    VAR_SW = "ADD"
    DTPfecha1.Value = Date
    'lblStatus.Caption = "Agregar registro"
    Fra_datos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    FrmABMDet2.Visible = False
    FrmABMDet.Visible = False
'    FrmABMDet3.Visible = False
    FraDet2.Visible = False
    FraDet1.Visible = False
'    FraDet3.Visible = False
    
    dtc_desc4.Enabled = False
'    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
'    rs_datos.AddNew
    Ado_datos.Recordset.AddNew
    dtc_desc11.SetFocus
    'dtc_desc1.BackColor = &H80000005
    dtc_codigo1.Text = parametro
    dtc_desc1.BoundText = dtc_codigo1.BoundText
'    dtc_aux1.BoundText = dtc_codigo1.BoundText
    dtc_desc2.Locked = True
    Select Case parametro
        Case "DVTA"        'INI COMERCIAL
            dtc_codigo2.Text = 3
        Case "COMEX"        'INI COMEX
            dtc_codigo2.Text = 3
        Case "DNINS"                        'INI GRABA INSTALACIONES
            '
            dtc_codigo2.Text = 4
        Case "DNAJS"
            '
            dtc_codigo2.Text = 4
        Case "DNMAN"
            '
            dtc_codigo2.Text = 4
        Case "DRRHH"
            '
            dtc_codigo2.Text = 11
            dtc_codigo3 = "20101-0"
        Case Else
            dtc_codigo2.Text = 5
    End Select
    DTPfecha1.Value = Date
    dtc_desc2.BoundText = dtc_codigo2.BoundText
'    dtc_codigo5.Text = "COM"
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'    dtc_codigo6.Text = "COM-01"
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'    dtc_codigo7.Text = "COM-01-02"
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'    BtnVer.Visible = False
'    dtc_codigo9.Enabled = False
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  rs_datos.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Function ExisteReg(Unidad As String, Codigo As Integer) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ro_rrhh_cabecera WHERE unidad_codigo = '" & Unidad & "' and solicitud_codigo = " & Codigo & " "
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub OptFilGral1_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "Select * from ao_solicitud where estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "' "
    queryinicial = "Select * from ao_solicitud where estado_codigo = 'REG' AND solicitud_tipo = '11' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "Select * from ao_solicitud where unidad_codigo = '" & parametro & "' "
    queryinicial = "Select * from ao_solicitud where solicitud_tipo = '11' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_obs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form mw_solicitud_calculo_trafico_mod 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Módulo Modernizacion - Parámetros de Cálculo"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13950
   Icon            =   "mw_solicitud_calculo_trafico_mod.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   13950
   WindowState     =   2  'Maximized
   Begin VB.Frame FrmDetalle 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ELEGIR EL EQUIPO (DEL EDIFICIO) A SER MODERNIZADO"
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
      Height          =   2745
      Left            =   120
      TabIndex        =   183
      Top             =   3840
      Visible         =   0   'False
      Width           =   19095
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   18825
         TabIndex        =   184
         Top             =   240
         Width           =   18825
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   7320
            Picture         =   "mw_solicitud_calculo_trafico_mod.frx":0A02
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   185
            ToolTipText     =   "Modifica datos del Grupo elegido"
            Top             =   0
            Width           =   1430
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edificio:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   225
            TabIndex        =   187
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lbl_texto2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   300
            Left            =   1320
            TabIndex        =   186
            Top             =   120
            Width           =   615
         End
      End
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":1317
         Height          =   1740
         Left            =   120
         TabIndex        =   188
         Top             =   960
         Width           =   18855
         _ExtentX        =   33258
         _ExtentY        =   3069
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
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
         ColumnCount     =   17
         BeginProperty Column00 
            DataField       =   "edif_codigo"
            Caption         =   "Edificio"
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
            Caption         =   "Codigo.Equipo"
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
            DataField       =   "tipo_eqp_descripcion"
            Caption         =   "Tipo.Equipo"
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
            DataField       =   "marca_descripcion"
            Caption         =   "Marca.de.Equipo"
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
         BeginProperty Column04 
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo.de.Equipo"
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
            DataField       =   "trafico_num_paradas"
            Caption         =   "#Paradas"
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
            DataField       =   "recorrido_codigo"
            Caption         =   "Recorrido"
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
         BeginProperty Column07 
            DataField       =   "pasajeros_numero"
            Caption         =   "#Pasajeros"
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
         BeginProperty Column08 
            DataField       =   "pasajeros_capacidad_km_d"
            Caption         =   "Capacidad.Kg"
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
         BeginProperty Column09 
            DataField       =   "vel_equipo_m_s"
            Caption         =   "Velocidad.m/s"
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
         BeginProperty Column10 
            DataField       =   "tipo_puerta_descripcion"
            Caption         =   "Tipo.Puerta"
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
         BeginProperty Column11 
            DataField       =   "trafico_ancho_puerta"
            Caption         =   "Ancho.Puerta"
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
         BeginProperty Column12 
            DataField       =   "cabina_descripcion"
            Caption         =   "Cabina"
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
         BeginProperty Column13 
            DataField       =   "tecnologia_descripcion"
            Caption         =   "Tecnología"
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
         BeginProperty Column14 
            DataField       =   "sist_puerta_descripcion"
            Caption         =   "Sistema.Puertas"
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
         BeginProperty Column15 
            DataField       =   "condicion_ventas_descripcion"
            Caption         =   "Condicion.Venta"
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
         BeginProperty Column16 
            DataField       =   "condicion_cabina_descripcion"
            Caption         =   "Cabina"
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
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column16 
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20520
      TabIndex        =   168
      Top             =   0
      Width           =   20520
      Begin VB.PictureBox BtnVer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7200
         Picture         =   "mw_solicitud_calculo_trafico_mod.frx":1332
         ScaleHeight     =   615
         ScaleWidth      =   1440
         TabIndex        =   182
         ToolTipText     =   "Ver Cálculos de Tráfico"
         Top             =   0
         Width           =   1440
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "mw_solicitud_calculo_trafico_mod.frx":1E3E
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   176
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   9960
         Picture         =   "mw_solicitud_calculo_trafico_mod.frx":2600
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8760
         Picture         =   "mw_solicitud_calculo_trafico_mod.frx":280A
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   174
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   105
         Picture         =   "mw_solicitud_calculo_trafico_mod.frx":2FC9
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   173
         ToolTipText     =   "Editar Datos de ""Cabecera Cronograma"""
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1440
         Picture         =   "mw_solicitud_calculo_trafico_mod.frx":38DE
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   172
         ToolTipText     =   "Anula el Registro Activo"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "mw_solicitud_calculo_trafico_mod.frx":402A
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   171
         ToolTipText     =   "Aprueba el Cronograma"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4200
         Picture         =   "mw_solicitud_calculo_trafico_mod.frx":485D
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   170
         ToolTipText     =   "Busca un Registro"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5640
         Picture         =   "mw_solicitud_calculo_trafico_mod.frx":5012
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   169
         ToolTipText     =   "Imprime Lista de Cronogramas"
         Top             =   0
         Width           =   1400
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO"
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
         Left            =   12600
         TabIndex        =   177
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   2400
      Left            =   120
      TabIndex        =   136
      Top             =   3600
      Width           =   5490
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         DataField       =   "estado_codigo_verif"
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
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   167
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         Height          =   290
         Left            =   3240
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   162
         Top             =   515
         Width           =   320
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4935
         TabIndex        =   154
         Top             =   1890
         Width           =   375
      End
      Begin VB.TextBox Txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         DataField       =   "trafico_codigo"
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
         Height          =   290
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   140
         Top             =   795
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.TextBox txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         DataField       =   "solicitud_codigo"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   139
         Top             =   1575
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.TextBox Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   138
         Top             =   1575
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5055
         TabIndex        =   137
         Top             =   1200
         Width           =   260
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":58DF
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4320
         TabIndex        =   141
         Top             =   840
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":58F8
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4320
         TabIndex        =   142
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5912
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   240
         TabIndex        =   143
         Top             =   1875
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "trafico_fecha"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   3740
         TabIndex        =   59
         Top             =   500
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   112656387
         CurrentDate     =   44914
         MaxDate         =   55153
         MinDate         =   32874
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":592B
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   240
         TabIndex        =   144
         Top             =   1185
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5944
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2080
         TabIndex        =   145
         Top             =   500
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "codigo1"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Proyecto de Edificación"
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
         TabIndex        =   164
         Top             =   915
         Width           =   2130
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   163
         Top             =   1620
         Width           =   1560
      End
      Begin VB.Label Txt_campo2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36NO"
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         TabIndex        =   161
         Top             =   500
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Cite Contrato                Tipo Edificación      Fecha Registro"
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
         TabIndex        =   146
         Top             =   240
         Width           =   4965
      End
   End
   Begin VB.PictureBox Fra_datos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   8325
      Left            =   5640
      ScaleHeight     =   8265
      ScaleWidth      =   12360
      TabIndex        =   80
      Top             =   810
      Width           =   12420
      Begin VB.ComboBox dtc_codigo014 
         DataField       =   "pais_continente4"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "mw_solicitud_calculo_trafico_mod.frx":595D
         Left            =   9600
         List            =   "mw_solicitud_calculo_trafico_mod.frx":596D
         TabIndex        =   49
         Top             =   4440
         Width           =   2685
      End
      Begin VB.ComboBox dtc_codigo013 
         DataField       =   "pais_continente3"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "mw_solicitud_calculo_trafico_mod.frx":598E
         Left            =   6840
         List            =   "mw_solicitud_calculo_trafico_mod.frx":599E
         TabIndex        =   35
         Top             =   4440
         Width           =   2685
      End
      Begin VB.ComboBox dtc_codigo012 
         DataField       =   "pais_continente2"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "mw_solicitud_calculo_trafico_mod.frx":59BF
         Left            =   4080
         List            =   "mw_solicitud_calculo_trafico_mod.frx":59CF
         TabIndex        =   21
         Top             =   4440
         Width           =   2685
      End
      Begin VB.ComboBox dtc_codigo011 
         DataField       =   "pais_continente"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "mw_solicitud_calculo_trafico_mod.frx":59F0
         Left            =   1320
         List            =   "mw_solicitud_calculo_trafico_mod.frx":5A00
         TabIndex        =   7
         Top             =   4440
         Width           =   2685
      End
      Begin VB.TextBox Txt_aux24 
         DataField       =   "recorrido_codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9705
         TabIndex        =   43
         Text            =   "0"
         Top             =   1200
         Width           =   2145
      End
      Begin VB.TextBox Txt_aux23 
         DataField       =   "recorrido_codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6960
         TabIndex        =   29
         Text            =   "0"
         Top             =   1200
         Width           =   2145
      End
      Begin VB.TextBox Txt_aux22 
         DataField       =   "recorrido_codigo2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4185
         TabIndex        =   15
         Text            =   "0"
         Top             =   1200
         Width           =   2145
      End
      Begin VB.TextBox Txt_aux21 
         DataField       =   "recorrido_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Text            =   "0"
         Top             =   1200
         Width           =   2145
      End
      Begin MSDataListLib.DataCombo dtc_campo54 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5A21
         DataField       =   "tipo_puerta4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   11040
         TabIndex        =   153
         Top             =   3000
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483629
         ListField       =   "tipo_puerta_sigla"
         BoundColumn     =   "tipo_puerta"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_campo53 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5A3B
         DataField       =   "tipo_puerta3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   8280
         TabIndex        =   152
         Top             =   3000
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483629
         ListField       =   "tipo_puerta_sigla"
         BoundColumn     =   "tipo_puerta"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_campo52 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5A55
         DataField       =   "tipo_puerta2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5520
         TabIndex        =   151
         Top             =   3000
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   8421504
         ListField       =   "tipo_puerta_sigla"
         BoundColumn     =   "tipo_puerta"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_campo51 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5A6F
         DataField       =   "tipo_puerta"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2880
         TabIndex        =   60
         Top             =   3000
         Visible         =   0   'False
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   12632256
         ListField       =   "tipo_puerta_sigla"
         BoundColumn     =   "tipo_puerta"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux34 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5A89
         DataField       =   "pasajeros_codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   10680
         TabIndex        =   150
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483629
         ListField       =   "pasajeros_capacidad_km_w"
         BoundColumn     =   "pasajeros_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux33 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5AA4
         DataField       =   "pasajeros_codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7680
         TabIndex        =   149
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483629
         ListField       =   "pasajeros_capacidad_km_w"
         BoundColumn     =   "pasajeros_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux32 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5ABF
         DataField       =   "pasajeros_codigo2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5040
         TabIndex        =   148
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "pasajeros_capacidad_km_w"
         BoundColumn     =   "pasajeros_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux31 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5ADA
         DataField       =   "pasajeros_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2400
         TabIndex        =   147
         Top             =   1560
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483637
         ListField       =   "pasajeros_capacidad_km_w"
         BoundColumn     =   "pasajeros_codigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox Txt_campo21 
         DataField       =   "trafico_num_paradas"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Text            =   "0"
         Top             =   720
         Width           =   2145
      End
      Begin VB.TextBox Txt_campo23 
         DataField       =   "trafico_num_paradas3"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   6960
         TabIndex        =   28
         Text            =   "0"
         Top             =   720
         Width           =   2145
      End
      Begin VB.TextBox Txt_campo24 
         DataField       =   "trafico_num_paradas4"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   9705
         TabIndex        =   42
         Text            =   "0"
         Top             =   720
         Width           =   2145
      End
      Begin VB.TextBox Txt_campo22 
         DataField       =   "trafico_num_paradas2"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   4185
         TabIndex        =   14
         Text            =   "0"
         Top             =   720
         Width           =   2145
      End
      Begin VB.TextBox Txt_campo32 
         DataField       =   "trafico_nro_equipos2"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   4185
         TabIndex        =   17
         Text            =   "0"
         Top             =   2160
         Width           =   2145
      End
      Begin VB.TextBox Txt_campo34 
         DataField       =   "trafico_nro_equipos4"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   9705
         TabIndex        =   45
         Text            =   "0"
         Top             =   2160
         Width           =   2145
      End
      Begin VB.TextBox Txt_campo33 
         DataField       =   "trafico_nro_equipos3"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   6960
         TabIndex        =   31
         Text            =   "0"
         Top             =   2160
         Width           =   2145
      End
      Begin VB.TextBox Txt_campo31 
         DataField       =   "trafico_nro_equipos"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Text            =   "0"
         Top             =   2160
         Width           =   2145
      End
      Begin VB.TextBox Txt_campo42 
         DataField       =   "trafico_ancho_puerta2"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   4185
         TabIndex        =   20
         Text            =   "0"
         Top             =   3600
         Width           =   2145
      End
      Begin VB.TextBox Txt_campo44 
         DataField       =   "trafico_ancho_puerta4"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   9705
         TabIndex        =   48
         Text            =   "0"
         Top             =   3600
         Width           =   2145
      End
      Begin VB.TextBox Txt_campo43 
         DataField       =   "trafico_ancho_puerta3"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   6960
         TabIndex        =   34
         Text            =   "0"
         Top             =   3600
         Width           =   2145
      End
      Begin VB.TextBox Txt_campo41 
         DataField       =   "trafico_ancho_puerta"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Text            =   "0"
         Top             =   3600
         Width           =   2145
      End
      Begin MSDataListLib.DataCombo dtc_aux51 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5AF4
         DataField       =   "tipo_puerta"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2280
         TabIndex        =   81
         Top             =   3000
         Visible         =   0   'False
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   12632256
         ListField       =   "tiempo_apertura_cierre"
         BoundColumn     =   "tipo_puerta"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux52 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5B0E
         DataField       =   "tipo_puerta2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4785
         TabIndex        =   82
         Top             =   3000
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   12632256
         ListField       =   "tiempo_apertura_cierre"
         BoundColumn     =   "tipo_puerta"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux53 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5B28
         DataField       =   "tipo_puerta3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7395
         TabIndex        =   83
         Top             =   3000
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483629
         ListField       =   "tiempo_apertura_cierre"
         BoundColumn     =   "tipo_puerta"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux54 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5B42
         DataField       =   "tipo_puerta4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   10140
         TabIndex        =   84
         Top             =   3000
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483629
         ListField       =   "tiempo_apertura_cierre"
         BoundColumn     =   "tipo_puerta"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux41 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5B5C
         DataField       =   "vel_equipo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2040
         TabIndex        =   85
         Top             =   2520
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "vel_tiempo_asc_desacel"
         BoundColumn     =   "vel_equipo_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux42 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5B76
         DataField       =   "vel_equipo_codigo2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5025
         TabIndex        =   86
         Top             =   2520
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   12632256
         ListField       =   "vel_tiempo_asc_desacel"
         BoundColumn     =   "vel_equipo_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux43 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5B90
         DataField       =   "vel_equipo_codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7635
         TabIndex        =   87
         Top             =   2520
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483629
         ListField       =   "vel_tiempo_asc_desacel"
         BoundColumn     =   "vel_equipo_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux44 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5BAA
         DataField       =   "vel_equipo_codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   10500
         TabIndex        =   88
         Top             =   2520
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483629
         ListField       =   "vel_tiempo_asc_desacel"
         BoundColumn     =   "vel_equipo_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo04 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5BC4
         DataField       =   "condicion_cabina4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   89
         Top             =   6840
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_cabina"
         BoundColumn     =   "condicion_cabina"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo94 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5BDF
         DataField       =   "condicion_ventas4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   90
         Top             =   6360
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_ventas"
         BoundColumn     =   "condicion_ventas"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo84 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5BFA
         DataField       =   "sist_puerta4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   91
         Top             =   5880
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "sist_puerta"
         BoundColumn     =   "sist_puerta"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo74 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5C15
         DataField       =   "tecnologia_codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   92
         Top             =   5400
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tecnologia_codigo"
         BoundColumn     =   "tecnologia_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo54 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5C30
         DataField       =   "tipo_puerta4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9480
         TabIndex        =   93
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tipo_puerta"
         BoundColumn     =   "tipo_puerta"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo64 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5C4B
         DataField       =   "cabina_codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   94
         Top             =   4920
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cabina_codigo"
         BoundColumn     =   "cabina_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo44 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5C66
         DataField       =   "vel_equipo_codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9840
         TabIndex        =   95
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "vel_equipo_codigo"
         BoundColumn     =   "vel_equipo_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo34 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5C81
         DataField       =   "pasajeros_codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9960
         TabIndex        =   96
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "pasajeros_codigo"
         BoundColumn     =   "pasajeros_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo63 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5C9C
         DataField       =   "cabina_codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   97
         Top             =   4920
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cabina_codigo"
         BoundColumn     =   "cabina_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo73 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5CB7
         DataField       =   "tecnologia_codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   98
         Top             =   5400
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tecnologia_codigo"
         BoundColumn     =   "tecnologia_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo83 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5CD2
         DataField       =   "sist_puerta3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   99
         Top             =   5880
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "sist_puerta"
         BoundColumn     =   "sist_puerta"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo93 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5CED
         DataField       =   "condicion_ventas3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   100
         Top             =   6360
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_ventas"
         BoundColumn     =   "condicion_ventas"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo03 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5D08
         DataField       =   "condicion_cabina3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   101
         Top             =   6840
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_cabina"
         BoundColumn     =   "condicion_cabina"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo62 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5D23
         DataField       =   "cabina_codigo2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   102
         Top             =   4920
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cabina_codigo"
         BoundColumn     =   "cabina_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo72 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5D3E
         DataField       =   "tecnologia_codigo2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   103
         Top             =   5400
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tecnologia_codigo"
         BoundColumn     =   "tecnologia_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo82 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5D59
         DataField       =   "sist_puerta2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   104
         Top             =   5880
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "sist_puerta"
         BoundColumn     =   "sist_puerta"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo92 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5D74
         DataField       =   "condicion_ventas2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   105
         Top             =   6360
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_ventas"
         BoundColumn     =   "condicion_ventas"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo02 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5D8F
         DataField       =   "condicion_cabina2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   106
         Top             =   6840
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_cabina"
         BoundColumn     =   "condicion_cabina"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo33 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5DAA
         DataField       =   "pasajeros_codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6960
         TabIndex        =   107
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "pasajeros_codigo"
         BoundColumn     =   "pasajeros_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo43 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5DC5
         DataField       =   "vel_equipo_codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6960
         TabIndex        =   108
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "vel_equipo_codigo"
         BoundColumn     =   "vel_equipo_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo53 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5DE0
         DataField       =   "tipo_puerta3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6720
         TabIndex        =   109
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tipo_puerta"
         BoundColumn     =   "tipo_puerta"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo32 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5DFB
         DataField       =   "pasajeros_codigo2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4320
         TabIndex        =   110
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "pasajeros_codigo"
         BoundColumn     =   "pasajeros_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo42 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5E16
         DataField       =   "vel_equipo_codigo2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4200
         TabIndex        =   111
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "vel_equipo_codigo"
         BoundColumn     =   "vel_equipo_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo52 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5E31
         DataField       =   "tipo_puerta2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   112
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tipo_puerta"
         BoundColumn     =   "tipo_puerta"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc31 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5E4C
         DataField       =   "pasajeros_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   1680
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "pasajeros_descripcion"
         BoundColumn     =   "pasajeros_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc32 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5E66
         DataField       =   "pasajeros_codigo2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4200
         TabIndex        =   16
         Top             =   1680
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "pasajeros_descripcion"
         BoundColumn     =   "pasajeros_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc33 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5E80
         DataField       =   "pasajeros_codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6960
         TabIndex        =   30
         Top             =   1680
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "pasajeros_descripcion"
         BoundColumn     =   "pasajeros_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc34 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5E9A
         DataField       =   "pasajeros_codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9705
         TabIndex        =   44
         Top             =   1680
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "pasajeros_descripcion"
         BoundColumn     =   "pasajeros_codigo"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_desc41 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5EB4
         DataField       =   "vel_equipo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   2640
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "vel_equipo_descripcion"
         BoundColumn     =   "vel_equipo_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc42 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5ECE
         DataField       =   "vel_equipo_codigo2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4185
         TabIndex        =   18
         Top             =   2640
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "vel_equipo_descripcion"
         BoundColumn     =   "vel_equipo_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc43 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5EE8
         DataField       =   "vel_equipo_codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6960
         TabIndex        =   32
         Top             =   2640
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "vel_equipo_descripcion"
         BoundColumn     =   "vel_equipo_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc44 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5F02
         DataField       =   "vel_equipo_codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9705
         TabIndex        =   46
         Top             =   2640
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "vel_equipo_descripcion"
         BoundColumn     =   "vel_equipo_codigo"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_desc51 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5F1C
         DataField       =   "tipo_puerta"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   3120
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tipo_puerta_descripcion"
         BoundColumn     =   "tipo_puerta"
         Text            =   "Todos"
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
      Begin MSDataListLib.DataCombo dtc_desc52 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5F36
         DataField       =   "tipo_puerta2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4185
         TabIndex        =   19
         Top             =   3120
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tipo_puerta_descripcion"
         BoundColumn     =   "tipo_puerta"
         Text            =   "Todos"
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
      Begin MSDataListLib.DataCombo dtc_desc53 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5F50
         DataField       =   "tipo_puerta3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6960
         TabIndex        =   33
         Top             =   3120
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tipo_puerta_descripcion"
         BoundColumn     =   "tipo_puerta"
         Text            =   "Todos"
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
      Begin MSDataListLib.DataCombo dtc_desc54 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5F6A
         DataField       =   "tipo_puerta4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9705
         TabIndex        =   47
         Top             =   3120
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tipo_puerta_descripcion"
         BoundColumn     =   "tipo_puerta"
         Text            =   "0"
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
      Begin MSDataListLib.DataCombo dtc_desc61 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5F84
         DataField       =   "cabina_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   4935
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cabina_descripcion"
         BoundColumn     =   "cabina_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc62 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5F9E
         DataField       =   "cabina_codigo2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   22
         Top             =   4935
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cabina_descripcion"
         BoundColumn     =   "cabina_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc63 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5FB8
         DataField       =   "cabina_codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   36
         Top             =   4935
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cabina_descripcion"
         BoundColumn     =   "cabina_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc64 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5FD2
         DataField       =   "cabina_codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   50
         Top             =   4935
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cabina_descripcion"
         BoundColumn     =   "cabina_codigo"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_desc71 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":5FEC
         DataField       =   "tecnologia_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   5445
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tecnologia_descripcion"
         BoundColumn     =   "tecnologia_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc72 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":6006
         DataField       =   "tecnologia_codigo2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   23
         Top             =   5445
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tecnologia_descripcion"
         BoundColumn     =   "tecnologia_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc73 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":6020
         DataField       =   "tecnologia_codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   37
         Top             =   5445
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tecnologia_descripcion"
         BoundColumn     =   "tecnologia_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc74 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":603A
         DataField       =   "tecnologia_codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   51
         Top             =   5445
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tecnologia_descripcion"
         BoundColumn     =   "tecnologia_codigo"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_desc81 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":6054
         DataField       =   "sist_puerta"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   5940
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "sist_puerta_descripcion"
         BoundColumn     =   "sist_puerta"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc82 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":606E
         DataField       =   "sist_puerta2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   24
         Top             =   5940
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "sist_puerta_descripcion"
         BoundColumn     =   "sist_puerta"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc83 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":6088
         DataField       =   "sist_puerta3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   38
         Top             =   5940
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "sist_puerta_descripcion"
         BoundColumn     =   "sist_puerta"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc84 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":60A2
         DataField       =   "sist_puerta4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   52
         Top             =   5940
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "sist_puerta_descripcion"
         BoundColumn     =   "sist_puerta"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_desc91 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":60BC
         DataField       =   "condicion_ventas"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   6435
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_ventas_descripcion"
         BoundColumn     =   "condicion_ventas"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc92 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":60D6
         DataField       =   "condicion_ventas2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   25
         Top             =   6435
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_ventas_descripcion"
         BoundColumn     =   "condicion_ventas"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc93 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":60F0
         DataField       =   "condicion_ventas3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   39
         Top             =   6435
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_ventas_descripcion"
         BoundColumn     =   "condicion_ventas"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc94 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":610A
         DataField       =   "condicion_ventas4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   53
         Top             =   6435
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_ventas_descripcion"
         BoundColumn     =   "condicion_ventas"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_desc01 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":6124
         DataField       =   "condicion_cabina"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   6945
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_cabina_descripcion"
         BoundColumn     =   "condicion_cabina"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc02 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":613E
         DataField       =   "condicion_cabina2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   26
         Top             =   6945
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_cabina_descripcion"
         BoundColumn     =   "condicion_cabina"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc03 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":6158
         DataField       =   "condicion_cabina3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   40
         Top             =   6945
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_cabina_descripcion"
         BoundColumn     =   "condicion_cabina"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc04 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":6172
         DataField       =   "condicion_cabina4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   55
         Top             =   6945
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_cabina_descripcion"
         BoundColumn     =   "condicion_cabina"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_codigo31 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":618C
         DataField       =   "pasajeros_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1680
         TabIndex        =   113
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "pasajeros_codigo"
         BoundColumn     =   "pasajeros_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo41 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":61A7
         DataField       =   "vel_equipo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   114
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "vel_equipo_codigo"
         BoundColumn     =   "vel_equipo_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo51 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":61C2
         DataField       =   "tipo_puerta"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   115
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tipo_puerta"
         BoundColumn     =   "tipo_puerta"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo61 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":61DD
         DataField       =   "cabina_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   116
         Top             =   4800
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cabina_codigo"
         BoundColumn     =   "cabina_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo71 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":61F8
         DataField       =   "tecnologia_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   117
         Top             =   5280
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "tecnologia_codigo"
         BoundColumn     =   "tecnologia_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo81 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":6213
         DataField       =   "sist_puerta"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   118
         Top             =   5760
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "sist_puerta"
         BoundColumn     =   "sist_puerta"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo91 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":622E
         DataField       =   "condicion_ventas"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   119
         Top             =   6240
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_ventas"
         BoundColumn     =   "condicion_ventas"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo01 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":6249
         DataField       =   "condicion_cabina"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   120
         Top             =   6720
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "condicion_cabina"
         BoundColumn     =   "condicion_cabina"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo11 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":6264
         DataField       =   "ctrlmaq_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   156
         Top             =   7200
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "ctrlmaq_codigo"
         BoundColumn     =   "ctrlmaq_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo12 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":627F
         DataField       =   "ctrlmaq_codigo2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   157
         Top             =   7320
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "ctrlmaq_codigo"
         BoundColumn     =   "ctrlmaq_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo13 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":629A
         DataField       =   "ctrlmaq_codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   158
         Top             =   7320
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "ctrlmaq_codigo"
         BoundColumn     =   "ctrlmaq_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo14 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":62B5
         DataField       =   "ctrlmaq_codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   159
         Top             =   7320
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "ctrlmaq_codigo"
         BoundColumn     =   "ctrlmaq_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc11 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":62D0
         DataField       =   "ctrlmaq_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   7440
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "ctrlmaq_descripcion"
         BoundColumn     =   "ctrlmaq_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc12 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":62EA
         DataField       =   "ctrlmaq_codigo2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   27
         Top             =   7440
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "ctrlmaq_descripcion"
         BoundColumn     =   "ctrlmaq_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc13 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":6304
         DataField       =   "ctrlmaq_codigo3"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   41
         Top             =   7440
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "ctrlmaq_descripcion"
         BoundColumn     =   "ctrlmaq_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc14 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":631E
         DataField       =   "ctrlmaq_codigo4"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   56
         Top             =   7440
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "ctrlmaq_descripcion"
         BoundColumn     =   "ctrlmaq_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_valor41 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":6338
         DataField       =   "vel_equipo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2880
         TabIndex        =   165
         Top             =   2520
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   12632256
         ListField       =   "vel_equipo_m_s"
         BoundColumn     =   "vel_equipo_codigo"
         Text            =   "Todos"
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFC0&
         X1              =   9555
         X2              =   9555
         Y1              =   -15
         Y2              =   8280
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFC0&
         X1              =   6795
         X2              =   6795
         Y1              =   -135
         Y2              =   8280
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "B.-  PARAMETROS COMPLEMENTARIOS (TABLA PRODUCTOS)"
         DataSource      =   "adoLista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   122
         Top             =   4035
         Width           =   5895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFC0&
         X1              =   4035
         X2              =   4035
         Y1              =   -135
         Y2              =   8280
      End
      Begin VB.Label lbl_campo14 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Industria Continente"
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
         Height          =   420
         Left            =   120
         TabIndex        =   166
         Top             =   4380
         Width           =   1155
      End
      Begin VB.Label lbl_ttt 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   160
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lbl_campo13 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Control / Máquinas"
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
         Height          =   420
         Left            =   120
         TabIndex        =   155
         Top             =   7380
         Width           =   1065
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"mw_solicitud_calculo_trafico_mod.frx":6352
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
         Height          =   270
         Left            =   1560
         TabIndex        =   135
         Top             =   30
         Width           =   10305
      End
      Begin VB.Label lbl_campo8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estética de Cabina"
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
         Height          =   420
         Left            =   120
         TabIndex        =   134
         Top             =   4875
         Width           =   1095
      End
      Begin VB.Label lbl_campo10 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sist.Operador de Puertas"
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
         Height          =   420
         Left            =   120
         TabIndex        =   133
         Top             =   5880
         Width           =   1155
      End
      Begin VB.Label lbl_campo9 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Línea / Tecnología"
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
         Height          =   420
         Left            =   120
         TabIndex        =   132
         Top             =   5385
         Width           =   1080
      End
      Begin VB.Label lbl_campo12 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Condicion Cabina"
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
         Height          =   420
         Left            =   120
         TabIndex        =   131
         Top             =   6885
         Width           =   1110
      End
      Begin VB.Label lbl_campo11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Condicion Ventas"
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
         Height          =   420
         Left            =   120
         TabIndex        =   130
         Top             =   6380
         Width           =   1095
      End
      Begin VB.Label lbl_campo7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho Puerta (mm)"
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
         Height          =   600
         Left            =   120
         TabIndex        =   129
         Top             =   3560
         Width           =   1215
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº de Equipos"
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
         Height          =   195
         Left            =   120
         TabIndex        =   128
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lbl_campo6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Puerta Piso (Apertura)"
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
         Height          =   480
         Left            =   120
         TabIndex        =   127
         Top             =   3020
         Width           =   1755
      End
      Begin VB.Label lbl_campo5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidad (m/s)"
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
         Height          =   195
         Left            =   120
         TabIndex        =   126
         Top             =   2640
         Width           =   1350
      End
      Begin VB.Label lbl_campo2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Recorrido (mt)"
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
         Height          =   195
         Left            =   120
         TabIndex        =   125
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Pasajeros"
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
         Height          =   195
         Left            =   120
         TabIndex        =   124
         Top             =   1680
         Width           =   1365
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº de Paradas "
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
         Height          =   195
         Left            =   120
         TabIndex        =   123
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "A.- PARAMETROS DE CALCULO "
         DataSource      =   "adoLista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   121
         Top             =   375
         Width           =   3045
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8400
      Left            =   120
      TabIndex        =   61
      Top             =   720
      Width           =   5535
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
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   3360
         TabIndex        =   58
         Top             =   2595
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
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   1080
         TabIndex        =   57
         Top             =   2595
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "mw_solicitud_calculo_trafico_mod.frx":63E6
         Height          =   2250
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   3969
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "solicitud_codigo"
            Caption         =   "No.Tramite"
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
            Caption         =   "U.Ejecutora"
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
            DataField       =   "edif_codigo"
            Caption         =   "Edificio"
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
            DataField       =   "estado_codigo_verif"
            Caption         =   "Verificado"
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
            DataField       =   "trafico_fecha"
            Caption         =   "Fecha_Registro"
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
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   2520
         Width           =   5265
         _ExtentX        =   9287
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
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "H. RESULTADOS OBTENIDOS (TOTALES)"
         DataSource      =   "adoLista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   67
         Top             =   5400
         Width           =   3855
      End
      Begin VB.Label lbl_campoh43 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "h_capacidad_trafico_result"
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
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   4080
         TabIndex        =   79
         Top             =   7560
         Width           =   1320
      End
      Begin VB.Label lbl_campoh42 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "h_capacidad_trafico_parametro"
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
         Height          =   315
         Left            =   2940
         TabIndex        =   78
         Top             =   7560
         Width           =   1005
      End
      Begin VB.Label lbl_campoh41 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "h_capacidad_trafico"
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
         Height          =   315
         Left            =   1800
         TabIndex        =   77
         Top             =   7560
         Width           =   1005
      End
      Begin VB.Label lbl_campoh33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "h_intervalo_trafico_result"
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
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   4080
         TabIndex        =   76
         Top             =   7080
         Width           =   1320
      End
      Begin VB.Label lbl_campoh32 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "h_intervalo_trafico_parametro"
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
         Height          =   315
         Left            =   2940
         TabIndex        =   75
         Top             =   7080
         Width           =   1005
      End
      Begin VB.Label lbl_campoh31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "h_intervalo_trafico"
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
         Height          =   315
         Left            =   1800
         TabIndex        =   74
         Top             =   7080
         Width           =   1005
      End
      Begin VB.Label lbl_campoh23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "h_partida_por_hora_result"
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
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   4080
         TabIndex        =   73
         Top             =   6600
         Width           =   1320
      End
      Begin VB.Label lbl_campoh22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "h_partida_por_hora_parametro"
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
         Height          =   315
         Left            =   2940
         TabIndex        =   72
         Top             =   6600
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lbl_campoh21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "h_partidas_por_hora"
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
         Height          =   315
         Left            =   1800
         TabIndex        =   71
         Top             =   6600
         Width           =   1005
      End
      Begin VB.Label lbl_campoh13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "h_nro_total_equipos_result"
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
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   4080
         TabIndex        =   70
         Top             =   6120
         Width           =   1320
      End
      Begin VB.Label lbl_campoh12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "h_nro_total_equipos_parametro"
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
         Height          =   315
         Left            =   2940
         TabIndex        =   69
         Top             =   6120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lbl_campoh11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "h_nro_total_equipos"
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
         Height          =   315
         Left            =   1800
         TabIndex        =   68
         Top             =   6120
         Width           =   1005
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Total Equipos"
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
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   6165
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Partidas por Hora"
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
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   6645
         Width           =   1500
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Intervalo Trafico"
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
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   7125
         Width           =   1425
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Capacidad.Trafico"
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
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   7605
         Width           =   1575
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Arreglos   Parámetros      Resultados   "
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
         Height          =   225
         Left            =   1620
         TabIndex        =   62
         Top             =   5745
         Width           =   3825
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
      Top             =   8400
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
   Begin MSAdodcLib.Adodc Ado_datos22 
      Height          =   330
      Left            =   6480
      Top             =   8400
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos22"
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
   Begin MSAdodcLib.Adodc Ado_datos21 
      Height          =   330
      Left            =   2160
      Top             =   8400
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos21"
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
   Begin MSAdodcLib.Adodc Ado_datos23 
      Height          =   330
      Left            =   8640
      Top             =   8400
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos23"
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
      Left            =   4320
      Top             =   8400
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
   Begin MSAdodcLib.Adodc Ado_datos41 
      Height          =   330
      Left            =   2160
      Top             =   8760
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Ado_datos41"
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
   Begin MSAdodcLib.Adodc Ado_datos42 
      Height          =   330
      Left            =   4320
      Top             =   8760
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Ado_datos42"
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
   Begin MSAdodcLib.Adodc Ado_datos51 
      Height          =   330
      Left            =   8640
      Top             =   8760
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos51"
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
   Begin MSAdodcLib.Adodc Ado_datos43 
      Height          =   330
      Left            =   6480
      Top             =   8760
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos43"
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
   Begin MSAdodcLib.Adodc Ado_datos52 
      Height          =   330
      Left            =   10800
      Top             =   8760
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos52"
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
   Begin MSAdodcLib.Adodc Ado_datos32 
      Height          =   330
      Left            =   12960
      Top             =   8400
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos32"
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
   Begin MSAdodcLib.Adodc Ado_datos53 
      Height          =   330
      Left            =   12960
      Top             =   8760
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos53"
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
   Begin MSAdodcLib.Adodc Ado_datos61 
      Height          =   330
      Left            =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos61"
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
   Begin MSAdodcLib.Adodc Ado_datos62 
      Height          =   330
      Left            =   2160
      Top             =   9120
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos62"
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
   Begin MSAdodcLib.Adodc Ado_datos63 
      Height          =   330
      Left            =   4320
      Top             =   9120
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos63"
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
   Begin MSAdodcLib.Adodc Ado_datos71 
      Height          =   330
      Left            =   6480
      Top             =   9120
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos71"
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
   Begin MSAdodcLib.Adodc Ado_datos82 
      Height          =   330
      Left            =   0
      Top             =   9480
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos82"
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
   Begin MSAdodcLib.Adodc Ado_datos72 
      Height          =   330
      Left            =   8640
      Top             =   9120
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos72"
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
   Begin MSAdodcLib.Adodc Ado_datos73 
      Height          =   330
      Left            =   10800
      Top             =   9120
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos73"
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
   Begin MSAdodcLib.Adodc Ado_datos81 
      Height          =   330
      Left            =   12960
      Top             =   9120
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos81"
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
   Begin MSAdodcLib.Adodc Ado_datos83 
      Height          =   330
      Left            =   2160
      Top             =   9480
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos83"
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
   Begin MSAdodcLib.Adodc Ado_datos91 
      Height          =   330
      Left            =   4320
      Top             =   9480
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos91"
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
   Begin MSAdodcLib.Adodc Ado_datos92 
      Height          =   330
      Left            =   6480
      Top             =   9480
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos92"
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
   Begin MSAdodcLib.Adodc Ado_datos93 
      Height          =   330
      Left            =   8640
      Top             =   9480
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos93"
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
   Begin MSAdodcLib.Adodc Ado_datos01 
      Height          =   330
      Left            =   10800
      Top             =   9480
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos01"
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
   Begin MSAdodcLib.Adodc Ado_datos03 
      Height          =   330
      Left            =   0
      Top             =   9840
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos03"
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
   Begin MSAdodcLib.Adodc Ado_datos02 
      Height          =   330
      Left            =   12960
      Top             =   9480
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos02"
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
   Begin MSAdodcLib.Adodc Ado_datos31 
      Height          =   330
      Left            =   10800
      Top             =   8400
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos31"
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
   Begin MSAdodcLib.Adodc Ado_datos33 
      Height          =   330
      Left            =   0
      Top             =   8760
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos33"
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
      Left            =   2160
      Top             =   9720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
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
   Begin MSAdodcLib.Adodc Ado_datos24 
      Height          =   330
      Left            =   2160
      Top             =   9840
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos24"
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
   Begin MSAdodcLib.Adodc Ado_datos34 
      Height          =   330
      Left            =   4320
      Top             =   9840
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos34"
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
   Begin MSAdodcLib.Adodc Ado_datos44 
      Height          =   330
      Left            =   6480
      Top             =   9840
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos44"
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
   Begin MSAdodcLib.Adodc Ado_datos54 
      Height          =   330
      Left            =   8640
      Top             =   9840
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos54"
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
   Begin MSAdodcLib.Adodc Ado_datos64 
      Height          =   330
      Left            =   10800
      Top             =   9840
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos64"
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
   Begin MSAdodcLib.Adodc Ado_datos84 
      Height          =   330
      Left            =   0
      Top             =   10200
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos84"
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
   Begin MSAdodcLib.Adodc Ado_datos94 
      Height          =   330
      Left            =   2160
      Top             =   10200
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos94"
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
   Begin MSAdodcLib.Adodc Ado_datos04 
      Height          =   330
      Left            =   4320
      Top             =   10200
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos04"
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
   Begin MSAdodcLib.Adodc Ado_datos74 
      Height          =   330
      Left            =   12960
      Top             =   9840
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos74"
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
      Left            =   6480
      Top             =   10200
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Left            =   8640
      Top             =   10200
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
   Begin MSAdodcLib.Adodc Ado_datos13 
      Height          =   330
      Left            =   10800
      Top             =   10200
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
   Begin MSAdodcLib.Adodc Ado_datos14 
      Height          =   330
      Left            =   12960
      Top             =   10200
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos14"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   -1560
      Top             =   23640
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Ado_datos23"
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
   Begin MSAdodcLib.Adodc Ado_datos05 
      Height          =   330
      Left            =   0
      Top             =   10560
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos05"
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
      ScaleWidth      =   20280
      TabIndex        =   178
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1560
         Picture         =   "mw_solicitud_calculo_trafico_mod.frx":63FE
         ScaleHeight     =   615
         ScaleWidth      =   1305
         TabIndex        =   180
         Top             =   0
         Width           =   1305
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2955
         Picture         =   "mw_solicitud_calculo_trafico_mod.frx":6BD4
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   179
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO2"
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
         Left            =   12600
         TabIndex        =   181
         Top             =   195
         Width           =   1035
      End
   End
   Begin MSAdodcLib.Adodc Ado_detalle2 
      Height          =   330
      Left            =   2160
      Top             =   10560
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
      ConnectStringType=   3
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
End
Attribute VB_Name = "mw_solicitud_calculo_trafico_mod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos21 As New ADODB.Recordset
Dim rs_datos22 As New ADODB.Recordset
Dim rs_datos23 As New ADODB.Recordset
Dim rs_datos24 As New ADODB.Recordset
Dim rs_datos31 As New ADODB.Recordset
Dim rs_datos32 As New ADODB.Recordset
Dim rs_datos33 As New ADODB.Recordset
Dim rs_datos34 As New ADODB.Recordset
Dim rs_datos41 As New ADODB.Recordset
Dim rs_datos42 As New ADODB.Recordset
Dim rs_datos43 As New ADODB.Recordset
Dim rs_datos44 As New ADODB.Recordset
Dim rs_datos51 As New ADODB.Recordset
Dim rs_datos52 As New ADODB.Recordset
Dim rs_datos53 As New ADODB.Recordset
Dim rs_datos54 As New ADODB.Recordset
Dim rs_datos61 As New ADODB.Recordset
Dim rs_datos62 As New ADODB.Recordset
Dim rs_datos63 As New ADODB.Recordset
Dim rs_datos64 As New ADODB.Recordset
Dim rs_datos71 As New ADODB.Recordset
Dim rs_datos72 As New ADODB.Recordset
Dim rs_datos73 As New ADODB.Recordset
Dim rs_datos74 As New ADODB.Recordset
Dim rs_datos81 As New ADODB.Recordset
Dim rs_datos82 As New ADODB.Recordset
Dim rs_datos83 As New ADODB.Recordset
Dim rs_datos84 As New ADODB.Recordset
Dim rs_datos91 As New ADODB.Recordset
Dim rs_datos92 As New ADODB.Recordset
Dim rs_datos93 As New ADODB.Recordset
Dim rs_datos94 As New ADODB.Recordset
Dim rs_datos01 As New ADODB.Recordset
Dim rs_datos02 As New ADODB.Recordset
Dim rs_datos03 As New ADODB.Recordset
Dim rs_datos04 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset
Dim rs_datos12 As New ADODB.Recordset
Dim rs_datos13 As New ADODB.Recordset
Dim rs_datos14 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset

Dim rstbeneficiario As New ADODB.Recordset
Dim rst_ben, rsNada As New ADODB.Recordset
Dim RsTmp As New ADODB.Recordset
Dim rs_aux1, rs_aux2 As New ADODB.Recordset
Dim rs_aux3, rs_aux4  As New ADODB.Recordset
Dim rs_aux5, rs_aux6 As New ADODB.Recordset

'Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

'OTROS
Dim VAR_MOD, VAR_MOD1, VAR_MOD2 As String
Dim SQL_FOR As String
Dim sql As String
Dim swnuevo As String
Dim sino As String
Dim NombreCarpeta, e As String
Dim VAR_DA, VAR_DPTO As String
Dim parametro, VAR_UORIGEN  As String
Dim var_cod, VAR_COD2, VAR_COD3 As String
Dim VAR_VAL, VAR_NO1, VAR_NO2, VAR_NO3 As String
Dim VAR_SW As String
Dim VAR_CONTI As String

Dim imag2 As Long

Dim i As Integer
Dim VAR_CONT1, VAR_COD4 As Integer

Dim VAR_AUX As Double
Dim var_campoc31, var_campoc32, var_campoc33, var_campoc34 As Double
Dim var_campod11, var_campod12, var_campod13, var_campod14 As Double
Dim var_campoe11, var_campoe12, var_campoe13, var_campoe14 As Double
Dim var_campoe21, var_campoe22, var_campoe23, var_campoe24 As Double
Dim var_campoe31, var_campoe32, var_campoe33, var_campoe34 As Double
Dim var_campoe41, var_campoe42, var_campoe43, var_campoe44 As Double
Dim var_campog11, var_campog12, var_campog13, var_campog14 As Double
Dim var_campog21, var_campog22, var_campog23, var_campog24 As Double

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
     '<-- Inicio                Identificación del Cliente                Fin -->
   If Ado_datos.Recordset.RecordCount > 0 Then
     If Ado_datos.Recordset!estado_codigo = "REG" Then
'        Call BtnModificar_Click
        BtnModificar.Visible = True
        BtnAprobar.Visible = True
'        If Ado_datos.Recordset!estado_codigo_verif = "APR" Then
'            'BtnBuscar.Visible = False
'            BtnAprobar.Visible = False
'        Else
'            'BtnBuscar.Visible = True
'            BtnAprobar.Visible = True
'        End If
     Else
        'BtnBuscar.Visible = False
        BtnModificar.Visible = False
        BtnAprobar.Visible = False
     End If
   End If
End Sub

Private Sub BtnAñadir_Click()
    GlUnidad = Ado_datos.Recordset!unidad_codigo
    GlSolicitud = Ado_datos.Recordset!solicitud_codigo
    mw_solicitud_calculo_trafico_mod_DET.Show
End Sub

Private Sub BtnAprobar_Click()
'   On Error GoTo UpdateErr
'   If rs_datos!estado_codigo_verif = "APR" And Val(Ado_datos.Recordset!trafico_num_paradas) > 0 Then
'        db.Execute "Update ao_solicitud Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'        db.Execute "Update ao_solicitud_cotiza_venta Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'        db.Execute "Update ao_negociacion_cabecera Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and negocia_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'
'        rs_datos!estado_codigo = "APR"
'        rs_datos!fecha_registro = Date
'        rs_datos!usr_codigo = glusuario
'        rs_datos.UpdateBatch adAffectAll
'   Else
'       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene detalle ...", vbExclamation, "Validación de Registro"
'   End If
'   Exit Sub
'UpdateErr:
'  MsgBox Err.Description

On Error GoTo UpdateErr
   '1 if
   If rs_datos!estado_codigo = "REG" And Txt_campo21 <> "" And Txt_aux21 <> "" And dtc_aux41.Text <> "" And dtc_codigo011.Text <> "NN" Then
   'If rs_datos!estado_codigo = "REG" And Val(Ado_datos.Recordset!trafico_num_paradas) > 0 And dtc_aux41.Text <> "" And dtc_codigo011.Text <> "NN" Then
   'If rs_datos!estado_codigo = "REG" And rs_datos!estado_codigo_verif = "REG" And Val(Ado_datos.Recordset!trafico_num_paradas) > 0 And dtc_aux41.Text <> "" And dtc_codigo011.Text <> "NN" Then
      sino = MsgBox("Está Seguro de VERIFICAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      '2 if
      If sino = vbYes Then
        GlSolicitud = rs_datos!solicitud_codigo
        VAR_COD4 = rs_datos!solicitud_codigo
        If Left(Txt_campo2, 4) = "36NO" Or Left(Txt_campo2, 4) = "OA36" Then
            rs_datos!unidad_codigo_ant = Trim(Txt_campo2)
        Else
            If VAR_COD4 < 10 Then
               rs_datos!unidad_codigo_ant = parametro + "-00000" + Trim(txt_codigo)
            End If
            If VAR_COD4 > 9 And VAR_COD4 < 100 Then
               rs_datos!unidad_codigo_ant = parametro + "-0000" + Trim(txt_codigo)
            End If
            If VAR_COD4 > 99 And VAR_COD4 < 1000 Then
               rs_datos!unidad_codigo_ant = parametro + "-000" + Trim(txt_codigo)
            End If
            If VAR_COD4 > 999 And VAR_COD4 < 10000 Then
               rs_datos!unidad_codigo_ant = parametro + "-00" + Trim(txt_codigo)
            End If
            If VAR_COD4 > 9999 And VAR_COD4 < 100000 Then
               rs_datos!unidad_codigo_ant = parametro + "-0" + Trim(txt_codigo)
            End If
            If VAR_COD4 > 99999 Then
               rs_datos!unidad_codigo_ant = parametro + "-" + Trim(txt_codigo)
            End If
        End If
        
        Set rs_aux5 = New ADODB.Recordset
        If rs_aux5.State = 1 Then rs_aux5.Close
        rs_aux5.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & "     ", db, adOpenKeyset, adLockOptimistic
'        SQL_FOR = "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & "  "
        '3 if
        If rs_aux5.RecordCount > 0 Then
           sino = MsgBox("La Cotización ya existe, desea volver a Generarla ? ...", vbYesNo + vbQuestion, "Atención ...")
           '4 if
           If sino = vbYes Then
              'OJO BORRAR ao_solicitud_costos
             Select Case dtc_codigo011.Text
                Case "AMERICA"
                    'Brasil 1   - AMERICA
                    VAR_CONTI = "AMERICA"
                Case "ASIA"
                    'Hypex  1   - ASIA
                    VAR_CONTI = "ASIA"
                Case "EUROPA"
                    'Xizi   XO  1   -   ESPAÑA
                    VAR_CONTI = "EUROPA"
                Case Else
                    VAR_CONTI = "AMERICA"
             End Select
             db.Execute "DELETE ao_solicitud_cotiza_modelo where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND trafico_codigo = " & CDbl(Txt_campo1) & "   "
             db.Execute "DELETE ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND trafico_codigo = " & CDbl(Txt_campo1) & "   "
              'corrprog = 0
             'Call GRABA_ARREGLO1
             Call GRABA_ARREGLO_1
             '6 if
             If Txt_campo22.Text <> "" And Txt_aux22 <> "" And dtc_codigo012.Text <> "NN" Then             'ARREGLO 2
                'Call GRABA_ARREGLO2
                Call GRABA_ARREGLO_2
             End If
             '6 end
             '7 if
             If Txt_campo23.Text <> "" And Txt_aux23 <> "" And dtc_codigo013.Text <> "NN" Then                   'ARREGLO 3
                'Call GRABA_ARREGLO3
                Call GRABA_ARREGLO_3
             End If
             '7 end
             '8 if
             If Txt_campo24.Text <> "" And Txt_aux24 <> "" And dtc_codigo014.Text <> "NN" Then                  'ARREGLO 4
                'Call GRABA_ARREGLO4
                Call GRABA_ARREGLO_4
             End If
             '8 end

           Else     '4 Else
             i = 1
             While (CDbl(Txt_campo31.Text) >= i)         'ARREGLO 1
                  Set rs_aux2 = New ADODB.Recordset
                  If rs_aux2.State = 1 Then rs_aux2.Close
                  rs_aux2.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 1 and cotiza_codigo = " & i & " ", db, adOpenStatic
                  If rs_aux2.RecordCount > 0 Then
                     db.Execute "Update ao_solicitud_cotiza_venta Set pais_continente = '" & Ado_datos.Recordset!pais_continente & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 1 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_venta Set modelo_codigo = '" & Ado_datos.Recordset!modelo_codigo & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 1 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_venta Set fecha_registro = '" & Date & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 1 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_venta Set usr_codigo = '" & glusuario & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 1 and cotiza_codigo = " & i & "  "
                      'rs_aux2!pais_continente = Ado_datos.Recordset!pais_continente
                      'rs_aux2!modelo_codigo = Ado_datos.Recordset!modelo_codigo
                      'rs_aux2!arreglo = 1     'arreglo1
                      'rs_aux2!fecha_registro = Date
                      'rs_aux2!usr_codigo = glusuario
                      'rs_aux2.Update
                  End If
                  Set rs_aux4 = New ADODB.Recordset
                  If rs_aux4.State = 1 Then rs_aux4.Close
                  rs_aux4.Open "Select * from ao_solicitud_cotiza_modelo where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 1 and cotiza_codigo = " & i & " ", db, adOpenStatic
                  If rs_aux4.RecordCount > 0 Then
                      'rs_aux4!pais_continente = Ado_datos.Recordset!pais_continente
                      'rs_aux4!modelo_codigo = Ado_datos.Recordset!modelo_codigo
                      'rs_aux4!arreglo = 1     'arreglo1
                      'rs_aux4!fecha_registro = Date
                      'rs_aux4!usr_codigo = glusuario
                      'rs_aux4.Update
                     db.Execute "Update ao_solicitud_cotiza_modelo Set pais_continente = '" & Ado_datos.Recordset!pais_continente & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 1 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_modelo Set modelo_codigo = '" & Ado_datos.Recordset!modelo_codigo & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 1 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_modelo Set fecha_registro = '" & Date & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 1 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_modelo Set usr_codigo = '" & glusuario & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 1 and cotiza_codigo = " & i & "  "
                  End If
             i = i + 1
             Wend
             '6 if
             If Txt_campo22.Text <> "" Then               'ARREGLO 2
                i = 1
                While (CDbl(Txt_campo32.Text) >= i)         'ARREGLO 2
                  Set rs_aux2 = New ADODB.Recordset
                  If rs_aux2.State = 1 Then rs_aux2.Close
                  rs_aux2.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 2 and cotiza_codigo = " & i & " ", db, adOpenStatic
                  If rs_aux2.RecordCount > 0 Then
                     db.Execute "Update ao_solicitud_cotiza_venta Set pais_continente = '" & Ado_datos.Recordset!pais_continente2 & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 2 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_venta Set modelo_codigo = '" & Ado_datos.Recordset!modelo_codigo2 & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 2 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_venta Set fecha_registro = '" & Date & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 2 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_venta Set usr_codigo = '" & glusuario & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 2 and cotiza_codigo = " & i & "  "
                      'rs_aux2!pais_continente = Ado_datos.Recordset!pais_continente2
                      'rs_aux2!modelo_codigo = Ado_datos.Recordset!modelo_codigo2
                      'rs_aux2!arreglo = 2     'arreglo2
                      'rs_aux2!fecha_registro = Date
                      'rs_aux2!usr_codigo = glusuario
                      'rs_aux2.Update
                  End If
                  Set rs_aux4 = New ADODB.Recordset
                  If rs_aux4.State = 1 Then rs_aux4.Close
                  rs_aux4.Open "Select * from ao_solicitud_cotiza_modelo where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 2 and cotiza_codigo = " & i & " ", db, adOpenStatic
                  If rs_aux4.RecordCount > 0 Then
                     db.Execute "Update ao_solicitud_cotiza_modelo Set pais_continente = '" & Ado_datos.Recordset!pais_continente2 & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 2 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_modelo Set modelo_codigo = '" & Ado_datos.Recordset!modelo_codigo2 & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 2 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_modelo Set fecha_registro = '" & Date & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 2 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_modelo Set usr_codigo = '" & glusuario & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 2 and cotiza_codigo = " & i & "  "
                      'rs_aux4!pais_continente = Ado_datos.Recordset!pais_continente
                      'rs_aux4!modelo_codigo = Ado_datos.Recordset!modelo_codigo
                      'rs_aux4!arreglo = 2     'arreglo2
                      'rs_aux4!fecha_registro = Date
                      'rs_aux4!usr_codigo = glusuario
                      'rs_aux4.Update
                  End If
                i = i + 1
                Wend
             End If
             '6 end
             '7 if
             If Txt_campo23.Text <> "" Then                   'ARREGLO 3
                i = 1
                While (CDbl(Txt_campo33.Text) >= i)         'ARREGLO 3
                  Set rs_aux2 = New ADODB.Recordset
                  If rs_aux2.State = 1 Then rs_aux2.Close
                  rs_aux2.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 3 and cotiza_codigo = " & i & " ", db, adOpenStatic
                  If rs_aux2.RecordCount > 0 Then
                     db.Execute "Update ao_solicitud_cotiza_venta Set pais_continente = '" & Ado_datos.Recordset!pais_continente3 & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 3 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_venta Set modelo_codigo = '" & Ado_datos.Recordset!modelo_codigo3 & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 3 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_venta Set fecha_registro = '" & Date & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 3 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_venta Set usr_codigo = '" & glusuario & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 3 and cotiza_codigo = " & i & "  "
                      'rs_aux2!pais_continente = Ado_datos.Recordset!pais_continente3
                      'rs_aux2!modelo_codigo = Ado_datos.Recordset!modelo_codigo3
                      'rs_aux2!arreglo = 3     'arreglo3
                      'rs_aux2!fecha_registro = Date
                      'rs_aux2!usr_codigo = glusuario
                      'rs_aux2.Update
                  End If
                  Set rs_aux4 = New ADODB.Recordset
                  If rs_aux4.State = 1 Then rs_aux4.Close
                  rs_aux4.Open "Select * from ao_solicitud_cotiza_modelo where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 3 and cotiza_codigo = " & i & " ", db, adOpenStatic
                  If rs_aux4.RecordCount > 0 Then
                     db.Execute "Update ao_solicitud_cotiza_modelo Set pais_continente = '" & Ado_datos.Recordset!pais_continente3 & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 3 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_modelo Set modelo_codigo = '" & Ado_datos.Recordset!modelo_codigo3 & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 3 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_modelo Set fecha_registro = '" & Date & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 3 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_modelo Set usr_codigo = '" & glusuario & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 3 and cotiza_codigo = " & i & "  "
                      'rs_aux4!pais_continente = Ado_datos.Recordset!pais_continente
                      'rs_aux4!modelo_codigo = Ado_datos.Recordset!modelo_codigo
                      'rs_aux4!arreglo = 3     'arreglo3
                      'rs_aux4!fecha_registro = Date
                      'rs_aux4!usr_codigo = glusuario
                      'rs_aux4.Update
                  End If
                i = i + 1
                Wend
             End If
             '7 end
             '8 if
             If Txt_campo24.Text <> "" Then                  'ARREGLO 4
                i = 1
                While (CDbl(Txt_campo34.Text) >= i)         'ARREGLO 4
                  Set rs_aux2 = New ADODB.Recordset
                  If rs_aux2.State = 1 Then rs_aux2.Close
                  rs_aux2.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 4 and cotiza_codigo = " & i & " ", db, adOpenStatic
                  If rs_aux2.RecordCount > 0 Then
                     db.Execute "Update ao_solicitud_cotiza_venta Set pais_continente4 = '" & Ado_datos.Recordset!pais_continente3 & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 4 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_venta Set modelo_codigo4 = '" & Ado_datos.Recordset!modelo_codigo3 & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 4 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_venta Set fecha_registro = '" & Date & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 4 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_venta Set usr_codigo = '" & glusuario & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 4 and cotiza_codigo = " & i & "  "
                      'rs_aux2!pais_continente = Ado_datos.Recordset!pais_continente4
                      'rs_aux2!modelo_codigo = Ado_datos.Recordset!modelo_codigo4
                      'rs_aux2!arreglo = 4     'arreglo4
                      'rs_aux2!fecha_registro = Date
                      'rs_aux2!usr_codigo = glusuario
                      'rs_aux2.Update
                  End If
                  Set rs_aux4 = New ADODB.Recordset
                  If rs_aux4.State = 1 Then rs_aux4.Close
                  rs_aux4.Open "Select * from ao_solicitud_cotiza_modelo where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 4 and cotiza_codigo = " & i & " ", db, adOpenStatic
                  If rs_aux4.RecordCount > 0 Then
                     db.Execute "Update ao_solicitud_cotiza_modelo Set pais_continente4 = '" & Ado_datos.Recordset!pais_continente3 & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 4 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_modelo Set modelo_codigo4 = '" & Ado_datos.Recordset!modelo_codigo3 & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 4 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_modelo Set fecha_registro = '" & Date & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 4 and cotiza_codigo = " & i & "  "
                     db.Execute "Update ao_solicitud_cotiza_modelo Set usr_codigo = '" & glusuario & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & " and arreglo = 4 and cotiza_codigo = " & i & "  "
                      'rs_aux4!pais_continente = Ado_datos.Recordset!pais_continente
                      'rs_aux4!modelo_codigo = Ado_datos.Recordset!modelo_codigo
                      'rs_aux4!arreglo = 4     'arreglo4
                      'rs_aux4!fecha_registro = Date
                      'rs_aux4!usr_codigo = glusuario
                      'rs_aux4.Update
                  End If
                i = i + 1
                Wend
             End If
             '8 end

           End If
           '4 end
        Else        '3 ELSE
            'VAR_VAL,
            VAR_COD4 = rs_datos!solicitud_codigo
            If Left(Txt_campo2, 4) = "36NO" Or Left(Txt_campo2, 4) = "OA36" Then
                rs_datos!unidad_codigo_ant = Trim(Txt_campo2)
            Else
                If VAR_COD4 < 10 Then
                   rs_datos!unidad_codigo_ant = parametro + "-00000" + Trim(txt_codigo)
                End If
                If VAR_COD4 > 9 And VAR_COD4 < 100 Then
                   rs_datos!unidad_codigo_ant = parametro + "-0000" + Trim(txt_codigo)
                End If
                If VAR_COD4 > 99 And VAR_COD4 < 1000 Then
                   rs_datos!unidad_codigo_ant = parametro + "-000" + Trim(txt_codigo)
                End If
                If VAR_COD4 > 999 And VAR_COD4 < 10000 Then
                   rs_datos!unidad_codigo_ant = parametro + "-00" + Trim(txt_codigo)
                End If
                If VAR_COD4 > 9999 And VAR_COD4 < 100000 Then
                   rs_datos!unidad_codigo_ant = parametro + "-0" + Trim(txt_codigo)
                End If
                If VAR_COD4 > 99999 Then
                   rs_datos!unidad_codigo_ant = parametro + "-" + Trim(txt_codigo)
                End If
            End If
            
            'Call GRABA_ARREGLO1
            Call GRABA_ARREGLO_1
            '6 if
            If Txt_campo22.Text <> "" Then               'ARREGLO 2
                'Call GRABA_ARREGLO2
                Call GRABA_ARREGLO_2
            End If
            '6 end
            '7 if
            If Txt_campo23.Text <> "" Then                   'ARREGLO 3
                'Call GRABA_ARREGLO3
                Call GRABA_ARREGLO_3
            End If
            '7 end
            '8 if
            If Txt_campo24.Text <> "" Then                  'ARREGLO 4
                'Call GRABA_ARREGLO4
                Call GRABA_ARREGLO_4
            End If
            '8 end
            
'            VAR_NO2 = VAR_NO2 + rs_datos!h_nro_total_equipos - 1
'            VAR_NO3 = "36NO-" + Trim(Str(VAR_NO2))
'            If rs_datos!h_nro_total_equipos > 1 Then
'                'If Right(VAR_NO3, 1) = 0 Then
'                    rs_datos!unidad_codigo_ant = VAR_NO1 + "-" + Right(VAR_NO3, 2)
'                'Else
'                '    rs_datos!unidad_codigo_ant = VAR_NO1 + "/" + Right(VAR_NO3, 1)
'                'End If
'            Else
'                rs_datos!unidad_codigo_ant = VAR_NO1
'            End If
'            rs_datos!unidad_codigo_ant = rs_datos!unidad_codigo + Trim(Str(rs_datos!solicitud_codigo))
    '        db.Execute "Update ao_solicitud Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
    '        db.Execute "Update ao_solicitud_cotiza_venta Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
    '        db.Execute "Update ao_negociacion_cabecera Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and negocia_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
        
            rs_datos!estado_codigo_verif = "APR"
            rs_datos!fecha_registro = Date
            rs_datos!usr_codigo = glusuario
            rs_datos.UpdateBatch adAffectAll
       End If
       '3 end
   Else
        MsgBox "No se puede VERIFICAR un registro Anulado o Aprobado o que no tiene PARAMETROS DE CALCULO ...", vbExclamation, "Validación de Registro"
   End If
   '2 end
 End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

'Private Sub GRABA_ARREGLO1()
'    i = 1
'    While (CDbl(Txt_campo31.Text) >= i)         'ARREGLO 1
'       Select Case dtc_codigo011.Text
'          Case "AMERICA"
'              'Brasil 1   - AMERICA
'              VAR_MOD = Trim(dtc_codigo61.Text) + Trim(dtc_codigo71.Text) + Trim(dtc_codigo31.Text) + Trim(dtc_codigo41.Text) + Trim(dtc_codigo51.Text) + Trim(dtc_codigo81.Text) + Trim(dtc_codigo91.Text) + Trim(dtc_codigo01.Text)
'              'rs_datos!modelo_codigo = Trim(VAR_MOD)
'          Case "ASIA"
'              'Hypex  1   - ASIA
'              If CDbl(dtc_aux31.Text) < 1000 Then     '1
'                  VAR_AUX = Trim("0" + Trim(dtc_aux31.Text))
'              Else
'                  VAR_AUX = Trim(dtc_aux31.Text)
'              End If
'              VAR_MOD = "P" + Trim(VAR_AUX) + "G" + Trim(dtc_codigo41.Text) + Trim(dtc_codigo91.Text) + "-" + Trim(dtc_campo51.Text)
'              'rs_datos!modelo_codigo_h1 = Trim(VAR_MOD)
'          Case "EUROPA"
'              'Xizi   XO  1   -   ESPAÑA
'              If CDbl(dtc_valor41.Text) < 3 Then
'                  VAR_MOD = "OH5000" + Trim(dtc_codigo31.Text) + Trim(dtc_codigo41.Text) + Trim(dtc_codigo51.Text) + Trim(dtc_codigo81.Text)
'              Else
'                  VAR_MOD = "XO8000" + Trim(dtc_codigo31.Text) + Trim(dtc_codigo41.Text) + Trim(dtc_codigo51.Text) + Trim(dtc_codigo81.Text)
'              End If
'              'rs_datos!modelo_codigo_x1 = Trim(VAR_MOD)
'          Case Else
'              VAR_MOD = Trim(dtc_codigo61.Text) + Trim(dtc_codigo71.Text) + Trim(dtc_codigo31.Text) + Trim(dtc_codigo41.Text) + Trim(dtc_codigo51.Text) + Trim(dtc_codigo81.Text) + Trim(dtc_codigo91.Text) + Trim(dtc_codigo01.Text)
'              'rs_datos!modelo_codigo = Trim(VAR_MOD)
'       End Select
'       rs_datos!modelo_codigo = Trim(VAR_MOD)
'       rs_datos!modelo_codigo_h1 = "S/M"      'Trim(VAR_MOD)
'       rs_datos!modelo_codigo_x1 = "S/M"      'Trim(VAR_MOD)
'       'Graba en Cotiza    1
'       Set rs_aux1 = New ADODB.Recordset
'       SQL_FOR = "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & "  "
'       'SQL_FOR = "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and pais_continente = '" & Ado_datos.Recordset!pais_continente & "'  "
'       rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'       '5 if
'       'If rs_aux1.RecordCount = 0 Then
'          Set rs_aux2 = New ADODB.Recordset
'          If rs_aux2.State = 1 Then rs_aux2.Close
'          rs_aux2.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & "   ", db, adOpenStatic
'          If Not rs_aux2.EOF Then
'               var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'          End If
'          rs_aux1.AddNew
'          rs_aux1!ges_gestion = Year(Date)
'          rs_aux1!unidad_codigo = parametro   'Ado_datos.Recordset!unidad_codigo
'          rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
'          rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'          rs_aux1!trafico_codigo = Ado_datos.Recordset!trafico_codigo
'          rs_aux1!pais_continente = Ado_datos.Recordset!pais_continente
'          rs_aux1!cotiza_codigo = var_cod
'          rs_aux1!arreglo = 1     'arreglo1
'          'correlativo Equipos            'WC2015
'          'Call correl_bien
'          'VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
'          'rs_aux1!bien_codigo = VAR_COD3  '"36NO-" + Trim(Str(VAR_COD2))
'          'WWWWWWWWWWWWWW AQUI JQA 16-feb-2015 WWWWWWWWWWW
'          VAR_COD3 = "NA" + Trim(Str(i))
'          rs_aux1!bien_codigo = VAR_COD3  '"NA" + Trim(Str(VAR_COD2))
'          'WWWWWWWWWWWWWW AQUI JQA 16-feb-2015 WWWWWWWWWWW
'          rs_aux1!modelo_codigo = Ado_datos.Recordset!modelo_codigo
'          rs_aux1!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h1
'          rs_aux1!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x1
'          rs_aux1!estado_codigo = "REG"
'          rs_aux1!fecha_registro = Date
'          rs_aux1!usr_codigo = glusuario
'          rs_aux1.Update
'          db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod & " Where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'          'GRABA ao_solicitud_cotiza_modelo
'            Set rs_aux4 = New ADODB.Recordset
'            If rs_aux4.State = 1 Then rs_aux4.Close
'            rs_aux4.Open "Select * from ao_solicitud_cotiza_modelo where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & "  ", db, adOpenKeyset, adLockOptimistic
'            rs_aux4.AddNew
'            rs_aux4!ges_gestion = Year(Date)
'            rs_aux4!unidad_codigo = parametro
'            rs_aux4!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
'            rs_aux4!edif_codigo = Ado_datos.Recordset!edif_codigo
'            rs_aux4!trafico_codigo = Ado_datos.Recordset!trafico_codigo
'            rs_aux4!pais_continente = Ado_datos.Recordset!pais_continente
'            rs_aux4!cotiza_codigo = var_cod
'            rs_aux4!modelo_codigo = Ado_datos.Recordset!modelo_codigo
'            rs_aux4!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo
'            rs_aux4!arreglo = 1     'arreglo1
'            rs_aux4!bien_codigo = VAR_COD3
'            rs_aux4!estado_codigo = "REG"
'            rs_aux4!fecha_registro = Date
'            rs_aux4!usr_codigo = glusuario
'            rs_aux4.Update
'
'          'GRABA EN AC_BIENES
'          If i = 1 Then
'              VAR_NO1 = VAR_COD3
'              VAR_NO2 = i         'VAR_COD2
'          End If
'          'Call GRABA_BIENES          'WC2015
'       'End If
'       '5 end
'    i = i + 1
'    Wend
'End Sub

Private Sub GRABA_ARREGLO_1()
    'WWWWWWWWWWWWWW AQUI JQA 29-OCT-2015 WWWWWWWWWWW
'    i = 1
'    While (CDbl(Txt_campo31.Text) >= i)         'ARREGLO 1
'        Set rs_aux3 = New ADODB.Recordset
'        If rs_aux3.State = 1 Then rs_aux3.Close
'        rs_aux3.Open "Select * from ao_solicitud_bienes where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenStatic
'
'        'WWWWWWWWWWWWWW AQUI JQA 16-feb-2015 WWWWWWWWWWW
'          VAR_COD3 = "NA" + Trim(Str(i))
'          rs_aux3!bien_codigo = VAR_COD3  '"NA" + Trim(Str(VAR_COD2))
'          'WWWWWWWWWWWWWW AQUI JQA 16-feb-2015 WWWWWWWWWWW
'        'GRABA EN AC_BIENES
'          If i = 1 Then
'              VAR_NO1 = VAR_COD3
'              VAR_NO2 = i         'VAR_COD2
'          End If
'    i = i + 1
'    Wend
    'WWWWWWWWWWWWWW AQUI JQA 29-OCT-2015 WWWWWWWWWWW
       Select Case dtc_codigo011.Text
          Case "AMERICA"
              'Brasil 1   - AMERICA
              VAR_MOD = Trim(dtc_codigo61.Text) + Trim(dtc_codigo71.Text) + Trim(dtc_codigo31.Text) + Trim(dtc_codigo41.Text) + Trim(dtc_codigo51.Text) + Trim(dtc_codigo81.Text) + Trim(dtc_codigo91.Text) + Trim(dtc_codigo01.Text)
              'rs_datos!modelo_codigo = Trim(VAR_MOD)
          Case "ASIA"
              'Hypex  1   - ASIA
              If CDbl(dtc_aux31.Text) < 1000 Then     '1
                  VAR_AUX = Trim("0" + Trim(dtc_aux31.Text))
              Else
                  VAR_AUX = Trim(dtc_aux31.Text)
              End If
              VAR_MOD = "P" + Trim(VAR_AUX) + "G" + Trim(dtc_codigo41.Text) + Trim(dtc_codigo91.Text) + "-" + Trim(dtc_campo51.Text)
              'rs_datos!modelo_codigo_h1 = Trim(VAR_MOD)
          Case "EUROPA"
              'Xizi   XO  1   -   ESPAÑA
              If CDbl(dtc_valor41.Text) < 3 Then
                  VAR_MOD = "OH5000" + Trim(dtc_codigo31.Text) + Trim(dtc_codigo41.Text) + Trim(dtc_codigo51.Text) + Trim(dtc_codigo81.Text)
              Else
                  VAR_MOD = "XO8000" + Trim(dtc_codigo31.Text) + Trim(dtc_codigo41.Text) + Trim(dtc_codigo51.Text) + Trim(dtc_codigo81.Text)
              End If
              'rs_datos!modelo_codigo_x1 = Trim(VAR_MOD)
          Case Else
              VAR_MOD = Trim(dtc_codigo61.Text) + Trim(dtc_codigo71.Text) + Trim(dtc_codigo31.Text) + Trim(dtc_codigo41.Text) + Trim(dtc_codigo51.Text) + Trim(dtc_codigo81.Text) + Trim(dtc_codigo91.Text) + Trim(dtc_codigo01.Text)
              'rs_datos!modelo_codigo = Trim(VAR_MOD)
       End Select
       rs_datos!modelo_codigo = Trim(VAR_MOD)
       rs_datos!modelo_codigo_h1 = "S/M"      'Trim(VAR_MOD)
       rs_datos!modelo_codigo_x1 = "S/M"      'Trim(VAR_MOD)
       'Graba en Cotiza    1
       
       Set rs_aux1 = New ADODB.Recordset
       SQL_FOR = "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & "  "
       'SQL_FOR = "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and pais_continente = '" & Ado_datos.Recordset!pais_continente & "'  "
       rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
       '5 if
       'If rs_aux1.RecordCount = 0 Then
          Set rs_aux2 = New ADODB.Recordset
          If rs_aux2.State = 1 Then rs_aux2.Close
          rs_aux2.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & "   ", db, adOpenStatic
          'JQA NEW
          'rs_aux2.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "'    ", db, adOpenStatic
          If Not rs_aux2.EOF Then
               var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
          End If
          rs_aux1.AddNew
          rs_aux1!ges_gestion = Year(Date)
          rs_aux1!unidad_codigo = parametro   'Ado_datos.Recordset!unidad_codigo
          rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
          rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
          rs_aux1!trafico_codigo = Ado_datos.Recordset!trafico_codigo
          rs_aux1!pais_continente = Ado_datos.Recordset!pais_continente
          rs_aux1!cotiza_codigo = var_cod
          rs_aux1!arreglo = Ado_datos.Recordset!arreglo1    '1     'arreglo1
          rs_aux1!cotiza_cantidad = Ado_datos.Recordset!trafico_nro_equipos
          'correlativo Equipos            'WC2015
          'Call correl_bien
          'VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
          'rs_aux1!bien_codigo = VAR_COD3  '"36NO-" + Trim(Str(VAR_COD2))
          'WWWWWWWWWWWWWW AQUI JQA 16-feb-2015 WWWWWWWWWWW
          'VAR_COD3 = "NA" + Trim(Str(i))
          rs_aux1!bien_codigo = "NA1"  'VAR_COD3  '"NA" + Trim(Str(VAR_COD2))
          'WWWWWWWWWWWWWW AQUI JQA 16-feb-2015 WWWWWWWWWWW
          rs_aux1!modelo_codigo = Ado_datos.Recordset!modelo_codigo
          'rs_aux1!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h1
          'rs_aux1!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x1
          rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
          rs_aux1!estado_codigo = "REG"
          rs_aux1!estado_codigo_verif = "REG"
          rs_aux1!fecha_registro = Date
          rs_aux1!usr_codigo = glusuario
          rs_aux1.Update
          db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod & " Where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
          'GRABA ao_solicitud_cotiza_modelo
            Set rs_aux4 = New ADODB.Recordset
            If rs_aux4.State = 1 Then rs_aux4.Close
            rs_aux4.Open "Select * from ao_solicitud_cotiza_modelo where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & "  ", db, adOpenKeyset, adLockOptimistic
            rs_aux4.AddNew
            rs_aux4!ges_gestion = Year(Date)
            rs_aux4!unidad_codigo = parametro
            rs_aux4!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
            rs_aux4!edif_codigo = Ado_datos.Recordset!edif_codigo
            rs_aux4!trafico_codigo = Ado_datos.Recordset!trafico_codigo
            rs_aux4!pais_continente = Ado_datos.Recordset!pais_continente
            rs_aux4!cotiza_codigo = var_cod
            rs_aux4!modelo_codigo = Ado_datos.Recordset!modelo_codigo
            rs_aux4!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo
            rs_aux4!arreglo = Ado_datos.Recordset!arreglo1   '1     'arreglo1
            rs_aux4!cotiza_cantidad = Ado_datos.Recordset!trafico_nro_equipos
            rs_aux4!bien_codigo = "NA1"
            rs_aux4!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
            rs_aux4!beneficiario_codigo_resp = Ado_datos.Recordset!beneficiario_codigo_resp
            rs_aux4!estado_codigo = "REG"
            rs_aux4!estado_codigo_verif = "REG"
            rs_aux4!fecha_registro = Date
            rs_aux4!usr_codigo = glusuario
            rs_aux4.Update

          'GRABA EN AC_BIENES
          'If i = 1 Then
          '    VAR_NO1 = VAR_COD3
          '    VAR_NO2 = i         'VAR_COD2
          'End If
          'Call GRABA_BIENES          'WC2015
       'End If
       '5 end
'    i = i + 1
'    Wend
End Sub

'Private Sub GRABA_ARREGLO2()
'    i = 1
'    While (CDbl(Txt_campo32.Text) >= i)
'      Select Case dtc_codigo012.Text
'          Case "AMERICA"
'              'Brasil 2   - AMERICA
'              VAR_MOD = Trim(dtc_codigo62.Text) + Trim(dtc_codigo72.Text) + Trim(dtc_codigo32.Text) + Trim(dtc_codigo42.Text) + Trim(dtc_codigo52.Text) + Trim(dtc_codigo82.Text) + Trim(dtc_codigo92.Text) + Trim(dtc_codigo02.Text)
'              'rs_datos!modelo_codigo = Trim(VAR_MOD)
'          Case "ASIA"
'              'Hypex  2   - ASIA
'              If CDbl(dtc_aux32.Text) < 1000 Then
'                  VAR_AUX = Trim("0" + Trim(dtc_aux32.Text))
'              Else
'                  VAR_AUX = Trim(dtc_aux32.Text)
'              End If
'              VAR_MOD = "P" + Trim(VAR_AUX) + "G" + Trim(dtc_codigo42.Text) + Trim(dtc_codigo92.Text) + "-" + Trim(dtc_campo52.Text)
'              'rs_datos!modelo_codigo_h1 = Trim(VAR_MOD)
'          Case "EUROPA"
'              'Xizi   XO  2   -   ESPAÑA
'              If CDbl(dtc_desc42.Text) < 3 Then
'                  VAR_MOD = "OH5000" + Trim(dtc_codigo32.Text) + Trim(dtc_codigo42.Text) + Trim(dtc_codigo52.Text) + Trim(dtc_codigo82.Text)
'              Else
'                  VAR_MOD = "XO8000" + Trim(dtc_codigo32.Text) + Trim(dtc_codigo42.Text) + Trim(dtc_codigo52.Text) + Trim(dtc_codigo82.Text)
'              End If
'              'rs_datos!modelo_codigo_x1 = Trim(VAR_MOD)
'          Case Else
'              VAR_MOD = Trim(dtc_codigo62.Text) + Trim(dtc_codigo72.Text) + Trim(dtc_codigo32.Text) + Trim(dtc_codigo42.Text) + Trim(dtc_codigo52.Text) + Trim(dtc_codigo82.Text) + Trim(dtc_codigo92.Text) + Trim(dtc_codigo02.Text)
'              'rs_datos!modelo_codigo = Trim(VAR_MOD)
'      End Select
'      rs_datos!modelo_codigo2 = Trim(VAR_MOD)
'      rs_datos!modelo_codigo_h2 = "S/M"      'Trim(VAR_MOD)
'      rs_datos!modelo_codigo_x2 = "S/M"      'Trim(VAR_MOD)
'
'      'Graba en Cotiza    2
'      Set rs_aux2 = New ADODB.Recordset
'      If rs_aux2.State = 1 Then rs_aux2.Close
'      rs_aux2.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and pais_continente = '" & Ado_datos.Recordset!pais_continente & "'   ", db, adOpenStatic
'      If Not rs_aux2.EOF Then
'           var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'      End If
'      rs_aux1.AddNew
'      rs_aux1!ges_gestion = Year(Date)
'      rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
'      rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
'      rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'      rs_aux1!trafico_codigo = Ado_datos.Recordset!trafico_codigo
'      rs_aux1!pais_continente = Ado_datos.Recordset!pais_continente2
'      rs_aux1!cotiza_codigo = var_cod
'      rs_aux1!arreglo = 2     'arreglo2
'      'Call correl_bien
'      'VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
'      VAR_CONT1 = rs_datos!trafico_nro_equipos + i
'      VAR_COD3 = "NA" + Trim(Str(VAR_CONT1))
'      rs_aux1!bien_codigo = VAR_COD3  '"NA" + Trim(Str(VAR_COD2))
'
'      rs_aux1!modelo_codigo = Ado_datos.Recordset!modelo_codigo2
'      rs_aux1!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h2
'      rs_aux1!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x2
'      rs_aux1!estado_codigo = "REG"
'      rs_aux1!fecha_registro = Date
'      rs_aux1!usr_codigo = glusuario
'      rs_aux1.Update
'      db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'      'GRABA EN AC_BIENES
'      'Call GRABA_BIENES
'    i = i + 1
'    Wend
'End Sub

Private Sub GRABA_ARREGLO_2()
'    i = 1
'    While (CDbl(Txt_campo32.Text) >= i)
      
      Select Case dtc_codigo012.Text
          Case "AMERICA"
              'Brasil 2   - AMERICA
              VAR_MOD = Trim(dtc_codigo62.Text) + Trim(dtc_codigo72.Text) + Trim(dtc_codigo32.Text) + Trim(dtc_codigo42.Text) + Trim(dtc_codigo52.Text) + Trim(dtc_codigo82.Text) + Trim(dtc_codigo92.Text) + Trim(dtc_codigo02.Text)
              'rs_datos!modelo_codigo = Trim(VAR_MOD)
          Case "ASIA"
              'Hypex  2   - ASIA
              If CDbl(dtc_aux32.Text) < 1000 Then
                  VAR_AUX = Trim("0" + Trim(dtc_aux32.Text))
              Else
                  VAR_AUX = Trim(dtc_aux32.Text)
              End If
              VAR_MOD = "P" + Trim(VAR_AUX) + "G" + Trim(dtc_codigo42.Text) + Trim(dtc_codigo92.Text) + "-" + Trim(dtc_campo52.Text)
              'rs_datos!modelo_codigo_h1 = Trim(VAR_MOD)
          Case "EUROPA"
              'Xizi   XO  2   -   ESPAÑA
              If CDbl(dtc_desc42.Text) < 3 Then
                  VAR_MOD = "OH5000" + Trim(dtc_codigo32.Text) + Trim(dtc_codigo42.Text) + Trim(dtc_codigo52.Text) + Trim(dtc_codigo82.Text)
              Else
                  VAR_MOD = "XO8000" + Trim(dtc_codigo32.Text) + Trim(dtc_codigo42.Text) + Trim(dtc_codigo52.Text) + Trim(dtc_codigo82.Text)
              End If
              'rs_datos!modelo_codigo_x1 = Trim(VAR_MOD)
          Case Else
              VAR_MOD = Trim(dtc_codigo62.Text) + Trim(dtc_codigo72.Text) + Trim(dtc_codigo32.Text) + Trim(dtc_codigo42.Text) + Trim(dtc_codigo52.Text) + Trim(dtc_codigo82.Text) + Trim(dtc_codigo92.Text) + Trim(dtc_codigo02.Text)
              'rs_datos!modelo_codigo = Trim(VAR_MOD)
      End Select
      rs_datos!modelo_codigo2 = Trim(VAR_MOD)
      rs_datos!modelo_codigo_h2 = "S/M"      'Trim(VAR_MOD)
      rs_datos!modelo_codigo_x2 = "S/M"      'Trim(VAR_MOD)
      'Graba en Cotiza    2
      Set rs_aux1 = New ADODB.Recordset
      If rs_aux1.State = 1 Then rs_aux1.Close
      SQL_FOR = "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & "  "
       'SQL_FOR = "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and pais_continente = '" & Ado_datos.Recordset!pais_continente & "'  "
      rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
      'Graba en Cotiza    2
      Set rs_aux2 = New ADODB.Recordset
      If rs_aux2.State = 1 Then rs_aux2.Close
      rs_aux2.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and pais_continente = '" & Ado_datos.Recordset!pais_continente & "'   ", db, adOpenStatic
      If Not rs_aux2.EOF Then
           var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
      End If
      rs_aux1.AddNew
      rs_aux1!ges_gestion = Year(Date)
      rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
      rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
      rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
      rs_aux1!trafico_codigo = Ado_datos.Recordset!trafico_codigo
      rs_aux1!pais_continente = Ado_datos.Recordset!pais_continente2
      rs_aux1!cotiza_codigo = var_cod
      rs_aux1!arreglo = Ado_datos.Recordset!arreglo2    '2
      rs_aux1!cotiza_cantidad = Ado_datos.Recordset!trafico_nro_equipos2
      'Call correl_bien
 '     'VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
 '     VAR_CONT1 = rs_datos!trafico_nro_equipos + i
 '     VAR_COD3 = "NA" + Trim(Str(VAR_CONT1))
 '     rs_aux1!bien_codigo = VAR_COD3  '"NA" + Trim(Str(VAR_COD2))
      rs_aux1!bien_codigo = "NA2"

      rs_aux1!modelo_codigo = Ado_datos.Recordset!modelo_codigo2
      'rs_aux1!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h2
      'rs_aux1!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x2
      rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
      rs_aux1!estado_codigo = "REG"
      rs_aux1!estado_codigo_verif = "REG"
      rs_aux1!fecha_registro = Date
      rs_aux1!usr_codigo = glusuario
      rs_aux1.Update
      db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
          'GRABA ao_solicitud_cotiza_modelo
            Set rs_aux4 = New ADODB.Recordset
            If rs_aux4.State = 1 Then rs_aux4.Close
            rs_aux4.Open "Select * from ao_solicitud_cotiza_modelo where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & "  ", db, adOpenKeyset, adLockOptimistic
            rs_aux4.AddNew
            rs_aux4!ges_gestion = Year(Date)
            rs_aux4!unidad_codigo = parametro
            rs_aux4!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
            rs_aux4!edif_codigo = Ado_datos.Recordset!edif_codigo
            rs_aux4!trafico_codigo = Ado_datos.Recordset!trafico_codigo
            rs_aux4!pais_continente = Ado_datos.Recordset!pais_continente2
            rs_aux4!cotiza_codigo = var_cod
            rs_aux4!modelo_codigo = Ado_datos.Recordset!modelo_codigo2
            rs_aux4!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo2
            rs_aux4!arreglo = 2     'arreglo2
            rs_aux4!cotiza_cantidad = Ado_datos.Recordset!trafico_nro_equipos2
            rs_aux4!bien_codigo = "NA2"
            rs_aux4!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
            rs_aux4!beneficiario_codigo_resp = Ado_datos.Recordset!beneficiario_codigo_resp
            rs_aux4!estado_codigo = "REG"
            rs_aux4!estado_codigo_verif = "REG"
            rs_aux4!fecha_registro = Date
            rs_aux4!usr_codigo = glusuario
            rs_aux4.Update
      
      'GRABA EN AC_BIENES
      'Call GRABA_BIENES
'    i = i + 1
'    Wend
End Sub

'Private Sub GRABA_ARREGLO3()
'    i = 1
'    While (CDbl(Txt_campo33.Text) >= i)
'      Select Case dtc_codigo013.Text
'          Case "AMERICA"
'              'Brasil 3   - AMERICA
'              VAR_MOD = Trim(dtc_codigo63.Text) + Trim(dtc_codigo73.Text) + Trim(dtc_codigo33.Text) + Trim(dtc_codigo43.Text) + Trim(dtc_codigo53.Text) + Trim(dtc_codigo83.Text) + Trim(dtc_codigo93.Text) + Trim(dtc_codigo03.Text)
'              'rs_datos!modelo_codigo = Trim(VAR_MOD)
'          Case "ASIA"
'              'Hypex  3   - ASIA
'              If CDbl(dtc_aux33.Text) < 1000 Then
'                  VAR_AUX = Trim("0" + Trim(dtc_aux33.Text))
'              Else
'                  VAR_AUX = Trim(dtc_aux33.Text)
'              End If
'              VAR_MOD = "P" + Trim(VAR_AUX) + "G" + Trim(dtc_codigo43.Text) + Trim(dtc_codigo93.Text) + "-" + Trim(dtc_campo53.Text)
'              'rs_datos!modelo_codigo_h1 = Trim(VAR_MOD)
'          Case "EUROPA"
'              'Xizi   XO  3   -   ESPAÑA
'              If CDbl(dtc_desc43.Text) < 3 Then
'                  VAR_MOD = "OH5000" + Trim(dtc_codigo33.Text) + Trim(dtc_codigo43.Text) + Trim(dtc_codigo53.Text) + Trim(dtc_codigo83.Text)
'              Else
'                  VAR_MOD = "XO8000" + Trim(dtc_codigo33.Text) + Trim(dtc_codigo43.Text) + Trim(dtc_codigo53.Text) + Trim(dtc_codigo83.Text)
'              End If
'              'rs_datos!modelo_codigo_x1 = Trim(VAR_MOD)
'          Case Else
'              VAR_MOD = Trim(dtc_codigo63.Text) + Trim(dtc_codigo73.Text) + Trim(dtc_codigo33.Text) + Trim(dtc_codigo43.Text) + Trim(dtc_codigo53.Text) + Trim(dtc_codigo83.Text) + Trim(dtc_codigo93.Text) + Trim(dtc_codigo03.Text)
'              'rs_datos!modelo_codigo = Trim(VAR_MOD)
'      End Select
'      rs_datos!modelo_codigo3 = Trim(VAR_MOD)
'      rs_datos!modelo_codigo_h3 = "S/M"      'Trim(VAR_MOD)
'      rs_datos!modelo_codigo_x3 = "S/M"      'Trim(VAR_MOD)
'
'      'Graba en Cotiza    3
'      Set rs_aux2 = New ADODB.Recordset
'      If rs_aux2.State = 1 Then rs_aux2.Close
'      rs_aux2.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'   ", db, adOpenStatic
'      If Not rs_aux2.EOF Then
'           var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'      End If
'      rs_aux1.AddNew
'      rs_aux1!ges_gestion = Year(Date)
'      rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
'      rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
'      rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'      rs_aux1!trafico_codigo = Ado_datos.Recordset!trafico_codigo
'      rs_aux1!pais_continente = Ado_datos.Recordset!pais_continente3
'      rs_aux1!cotiza_codigo = var_cod
'      rs_aux1!arreglo = 3     'arreglo3
'      'Call correl_bien
'      'VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
'      'rs_aux1!bien_codigo = VAR_COD3
'      VAR_CONT1 = rs_datos!trafico_nro_equipos + rs_datos!trafico_nro_equipos2 + i
'      VAR_COD3 = "NA" + Trim(Str(VAR_CONT1))
'      rs_aux1!bien_codigo = VAR_COD3  '"NA" + Trim(Str(VAR_COD2))
'
'      rs_aux1!modelo_codigo = Ado_datos.Recordset!modelo_codigo3
'      rs_aux1!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h3
'      rs_aux1!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x3
'      rs_aux1!estado_codigo = "REG"
'      rs_aux1!fecha_registro = Date
'      rs_aux1!usr_codigo = glusuario
'      rs_aux1.Update
'      db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'      'GRABA EN AC_BIENES
'      'Call GRABA_BIENES
'    i = i + 1
'    Wend
'End Sub

Private Sub GRABA_ARREGLO_3()
'    i = 1
'    While (CDbl(Txt_campo33.Text) >= i)
      Select Case dtc_codigo013.Text
          Case "AMERICA"
              'Brasil 3   - AMERICA
              VAR_MOD = Trim(dtc_codigo63.Text) + Trim(dtc_codigo73.Text) + Trim(dtc_codigo33.Text) + Trim(dtc_codigo43.Text) + Trim(dtc_codigo53.Text) + Trim(dtc_codigo83.Text) + Trim(dtc_codigo93.Text) + Trim(dtc_codigo03.Text)
              'rs_datos!modelo_codigo = Trim(VAR_MOD)
          Case "ASIA"
              'Hypex  3   - ASIA
              If CDbl(dtc_aux33.Text) < 1000 Then
                  VAR_AUX = Trim("0" + Trim(dtc_aux33.Text))
              Else
                  VAR_AUX = Trim(dtc_aux33.Text)
              End If
              VAR_MOD = "P" + Trim(VAR_AUX) + "G" + Trim(dtc_codigo43.Text) + Trim(dtc_codigo93.Text) + "-" + Trim(dtc_campo53.Text)
              'rs_datos!modelo_codigo_h1 = Trim(VAR_MOD)
          Case "EUROPA"
              'Xizi   XO  3   -   ESPAÑA
              If CDbl(dtc_desc43.Text) < 3 Then
                  VAR_MOD = "OH5000" + Trim(dtc_codigo33.Text) + Trim(dtc_codigo43.Text) + Trim(dtc_codigo53.Text) + Trim(dtc_codigo83.Text)
              Else
                  VAR_MOD = "XO8000" + Trim(dtc_codigo33.Text) + Trim(dtc_codigo43.Text) + Trim(dtc_codigo53.Text) + Trim(dtc_codigo83.Text)
              End If
              'rs_datos!modelo_codigo_x1 = Trim(VAR_MOD)
          Case Else
              VAR_MOD = Trim(dtc_codigo63.Text) + Trim(dtc_codigo73.Text) + Trim(dtc_codigo33.Text) + Trim(dtc_codigo43.Text) + Trim(dtc_codigo53.Text) + Trim(dtc_codigo83.Text) + Trim(dtc_codigo93.Text) + Trim(dtc_codigo03.Text)
              'rs_datos!modelo_codigo = Trim(VAR_MOD)
      End Select
      rs_datos!modelo_codigo3 = Trim(VAR_MOD)
      rs_datos!modelo_codigo_h3 = "S/M"      'Trim(VAR_MOD)
      rs_datos!modelo_codigo_x3 = "S/M"      'Trim(VAR_MOD)

      'Graba en Cotiza    3
      Set rs_aux2 = New ADODB.Recordset
      If rs_aux2.State = 1 Then rs_aux2.Close
      rs_aux2.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'   ", db, adOpenStatic
      If Not rs_aux2.EOF Then
           var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
      End If
      rs_aux1.AddNew
      rs_aux1!ges_gestion = Year(Date)
      rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
      rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
      rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
      rs_aux1!trafico_codigo = Ado_datos.Recordset!trafico_codigo
      rs_aux1!pais_continente = Ado_datos.Recordset!pais_continente3
      rs_aux1!cotiza_codigo = var_cod
      rs_aux1!arreglo = Ado_datos.Recordset!arreglo3    '3     '
      rs_aux1!cotiza_cantidad = Ado_datos.Recordset!trafico_nro_equipos3
      'Call correl_bien
      'VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
      'rs_aux1!bien_codigo = VAR_COD3
'      VAR_CONT1 = rs_datos!trafico_nro_equipos + rs_datos!trafico_nro_equipos2 + i
'      VAR_COD3 = "NA" + Trim(Str(VAR_CONT1))
'      rs_aux1!bien_codigo = VAR_COD3  '"NA" + Trim(Str(VAR_COD2))
      rs_aux1!bien_codigo = "NA3"
     
      rs_aux1!modelo_codigo = Ado_datos.Recordset!modelo_codigo3
      rs_aux1!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h3
      rs_aux1!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x3
      rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
      rs_aux1!estado_codigo = "REG"
      rs_aux1!estado_codigo_verif = "REG"
      rs_aux1!fecha_registro = Date
      rs_aux1!usr_codigo = glusuario
      rs_aux1.Update
      db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
        'GRABA ao_solicitud_cotiza_modelo
            Set rs_aux4 = New ADODB.Recordset
            If rs_aux4.State = 1 Then rs_aux4.Close
            rs_aux4.Open "Select * from ao_solicitud_cotiza_modelo where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & "  ", db, adOpenKeyset, adLockOptimistic
            rs_aux4.AddNew
            rs_aux4!ges_gestion = Year(Date)
            rs_aux4!unidad_codigo = parametro
            rs_aux4!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
            rs_aux4!edif_codigo = Ado_datos.Recordset!edif_codigo
            rs_aux4!trafico_codigo = Ado_datos.Recordset!trafico_codigo
            rs_aux4!pais_continente = Ado_datos.Recordset!pais_continente3
            rs_aux4!cotiza_codigo = var_cod
            rs_aux4!modelo_codigo = Ado_datos.Recordset!modelo_codigo3
            rs_aux4!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo3
            rs_aux4!arreglo = 3     'arreglo3
            rs_aux4!cotiza_cantidad = Ado_datos.Recordset!trafico_nro_equipos3
            rs_aux4!bien_codigo = "NA3"
            rs_aux4!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
            rs_aux4!beneficiario_codigo_resp = Ado_datos.Recordset!beneficiario_codigo_resp
            rs_aux4!estado_codigo = "REG"
            rs_aux4!estado_codigo_verif = "REG"
            rs_aux4!fecha_registro = Date
            rs_aux4!usr_codigo = glusuario
            rs_aux4.Update

      'GRABA EN AC_BIENES
      'Call GRABA_BIENES
'    i = i + 1
'    Wend
End Sub

'Private Sub GRABA_ARREGLO4()
'    i = 1
'    While (CDbl(Txt_campo34.Text) >= i)
'      Select Case dtc_codigo014.Text
'          Case "AMERICA"
'              'Brasil 4   - AMERICA
'              VAR_MOD = Trim(dtc_codigo64.Text) + Trim(dtc_codigo74.Text) + Trim(dtc_codigo34.Text) + Trim(dtc_codigo44.Text) + Trim(dtc_codigo54.Text) + Trim(dtc_codigo84.Text) + Trim(dtc_codigo94.Text) + Trim(dtc_codigo04.Text)
'              'rs_datos!modelo_codigo = Trim(VAR_MOD)
'          Case "ASIA"
'              'Hypex  4   - ASIA
'              If CDbl(dtc_aux34.Text) < 1000 Then
'                  VAR_AUX = Trim("0" + Trim(dtc_aux34.Text))
'              Else
'                  VAR_AUX = Trim(dtc_aux34.Text)
'              End If
'              VAR_MOD = "P" + Trim(VAR_AUX) + "G" + Trim(dtc_codigo44.Text) + Trim(dtc_codigo94.Text) + "-" + Trim(dtc_campo54.Text)
'              'rs_datos!modelo_codigo_h1 = Trim(VAR_MOD)
'          Case "EUROPA"
'              'Xizi   XO  4   -   ESPAÑA
'              If CDbl(dtc_desc44.Text) < 3 Then
'                  VAR_MOD = "OH5000" + Trim(dtc_codigo34.Text) + Trim(dtc_codigo44.Text) + Trim(dtc_codigo54.Text) + Trim(dtc_codigo84.Text)
'              Else
'                  VAR_MOD = "XO8000" + Trim(dtc_codigo34.Text) + Trim(dtc_codigo44.Text) + Trim(dtc_codigo54.Text) + Trim(dtc_codigo84.Text)
'              End If
'              'rs_datos!modelo_codigo_x1 = Trim(VAR_MOD)
'          Case Else
'              VAR_MOD = Trim(dtc_codigo64.Text) + Trim(dtc_codigo74.Text) + Trim(dtc_codigo34.Text) + Trim(dtc_codigo44.Text) + Trim(dtc_codigo54.Text) + Trim(dtc_codigo84.Text) + Trim(dtc_codigo94.Text) + Trim(dtc_codigo04.Text)
'              'rs_datos!modelo_codigo = Trim(VAR_MOD)
'      End Select
'      rs_datos!modelo_codigo4 = Trim(VAR_MOD)
'      rs_datos!modelo_codigo_h4 = "S/M"      'Trim(VAR_MOD)
'      rs_datos!modelo_codigo_x4 = "S/M"      'Trim(VAR_MOD)
'
'      'Graba en Cotiza    4
'      Set rs_aux2 = New ADODB.Recordset
'      If rs_aux2.State = 1 Then rs_aux2.Close
'      rs_aux2.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'   ", db, adOpenStatic
'      If Not rs_aux2.EOF Then
'           var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'      End If
'      rs_aux1.AddNew
'      rs_aux1!ges_gestion = Year(Date)
'      rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
'      rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
'      rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'      rs_aux1!trafico_codigo = Ado_datos.Recordset!trafico_codigo
'      rs_aux1!pais_continente = Ado_datos.Recordset!pais_continente4
'      rs_aux1!cotiza_codigo = var_cod
'      rs_aux1!arreglo = 4     'arreglo4
'      'Call correl_bien
'      'VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
'      'rs_aux1!bien_codigo = VAR_COD3
'      VAR_CONT1 = rs_datos!trafico_nro_equipos + rs_datos!trafico_nro_equipos2 + rs_datos!trafico_nro_equipos3 + i
'      VAR_COD3 = "NA" + Trim(Str(VAR_CONT1))
'      rs_aux1!bien_codigo = VAR_COD3  '"NA" + Trim(Str(VAR_COD2))
'
'      rs_aux1!modelo_codigo = Ado_datos.Recordset!modelo_codigo4
'      rs_aux1!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h4
'      rs_aux1!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x4
'      rs_aux1!estado_codigo = "REG"
'      rs_aux1!fecha_registro = Date
'      rs_aux1!usr_codigo = glusuario
'      rs_aux1.Update
'      db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'      'GRABA EN AC_BIENES
'      'Call GRABA_BIENES
'    i = i + 1
'    Wend
'End Sub

Private Sub GRABA_ARREGLO_4()
'    i = 1
'    While (CDbl(Txt_campo34.Text) >= i)
      Select Case dtc_codigo014.Text
          Case "AMERICA"
              'Brasil 4   - AMERICA
              VAR_MOD = Trim(dtc_codigo64.Text) + Trim(dtc_codigo74.Text) + Trim(dtc_codigo34.Text) + Trim(dtc_codigo44.Text) + Trim(dtc_codigo54.Text) + Trim(dtc_codigo84.Text) + Trim(dtc_codigo94.Text) + Trim(dtc_codigo04.Text)
              'rs_datos!modelo_codigo = Trim(VAR_MOD)
          Case "ASIA"
              'Hypex  4   - ASIA
              If CDbl(dtc_aux34.Text) < 1000 Then
                  VAR_AUX = Trim("0" + Trim(dtc_aux34.Text))
              Else
                  VAR_AUX = Trim(dtc_aux34.Text)
              End If
              VAR_MOD = "P" + Trim(VAR_AUX) + "G" + Trim(dtc_codigo44.Text) + Trim(dtc_codigo94.Text) + "-" + Trim(dtc_campo54.Text)
              'rs_datos!modelo_codigo_h1 = Trim(VAR_MOD)
          Case "EUROPA"
              'Xizi   XO  4   -   ESPAÑA
              If CDbl(dtc_desc44.Text) < 3 Then
                  VAR_MOD = "OH5000" + Trim(dtc_codigo34.Text) + Trim(dtc_codigo44.Text) + Trim(dtc_codigo54.Text) + Trim(dtc_codigo84.Text)
              Else
                  VAR_MOD = "XO8000" + Trim(dtc_codigo34.Text) + Trim(dtc_codigo44.Text) + Trim(dtc_codigo54.Text) + Trim(dtc_codigo84.Text)
              End If
              'rs_datos!modelo_codigo_x1 = Trim(VAR_MOD)
          Case Else
              VAR_MOD = Trim(dtc_codigo64.Text) + Trim(dtc_codigo74.Text) + Trim(dtc_codigo34.Text) + Trim(dtc_codigo44.Text) + Trim(dtc_codigo54.Text) + Trim(dtc_codigo84.Text) + Trim(dtc_codigo94.Text) + Trim(dtc_codigo04.Text)
              'rs_datos!modelo_codigo = Trim(VAR_MOD)
      End Select
      rs_datos!modelo_codigo4 = Trim(VAR_MOD)
      rs_datos!modelo_codigo_h4 = "S/M"      'Trim(VAR_MOD)
      rs_datos!modelo_codigo_x4 = "S/M"      'Trim(VAR_MOD)
      
      'Graba en Cotiza    4
      Set rs_aux2 = New ADODB.Recordset
      If rs_aux2.State = 1 Then rs_aux2.Close
      rs_aux2.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'   ", db, adOpenStatic
      If Not rs_aux2.EOF Then
           var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
      End If
      rs_aux1.AddNew
      rs_aux1!ges_gestion = Year(Date)
      rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
      rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
      rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
      rs_aux1!trafico_codigo = Ado_datos.Recordset!trafico_codigo
      rs_aux1!pais_continente = Ado_datos.Recordset!pais_continente4
      rs_aux1!cotiza_codigo = var_cod
      rs_aux1!arreglo = 4     'arreglo4
      rs_aux1!cotiza_cantidad = Ado_datos.Recordset!trafico_nro_equipos4
      'Call correl_bien
      'VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
      'rs_aux1!bien_codigo = VAR_COD3
      'VAR_CONT1 = rs_datos!trafico_nro_equipos + rs_datos!trafico_nro_equipos2 + rs_datos!trafico_nro_equipos3 + i
      'VAR_COD3 = "NA" + Trim(Str(VAR_CONT1))
      'rs_aux1!bien_codigo = VAR_COD3  '"NA" + Trim(Str(VAR_COD2))
      rs_aux1!bien_codigo = "NA4"
      
      rs_aux1!modelo_codigo = Ado_datos.Recordset!modelo_codigo4
      rs_aux1!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h4
      rs_aux1!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x4
      rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
      rs_aux1!estado_codigo = "REG"
      rs_aux1!estado_codigo_verif = "REG"
      rs_aux1!fecha_registro = Date
      rs_aux1!usr_codigo = glusuario
      rs_aux1.Update
      db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
        'GRABA ao_solicitud_cotiza_modelo
            Set rs_aux4 = New ADODB.Recordset
            If rs_aux4.State = 1 Then rs_aux4.Close
            rs_aux4.Open "Select * from ao_solicitud_cotiza_modelo where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and trafico_codigo = " & Ado_datos.Recordset!trafico_codigo & "  ", db, adOpenKeyset, adLockOptimistic
            rs_aux4.AddNew
            rs_aux4!ges_gestion = Year(Date)
            rs_aux4!unidad_codigo = parametro
            rs_aux4!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
            rs_aux4!edif_codigo = Ado_datos.Recordset!edif_codigo
            rs_aux4!trafico_codigo = Ado_datos.Recordset!trafico_codigo
            rs_aux4!pais_continente = Ado_datos.Recordset!pais_continente4
            rs_aux4!cotiza_codigo = var_cod
            rs_aux4!modelo_codigo = Ado_datos.Recordset!modelo_codigo4
            rs_aux4!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo4
            rs_aux4!arreglo = 4     'arreglo4
            rs_aux4!cotiza_cantidad = Ado_datos.Recordset!trafico_nro_equipos4
            rs_aux4!bien_codigo = "NA4"
            rs_aux4!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
            rs_aux4!beneficiario_codigo_resp = Ado_datos.Recordset!beneficiario_codigo_resp
            rs_aux4!estado_codigo = "REG"
            rs_aux4!estado_codigo_verif = "REG"
            rs_aux4!fecha_registro = Date
            rs_aux4!usr_codigo = glusuario
            rs_aux4.Update

      'GRABA EN AC_BIENES
      'Call GRABA_BIENES
'    i = i + 1
'    Wend
End Sub

Private Sub GRABA_BIENES()
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            rs_aux2.Open "Select * from ac_bienes where bien_codigo = '" & VAR_COD3 & "' ", db, adOpenKeyset
            If rs_aux2.RecordCount > 0 Then
                'db.Execute "Update ac_bienes Set grupo_codigo = '40000', subgrupo_codigo= '43000', par_codigo= '43340', bien_codigo= '" & VAR_COD3 & "', bien_descripcion = 'EQUIPO DE CAPACIDAD ' + '" & dtc_desc31.Text & "' + ' PERSONAS Y VELOCIDAD ' + '" & dtc_valor41.Text & "' + ' m/s' Where bien_codigo = '" & VAR_COD3 & "'  "
                'WWWWWWWWWWW JQA FEB-2015 WWWWWWWWWWWW
                'db.Execute "Update ac_bienes Set bien_codigo= '" & VAR_COD3 & "', bien_descripcion = 'EQUIPO DE CAPACIDAD ' + '" & dtc_desc31.Text & "' + ' PERSONAS Y VELOCIDAD ' + '" & dtc_valor41.Text & "' + ' m/s' Where bien_codigo = '" & VAR_COD3 & "'  "
                'WWWWWWWWWWW JQA FEB-2015 WWWWWWWWWWWW
            Else
               ' db.Execute "insert into gc_usuarios(usr_codigo, beneficiario_codigo, usr_nombres, usr_primer_apellido, usr_segundo_apellido, usr_clave, IdNivelAcceso, estado_codigo, fecha_registro, dgral_codigo, da_codigo, unidad_codigo, ocup_codigo, usr_observaciones)" & _
'            "values ('" & Left(Ado_datos.Recordset("beneficiario_nombres"), 1) & "' + '" & Ado_datos.Recordset("beneficiario_primer_apellido") & "', '" & Ado_datos.Recordset("beneficiario_codigo") & "','" & Trim(Ado_datos.Recordset("beneficiario_nombres")) & "', '" & Ado_datos.Recordset("beneficiario_primer_apellido") & "','" & Trim(Ado_datos.Recordset("beneficiario_segundo_apellido")) & "','" & Ado_datos.Recordset("beneficiario_codigo") & "', '1', 'REG', '" & Date & "', '0', '0', '0', '0', '0') "

                db.Execute "insert into ac_bienes(grupo_codigo, subgrupo_codigo, bien_codigo, par_codigo, bien_descripcion, bien_precio_compra, bien_precio_venta_base, bien_precio_venta_final, unimed_codigo, unimed_codigo_empaque, bien_cantidad_por_empaque, marca_codigo, bien_stock_minimo, bien_stock_inicial, bien_stock_ingreso, bien_stock_salida, bien_stock_actual, bien_total_compra_bs, bien_total_venta_bs, bien_utilidad_Bs, bien_codigo_anterior, bien_codigo_universal, bien_descripcion_anterior, pais_codigo, archivo_foto2, archivo_foto, estado_codigo, fecha_registro, usr_codigo) " & _
                "VALUES ('40000', '43000', '" & VAR_COD3 & "', '43340', 'EQUIPO DE CAPACIDAD ' + '" & dtc_desc31.Text & "' + ' PERSONAS Y VELOCIDAD ' + '" & dtc_valor41.Text & "' + ' m/s', " & var_cod & ", '0', '0', 'EQP', 'EQP', '1', 'S/M', '1', '0', '0', '0', '0', '0', '0', '0', '-', '-', '-', 'NN', '" & VAR_COD3 & "' + '2.JPG', '" & VAR_COD3 & "' + '.JPG', 'REG', '" & Date & "', '" & glusuario & "') "
            End If
'                rs_aux2.AddNew
'                rs_aux2!grupo_codigo = "40000"
'                rs_aux2!subgrupo_codigo = "43000"
'                rs_aux2!bien_codigo = VAR_COD3
'                rs_aux2!par_codigo = "43340"
'                rs_aux2!bien_descripcion = "EQUIPO DE CAPACIDAD " + dtc_desc31.Text + " PERSONAS Y VELOCIDAD " + dtc_valor41.Text + " m/s"
'                rs_aux2!bien_precio_compra = var_cod
'                rs_aux2!bien_precio_venta_base = 0
'                rs_aux2!bien_precio_venta_final = 0
'                rs_aux2!unimed_codigo = "EQP"
'                rs_aux2!unimed_codigo_empaque = "EQP"
'                rs_aux2!bien_cantidad_por_empaque = 1
'                rs_aux2!marca_codigo = "S/M"
'                rs_aux2!bien_stock_minimo = 1
'                rs_aux2!bien_stock_inicial = 0
'                rs_aux2!bien_stock_ingreso = 0
'                rs_aux2!bien_stock_salida = 0
'                rs_aux2!bien_stock_actual = 0
'                rs_aux2!bien_total_compra_bs = 0
'                rs_aux2!bien_total_venta_bs = 0
'                rs_aux2!bien_utilidad_Bs = 0
'                rs_aux2!bien_codigo_anterior = ""
'                rs_aux2!bien_codigo_universal = ""
'                rs_aux2!bien_descripcion_anterior = ""
'                rs_aux2!pais_codigo = "NN"
'                rs_aux2!archivo_foto2 = VAR_COD3 + "2.JPG"
'                rs_aux2!archivo_foto = VAR_COD3 + ".JPG"
'
'                rs_aux2!estado_codigo = "REG"
'                rs_aux2!fecha_registro = Date
'                rs_aux2!usr_codigo = GlUsuario
'                rs_aux2.Update
                
         
End Sub

Private Sub BtnBuscar_Click()
'   On Error GoTo UpdateErr
'   If rs_datos!estado_codigo = "REG" And rs_datos!estado_codigo_verif = "REG" And Val(Ado_datos.Recordset!trafico_num_paradas) > 0 Then
'      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'        i = 1
'        While (CDbl(Txt_campo31.Text) >= i)
'            'Brasil 1
'            VAR_MOD = Trim(dtc_codigo61.Text) + Trim(dtc_codigo71.Text) + Trim(dtc_codigo31.Text) + Trim(dtc_codigo41.Text) + Trim(dtc_codigo51.Text) + Trim(dtc_codigo81.Text) + Trim(dtc_codigo91.Text) + Trim(dtc_codigo01.Text)
'            rs_datos!modelo_codigo = Trim(VAR_MOD)
'            'Hypex  1
'            If CDbl(dtc_aux31.Text) < 1000 Then     '1
'                VAR_AUX = Trim("0" + Trim(dtc_aux31.Text))
'            Else
'                VAR_AUX = Trim(dtc_aux31.Text)
'            End If
'            VAR_MOD = "P" + Trim(VAR_AUX) + "G" + Trim(dtc_codigo41.Text) + Trim(dtc_codigo91.Text) + "-" + Trim(dtc_campo51.Text)
'            rs_datos!modelo_codigo_h1 = Trim(VAR_MOD)
'            'Xizi   XO  1
'            If CDbl(dtc_valor41.Text) < 3 Then
'                VAR_MOD = "OH5000" + Trim(dtc_codigo31.Text) + Trim(dtc_codigo41.Text) + Trim(dtc_codigo51.Text) + Trim(dtc_codigo81.Text)
'            Else
'                VAR_MOD = "XO8000" + Trim(dtc_codigo31.Text) + Trim(dtc_codigo41.Text) + Trim(dtc_codigo51.Text) + Trim(dtc_codigo81.Text)
'            End If
'            rs_datos!modelo_codigo_x1 = Trim(VAR_MOD)
'            'Graba en Cotiza    1
'            Set rs_aux1 = New ADODB.Recordset
'            'SQL_FOR = "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'            SQL_FOR = "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'            rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'                Set rs_aux2 = New ADODB.Recordset
'                If rs_aux2.State = 1 Then rs_aux2.Close
'                rs_aux2.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'   ", db, adOpenStatic
'                If Not rs_aux2.EOF Then
'                     var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                End If
'                rs_aux1.AddNew
'                rs_aux1!ges_gestion = Year(Date)
'                rs_aux1!unidad_codigo = parametro   'Ado_datos.Recordset!unidad_codigo
'                rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
'                rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'                rs_aux1!trafico_codigo = Ado_datos.Recordset!trafico_codigo
'                rs_aux1!cotiza_codigo = var_cod
'                Call correl_bien
'                VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
'                rs_aux1!bien_codigo = VAR_COD3  '"36NO-" + Trim(Str(VAR_COD2))
'                'WWWWWWWWWWWWWW AQUI JQA 16-feb-2015 WWWWWWWWWWW
'                VAR_COD3 = "NA" + Trim(Str(i))
'                rs_aux1!bien_codigo = VAR_COD3  '"NA" + Trim(Str(VAR_COD2))
'                'WWWWWWWWWWWWWW AQUI JQA 16-feb-2015 WWWWWWWWWWW
'                rs_aux1!modelo_codigo = Ado_datos.Recordset!modelo_codigo
'                rs_aux1!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h1
'                rs_aux1!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x1
'                rs_aux1!estado_codigo = "REG"
'                rs_aux1!fecha_registro = Date
'                rs_aux1!usr_codigo = glusuario
'                rs_aux1.Update
'                db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod & " Where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'            'GRABA EN AC_BIENES
'            If i = 1 Then
'                VAR_NO1 = VAR_COD3
'                VAR_NO2 = VAR_COD2
'            End If
'            Call GRABA_BIENES
'        i = i + 1
'        Wend
'
'        If Txt_campo22.Text <> "" Then
'            i = 1
'            While (CDbl(Txt_campo32.Text) >= i)
'                'Brasil  2
'                VAR_MOD = Trim(dtc_codigo62.Text) + Trim(dtc_codigo72.Text) + Trim(dtc_codigo32.Text) + Trim(dtc_codigo42.Text) + Trim(dtc_codigo52.Text) + Trim(dtc_codigo82.Text) + Trim(dtc_codigo92.Text) + Trim(dtc_codigo02.Text)
'                rs_datos!modelo_codigo2 = Trim(VAR_MOD)
'                'Hypex  2
'                If CDbl(dtc_aux32.Text) < 1000 Then
'                    VAR_AUX = Trim("0" + Trim(dtc_aux32.Text))
'                Else
'                    VAR_AUX = Trim(dtc_aux32.Text)
'                End If
'                VAR_MOD = "P" + Trim(VAR_AUX) + "G" + Trim(dtc_codigo42.Text) + Trim(dtc_codigo92.Text) + "-" + Trim(dtc_campo52.Text)
'                rs_datos!modelo_codigo_h2 = Trim(VAR_MOD)
'                'Xizi   XO  2
'                If CDbl(dtc_desc42.Text) < 3 Then
'                    VAR_MOD = "OH5000" + Trim(dtc_codigo32.Text) + Trim(dtc_codigo42.Text) + Trim(dtc_codigo52.Text) + Trim(dtc_codigo82.Text)
'                Else
'                    VAR_MOD = "XO8000" + Trim(dtc_codigo32.Text) + Trim(dtc_codigo42.Text) + Trim(dtc_codigo52.Text) + Trim(dtc_codigo82.Text)
'                End If
'                rs_datos!modelo_codigo_x1 = Trim(VAR_MOD)
'                'Graba en Cotiza    2
'                Set rs_aux2 = New ADODB.Recordset
'                If rs_aux2.State = 1 Then rs_aux2.Close
'                rs_aux2.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'   ", db, adOpenStatic
'                If Not rs_aux2.EOF Then
'                     var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                End If
'                rs_aux1.AddNew
'                rs_aux1!ges_gestion = Year(Date)
'                rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
'                rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
'                rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'                rs_aux1!trafico_codigo = Ado_datos.Recordset!trafico_codigo
'                rs_aux1!cotiza_codigo = var_cod
'                Call correl_bien
'                VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
'                rs_aux1!bien_codigo = VAR_COD3
'                rs_aux1!modelo_codigo = Ado_datos.Recordset!modelo_codigo2
'                'rs_aux1!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h2
'                'rs_aux1!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x2
'                rs_aux1!estado_codigo = "REG"
'                rs_aux1!fecha_registro = Date
'                rs_aux1!usr_codigo = glusuario
'                rs_aux1.Update
'                db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'                'GRABA EN AC_BIENES
'                Call GRABA_BIENES
'            i = i + 1
'            Wend
'        End If
'
'        If Txt_campo23.Text <> "" Then
'            i = 1
'            While (CDbl(Txt_campo33.Text) >= i)
'                'Brasil  3
'                VAR_MOD = Trim(dtc_codigo63.Text) + Trim(dtc_codigo73.Text) + Trim(dtc_codigo33.Text) + Trim(dtc_codigo43.Text) + Trim(dtc_codigo53.Text) + Trim(dtc_codigo83.Text) + Trim(dtc_codigo93.Text) + Trim(dtc_codigo03.Text)
'                rs_datos!modelo_codigo3 = Trim(VAR_MOD)
'                'Hypex  3
'                If CDbl(dtc_aux33.Text) < 1000 Then
'                    VAR_AUX = Trim("0" + Trim(dtc_aux33.Text))
'                Else
'                    VAR_AUX = Trim(dtc_aux33.Text)
'                End If
'                VAR_MOD = "P" + Trim(VAR_AUX) + "G" + Trim(dtc_codigo43.Text) + Trim(dtc_codigo93.Text) + "-" + Trim(dtc_campo53.Text)
'                rs_datos!modelo_codigo_h3 = Trim(VAR_MOD)
'                'Xizi   XO  3
'                If CDbl(dtc_desc43.Text) < 3 Then
'                    VAR_MOD = "OH5000" + Trim(dtc_codigo33.Text) + Trim(dtc_codigo43.Text) + Trim(dtc_codigo53.Text) + Trim(dtc_codigo83.Text)
'                Else
'                    VAR_MOD = "XO8000" + Trim(dtc_codigo33.Text) + Trim(dtc_codigo43.Text) + Trim(dtc_codigo53.Text) + Trim(dtc_codigo83.Text)
'                End If
'                rs_datos!modelo_codigo_x3 = Trim(VAR_MOD)
'                'Graba en Cotiza    3
'                Set rs_aux2 = New ADODB.Recordset
'                If rs_aux2.State = 1 Then rs_aux2.Close
'                rs_aux2.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'   ", db, adOpenStatic
'                If Not rs_aux2.EOF Then
'                     var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                End If
'                rs_aux1.AddNew
'                rs_aux1!ges_gestion = Year(Date)
'                rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
'                rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
'                rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'                rs_aux1!trafico_codigo = Ado_datos.Recordset!trafico_codigo
'                rs_aux1!cotiza_codigo = var_cod
'                Call correl_bien
'                VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
'                rs_aux1!bien_codigo = VAR_COD3
'                rs_aux1!modelo_codigo = Ado_datos.Recordset!modelo_codigo3
'                rs_aux1!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h3
'                rs_aux1!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x3
'                rs_aux1!estado_codigo = "REG"
'                rs_aux1!fecha_registro = Date
'                rs_aux1!usr_codigo = glusuario
'                rs_aux1.Update
'                db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'                'GRABA EN AC_BIENES
'                Call GRABA_BIENES
'            i = i + 1
'            Wend
'        End If
'
'        If Txt_campo24.Text <> "" Then
'            i = 1
'            While (CDbl(Txt_campo34.Text) >= i)
'                'Brasil  4
'                VAR_MOD = Trim(dtc_codigo64.Text) + Trim(dtc_codigo74.Text) + Trim(dtc_codigo34.Text) + Trim(dtc_codigo44.Text) + Trim(dtc_codigo54.Text) + Trim(dtc_codigo84.Text) + Trim(dtc_codigo94.Text) + Trim(dtc_codigo04.Text)
'                rs_datos!modelo_codigo4 = Trim(VAR_MOD)
'                'Hypex  4
'                If CDbl(dtc_aux34.Text) < 1000 Then
'                    VAR_AUX = Trim("0" + Trim(dtc_aux34.Text))
'                Else
'                    VAR_AUX = Trim(dtc_aux34.Text)
'                End If
'                VAR_MOD = "P" + Trim(VAR_AUX) + "G" + Trim(dtc_codigo44.Text) + Trim(dtc_codigo94.Text) + "-" + Trim(dtc_campo54.Text)
'                rs_datos!modelo_codigo_h4 = Trim(VAR_MOD)
'                'Xizi   XO  4
'                If CDbl(dtc_desc44.Text) < 3 Then
'                    VAR_MOD = "OH5000" + Trim(dtc_codigo34.Text) + Trim(dtc_codigo44.Text) + Trim(dtc_codigo54.Text) + Trim(dtc_codigo84.Text)
'                Else
'                    VAR_MOD = "XO8000" + Trim(dtc_codigo34.Text) + Trim(dtc_codigo44.Text) + Trim(dtc_codigo54.Text) + Trim(dtc_codigo84.Text)
'                End If
'                rs_datos!modelo_codigo_x4 = Trim(VAR_MOD)
'                'Graba en Cotiza    4
'                Set rs_aux2 = New ADODB.Recordset
'                If rs_aux2.State = 1 Then rs_aux2.Close
'                rs_aux2.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'   ", db, adOpenStatic
'                If Not rs_aux2.EOF Then
'                     var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                End If
'                rs_aux1.AddNew
'                rs_aux1!ges_gestion = Year(Date)
'                rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
'                rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
'                rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'                rs_aux1!trafico_codigo = Ado_datos.Recordset!trafico_codigo
'                rs_aux1!cotiza_codigo = var_cod
'                Call correl_bien
'                VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
'                rs_aux1!bien_codigo = VAR_COD3
'                rs_aux1!modelo_codigo = Ado_datos.Recordset!modelo_codigo4
'                rs_aux1!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h4
'                rs_aux1!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x4
'                rs_aux1!estado_codigo = "REG"
'                rs_aux1!fecha_registro = Date
'                rs_aux1!usr_codigo = glusuario
'                rs_aux1.Update
'                db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'                'GRABA EN AC_BIENES
'                Call GRABA_BIENES
'            i = i + 1
'            Wend
'        End If
'        'VAR_VAL,
'        VAR_NO2 = VAR_NO2 + rs_datos!h_nro_total_equipos - 1
'        VAR_NO3 = "36NO-" + Trim(Str(VAR_NO2))
'        If rs_datos!h_nro_total_equipos > 1 Then
'            'If Right(VAR_NO3, 1) = 0 Then
'                rs_datos!unidad_codigo_ant = VAR_NO1 + "-" + Right(VAR_NO3, 2)
'            'Else
'            '    rs_datos!unidad_codigo_ant = VAR_NO1 + "/" + Right(VAR_NO3, 1)
'            'End If
'        Else
'            rs_datos!unidad_codigo_ant = VAR_NO1
'        End If
'        db.Execute "Update ao_solicitud Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'        db.Execute "Update ao_solicitud_cotiza_venta Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'        db.Execute "Update ao_negociacion_cabecera Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and negocia_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'
'        rs_datos!estado_codigo_verif = "APR"
'        rs_datos!fecha_registro = Date
'        rs_datos!usr_codigo = glusuario
'        rs_datos.UpdateBatch adAffectAll
'      End If
'   Else
'       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene detalle ...", vbExclamation, "Validación de Registro"
'   End If
'   Exit Sub
'UpdateErr:
'  MsgBox Err.Description

    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexión = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = dg_datos
    ClBuscaGrid.QueryUtilizado = queryinicial
    Set ClBuscaGrid.RecordsetTrabajo = rs_datos
    'ClBuscaGrid.CamposVisibles = "11010011"
    ClBuscaGrid.Ejecutar
End Sub

Private Sub BtnCancelar_Click()
    On Error GoTo AddErr
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelUpdate
'        Call ABRIR_TABLA
        'rs_datos.MoveFirst
        FraNavega.Enabled = True
        Fra_datos.Enabled = False
        Frame2.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        VAR_SW = ""
    End If
Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
     rs_datos!trafico_fecha = IIf(DTPfecha1.Value = "01/01/1900", Date, DTPfecha1.Value)
     'A.
     rs_datos!arreglo1 = 1
     rs_datos!trafico_num_paradas = Txt_campo21.Text
     rs_datos!recorrido_codigo = Txt_aux21.Text   'dtc_codigo21.Text
     rs_datos!pasajeros_codigo = dtc_codigo31.Text
'     If CDbl(dtc_desc31.Text) = 0 Or dtc_desc31.Text = "" Then
'        MsgBox "Vuelva a registrar el: " + lbl_campo3.Caption
'        Txt_campo31 = Round(CDbl(lbl_campoh42.Caption) / 1, 0)
'     Else
'        Txt_campo31 = Round(CDbl(lbl_campoh42.Caption) / CDbl(dtc_desc31.Text), 0)
'     End If
     rs_datos!trafico_nro_equipos = Txt_campo31.Text
     rs_datos!vel_equipo_codigo = dtc_codigo41.Text
     rs_datos!tipo_puerta = dtc_codigo51.Text
     rs_datos!trafico_ancho_puerta = Txt_campo41.Text
     
     'B.
     rs_datos!cabina_codigo = dtc_codigo61.Text
     rs_datos!tecnologia_codigo = dtc_codigo71.Text
     rs_datos!sist_puerta = dtc_codigo81.Text
     rs_datos!condicion_ventas = dtc_codigo91.Text
     rs_datos!condicion_cabina = dtc_codigo01.Text
     rs_datos!ctrlmaq_codigo = dtc_codigo11.Text
     rs_datos!pais_continente = dtc_codigo011.Text
     'C.
     rs_datos!c_time_asc_desaceleracion = dtc_aux41.Text  'lbl_campoc11.Caption
     rs_datos!c_time_apertura_cierre = dtc_aux51.Text
     If CDbl(Txt_campo41.Text) < 1100 Then var_campoc31 = "2.4" Else var_campoc31 = "2" 'End If
     rs_datos!c_time_entrada_salida = Round(var_campoc31, 2)    'var_campoe31
     'D.
     '=+SI(B40>0;B9-(B9-1)*((B9-2)/(B9-1))^B39;0)
     If CDbl(Txt_campo31.Text) > 0 Then
        var_campod11 = Round(CDbl(Txt_campo21.Text) - (CDbl(Txt_campo21.Text) - 1) * ((CDbl(Txt_campo21.Text) - 2) / (CDbl(Txt_campo21.Text) - 1)) ^ CDbl(dtc_desc31.Text), 2)
     Else
        var_campod11 = 0
     End If
     rs_datos!d_num_paradas_probables = Round(var_campod11, 2)
     'E.
        '=+B10*2/B41
        var_campoe11 = CDbl(Txt_aux21.Text) * 2 / CDbl(dtc_valor41.Text)
     rs_datos!e_tiempo_recorrido = Round(var_campoe11, 2)
        '=+B53*B47/2
        var_campoe21 = CDbl(var_campod11) * CDbl(dtc_aux41.Text) / 2
     rs_datos!e_tiempo_asc_desaceleracion = Round(var_campoe21, 2)
        '=+B53*B48*1.1
        var_campoe31 = CDbl(var_campod11) * CDbl(dtc_aux51.Text) * 1.1
     rs_datos!e_tiempo_apertura_cierre = Round(var_campoe31, 2)
        '=+B39*B49*1.1
        var_campoe41 = CDbl(dtc_desc31.Text) * CDbl(var_campoc31) * 1.1
     rs_datos!e_tiempo_entrada_salida = Round(var_campoe41, 2)
     If CDbl(Txt_campo31.Text) > 0 Then
        rs_datos!e_tiempo_total = CDbl(var_campoe11) + CDbl(var_campoe21) + CDbl(var_campoe31) + CDbl(var_campoe41)
     Else
        rs_datos!e_tiempo_total = "0"
     End If
     'F.
     If rs_datos!e_tiempo_total > 0 Then
        rs_datos!f_tiempo_recorrido = Round(CDbl(var_campoe11) / rs_datos!e_tiempo_total, 2)
        rs_datos!f_time_asc_desaceleracion = Round(CDbl(var_campoe21) / rs_datos!e_tiempo_total, 2)
        rs_datos!f_time_apertura_cierre = Round(CDbl(var_campoe31) / rs_datos!e_tiempo_total, 2)
        rs_datos!f_time_entrada_salida = Round(CDbl(var_campoe41) / rs_datos!e_tiempo_total, 2)
     Else
        rs_datos!f_tiempo_recorrido = 0
        rs_datos!f_time_asc_desaceleracion = 0
        rs_datos!f_time_apertura_cierre = 0
        rs_datos!f_time_entrada_salida = 0
     End If
     'G.
        '=300*B39/B59
        var_campog11 = 300 * CDbl(dtc_desc31.Text) / rs_datos!e_tiempo_total
     rs_datos!g_capacidad_tiempo_cti = Round(var_campog11, 2)
        '=+SI(B40>0;B69*B40;0)
        If Txt_campo31.Text > 0 Then
           var_campog21 = CDbl(var_campog11) * CDbl(Txt_campo31.Text)
        Else
           var_campog21 = 0
        End If
     rs_datos!g_capacidad_total_arreglo = Round(var_campog21, 2)
     'TOOLTIPTEXT 1
     dtc_desc51.ToolTipText = dtc_desc51.Text
     dtc_desc61.ToolTipText = dtc_desc61.Text
     dtc_desc71.ToolTipText = dtc_desc71.Text
     dtc_desc81.ToolTipText = dtc_desc81.Text
     dtc_desc91.ToolTipText = dtc_desc91.Text
     dtc_desc01.ToolTipText = dtc_desc01.Text
     dtc_desc11.ToolTipText = dtc_desc11.Text
     
     'A.    2
   If Txt_campo22.Text <> "" Then
     rs_datos!arreglo2 = 2
     rs_datos!trafico_num_paradas2 = Txt_campo22.Text
     rs_datos!recorrido_codigo2 = Txt_aux22.Text    ' dtc_codigo22.Text
     rs_datos!pasajeros_codigo2 = dtc_codigo32.Text
'     If CDbl(dtc_desc32.Text) = 0 Or dtc_desc32.Text = "" Then
'        'Txt_campo32 = Round(CDbl(lbl_campoh42.Caption) / CDbl(dtc_desc32.Text), 0)
'        MsgBox "Vuelva a registrar el: " + lbl_campo3.Caption
'        Txt_campo32 = Round(CDbl(lbl_campoh42.Caption) / 1, 0)
'     Else
'        'MsgBox "Vuelva a registrar el: " + lbl_campo3.Caption
'        'Txt_campo32 = Round(CDbl(lbl_campoh42.Caption) / 1, 0)
'        Txt_campo32 = Round(CDbl(lbl_campoh42.Caption) / CDbl(dtc_desc32.Text), 0)
'     End If
     rs_datos!trafico_nro_equipos2 = Txt_campo32.Text       'Txt_campo31
     rs_datos!vel_equipo_codigo2 = dtc_codigo42.Text
     rs_datos!tipo_puerta2 = dtc_codigo52.Text
     rs_datos!trafico_ancho_puerta2 = Txt_campo42.Text
     'B.    2
     rs_datos!cabina_codigo2 = dtc_codigo62.Text
     rs_datos!tecnologia_codigo2 = dtc_codigo72.Text
     rs_datos!sist_puerta2 = dtc_codigo82.Text
     rs_datos!condicion_ventas2 = dtc_codigo92.Text
     rs_datos!condicion_cabina2 = dtc_codigo02.Text
     rs_datos!ctrlmaq_codigo2 = dtc_codigo12.Text
     rs_datos!pais_continente2 = dtc_codigo012.Text
     'C.    2
     rs_datos!c_time_asc_desaceleracion2 = dtc_aux42.Text
     rs_datos!c_time_apertura_cierre2 = dtc_aux52.Text
     If CDbl(Txt_campo42.Text) < 1100 Then var_campoc32 = "2.4" Else var_campoc32 = "2" 'End If
     rs_datos!c_time_entrada_salida2 = var_campoc32
     'D.    2
     '=+SI(B40>0;B9-(B9-1)*((B9-2)/(B9-1))^B39;0)
     If CDbl(Txt_campo32.Text) > 0 Then
        var_campod12 = Round(CDbl(Txt_campo22.Text) - (CDbl(Txt_campo22.Text) - 1) * ((CDbl(Txt_campo22.Text) - 2) / (CDbl(Txt_campo22.Text) - 1)) ^ CDbl(dtc_desc32.Text), 2)
     Else
        var_campod12 = 0
     End If
     rs_datos!d_num_paradas_probables2 = Round(var_campod12, 2)
     'E.    2
        '=+B10*2/B41
        var_campoe12 = CDbl(Txt_aux22.Text) * 2 / CDbl(dtc_desc42.Text)
     rs_datos!e_tiempo_recorrido2 = Round(var_campoe12, 2)
        '=+B53*B47/2
        var_campoe22 = CDbl(var_campod12) * CDbl(dtc_aux42.Text) / 2
     rs_datos!e_tiempo_asc_desaceleracion2 = Round(var_campoe22, 2)
        '=+B53*B48*1.1
        var_campoe32 = CDbl(var_campod12) * CDbl(dtc_aux52.Text) * 1.1
     rs_datos!e_tiempo_apertura_cierre2 = Round(var_campoe32, 2)
        '=+B39*B49*1.1
        var_campoe42 = CDbl(dtc_desc32.Text) * CDbl(var_campoc32) * 1.1
     rs_datos!e_tiempo_entrada_salida2 = Round(var_campoe42, 2)
     If Txt_campo32.Text > 0 Then
        rs_datos!e_tiempo_total2 = Round(CDbl(var_campoe12) + CDbl(var_campoe22) + CDbl(var_campoe32) + CDbl(var_campoe42), 2)
     Else
        rs_datos!e_tiempo_total2 = "0"
     End If
     'F.    2
     If rs_datos!e_tiempo_total2 > 0 Then
        rs_datos!f_tiempo_recorrido2 = Round(CDbl(var_campoe12) / rs_datos!e_tiempo_total2, 2)
        rs_datos!f_time_asc_desaceleracion2 = Round(CDbl(var_campoe22) / rs_datos!e_tiempo_total2, 2)
        rs_datos!f_time_apertura_cierre2 = Round(CDbl(var_campoe32) / rs_datos!e_tiempo_total2, 2)
        rs_datos!f_time_entrada_salida2 = Round(CDbl(var_campoe42) / rs_datos!e_tiempo_total2, 2)
     Else
        rs_datos!f_tiempo_recorrido2 = 0
        rs_datos!f_time_asc_desaceleracion2 = 0
        rs_datos!f_time_apertura_cierre2 = 0
        rs_datos!f_time_entrada_salida2 = 0
     End If
     'G.    2
        '=300*B39/B59
        var_campog12 = 300 * CDbl(dtc_desc32.Text) / rs_datos!e_tiempo_total
     rs_datos!g_capacidad_tiempo_cti2 = Round(var_campog12, 2)
        '=+SI(B40>0;B69*B40;0)
        If Txt_campo32.Text > 0 Then
           var_campog22 = CDbl(var_campog12) * CDbl(Txt_campo32.Text)
        Else
           var_campog22 = 0
        End If
     rs_datos!g_capacidad_total_arreglo2 = Round(var_campog22, 2)
   End If
   'TOOLTIPTEXT 2
     dtc_desc52.ToolTipText = dtc_desc52.Text
     dtc_desc62.ToolTipText = dtc_desc62.Text
     dtc_desc72.ToolTipText = dtc_desc72.Text
     dtc_desc82.ToolTipText = dtc_desc82.Text
     dtc_desc92.ToolTipText = dtc_desc92.Text
     dtc_desc02.ToolTipText = dtc_desc02.Text
     dtc_desc12.ToolTipText = dtc_desc12.Text
     
     'A.    3
   If Txt_campo23.Text <> "" Then
     rs_datos!arreglo3 = 3
     rs_datos!trafico_num_paradas3 = Txt_campo23.Text
     rs_datos!recorrido_codigo3 = Txt_aux23.Text    'dtc_codigo23.Text
     rs_datos!pasajeros_codigo3 = dtc_codigo33.Text
'     If CDbl(dtc_desc33.Text) = 0 Or dtc_desc33.Text = "" Then
'        MsgBox "Vuelva a registrar el: " + lbl_campo3.Caption
'        Txt_campo33 = Round(CDbl(lbl_campoh42.Caption) / 1, 0)
'     Else
'        Txt_campo33 = Round(CDbl(lbl_campoh42.Caption) / CDbl(dtc_desc33.Text), 0)
'     End If
     rs_datos!trafico_nro_equipos3 = Txt_campo33.Text
     rs_datos!vel_equipo_codigo3 = dtc_codigo43.Text
     rs_datos!tipo_puerta3 = dtc_codigo53.Text
     rs_datos!trafico_ancho_puerta3 = Txt_campo43.Text
     'B.    3
     rs_datos!cabina_codigo3 = dtc_codigo63.Text
     rs_datos!tecnologia_codigo3 = dtc_codigo73.Text
     rs_datos!sist_puerta3 = dtc_codigo83.Text
     rs_datos!condicion_ventas3 = dtc_codigo93.Text
     rs_datos!condicion_cabina3 = dtc_codigo03.Text
     rs_datos!ctrlmaq_codigo3 = dtc_codigo13.Text
     rs_datos!pais_continente3 = dtc_codigo013.Text
     'C.    3
     rs_datos!c_time_asc_desaceleracion3 = dtc_aux43.Text  'lbl_campoc11.Caption
     rs_datos!c_time_apertura_cierre3 = dtc_aux53.Text
     If CDbl(Txt_campo43.Text) < 1100 Then var_campoc33 = "2.4" Else var_campoc33 = "2" 'End If
     rs_datos!c_time_entrada_salida3 = var_campoc33
     'D.    3
     '=+SI(B40>0;B9-(B9-1)*((B9-2)/(B9-1))^B39;0)
     If CDbl(Txt_campo33.Text) > 0 Then
        var_campod13 = CDbl(Txt_campo23.Text) - (CDbl(Txt_campo23.Text) - 1) * ((CDbl(Txt_campo23.Text) - 2) / (CDbl(Txt_campo23.Text) - 1)) ^ CDbl(dtc_desc33.Text)
     Else
        var_campod13 = 0
     End If
     rs_datos!d_num_paradas_probables3 = Round(var_campod13, 2)
     'E.    3
        '=+B10*2/B41
        var_campoe13 = CDbl(Txt_aux23.Text) * 2 / CDbl(dtc_desc43.Text)
     rs_datos!e_tiempo_recorrido3 = Round(var_campoe13, 2)
        '=+B53*B47/2
        var_campoe23 = CDbl(var_campod13) * CDbl(dtc_aux43.Text) / 2
     rs_datos!e_tiempo_asc_desaceleracion3 = Round(var_campoe23, 2)
        '=+B53*B48*1.1
        var_campoe33 = CDbl(var_campod13) * CDbl(dtc_aux53.Text) * 1.1
     rs_datos!e_tiempo_apertura_cierre3 = Round(var_campoe33, 2)
        '=+B39*B49*1.1
        var_campoe43 = CDbl(dtc_desc33.Text) * CDbl(var_campoc33) * 1.1
     rs_datos!e_tiempo_entrada_salida3 = Round(var_campoe43, 2)
     If CDbl(Txt_campo33.Text) > 0 Then
        rs_datos!e_tiempo_total3 = CDbl(var_campoe13) + CDbl(var_campoe23) + CDbl(var_campoe33) + CDbl(var_campoe43)
     Else
        rs_datos!e_tiempo_total3 = "0"
     End If
     'F.    3
     If rs_datos!e_tiempo_total3 > 0 Then
        rs_datos!f_tiempo_recorrido3 = Round(CDbl(var_campoe13) / rs_datos!e_tiempo_total3, 2)
        rs_datos!f_time_asc_desaceleracion3 = Round(CDbl(var_campoe23) / rs_datos!e_tiempo_total3, 2)
        rs_datos!f_time_apertura_cierre3 = Round(CDbl(var_campoe33) / rs_datos!e_tiempo_total3, 2)
        rs_datos!f_time_entrada_salida3 = Round(CDbl(var_campoe43) / rs_datos!e_tiempo_total3, 2)
     Else
        rs_datos!f_tiempo_recorrido3 = 0
        rs_datos!f_time_asc_desaceleracion3 = 0
        rs_datos!f_time_apertura_cierre3 = 0
        rs_datos!f_time_entrada_salida3 = 0
     End If
     'G.    3
        '=300*B39/B59
        var_campog13 = 300 * CDbl(dtc_desc33.Text) / rs_datos!e_tiempo_total
     rs_datos!g_capacidad_tiempo_cti3 = Round(var_campog13, 2)
        '=+SI(B40>0;B69*B40;0)
        If Txt_campo33.Text > 0 Then
           var_campog23 = CDbl(var_campog13) * CDbl(Txt_campo33.Text)
        Else
           var_campog23 = 0
        End If
     rs_datos!g_capacidad_total_arreglo3 = Round(var_campog23, 2)
    End If
    'TOOLTIPTEXT 3
     dtc_desc53.ToolTipText = dtc_desc53.Text
     dtc_desc63.ToolTipText = dtc_desc63.Text
     dtc_desc73.ToolTipText = dtc_desc73.Text
     dtc_desc83.ToolTipText = dtc_desc83.Text
     dtc_desc93.ToolTipText = dtc_desc93.Text
     dtc_desc03.ToolTipText = dtc_desc03.Text
     dtc_desc13.ToolTipText = dtc_desc13.Text
        
     'A.    4
   If Txt_campo24.Text <> "" Then
     rs_datos!arreglo4 = 4
     rs_datos!trafico_num_paradas4 = Txt_campo24.Text
     rs_datos!recorrido_codigo4 = Txt_aux24.Text          'dtc_codigo24.Text
     rs_datos!pasajeros_codigo4 = dtc_codigo34.Text
'     If CDbl(dtc_desc34.Text) = 0 Or dtc_desc34.Text = "" Then
'        MsgBox "Vuelva a registrar el: " + lbl_campo4.Caption
'        Txt_campo34 = Round(CDbl(lbl_campoh42.Caption) / 1, 0)
'     Else
'        Txt_campo34 = Round(CDbl(lbl_campoh42.Caption) / CDbl(dtc_desc34.Text), 0)
'     End If
     rs_datos!trafico_nro_equipos4 = Txt_campo34.Text
     rs_datos!vel_equipo_codigo4 = dtc_codigo44.Text
     rs_datos!tipo_puerta4 = dtc_codigo54.Text
     rs_datos!trafico_ancho_puerta4 = Txt_campo44.Text
     'B.    4
     rs_datos!cabina_codigo4 = dtc_codigo64.Text
     rs_datos!tecnologia_codigo4 = dtc_codigo74.Text
     rs_datos!sist_puerta4 = dtc_codigo84.Text
     rs_datos!condicion_ventas4 = dtc_codigo94.Text
     rs_datos!condicion_cabina4 = dtc_codigo04.Text
     rs_datos!ctrlmaq_codigo4 = dtc_codigo14.Text
     rs_datos!pais_continente4 = dtc_codigo014.Text
     'C.    4
     rs_datos!c_time_asc_desaceleracion4 = dtc_aux44.Text  'lbl_campoc11.Caption
     rs_datos!c_time_apertura_cierre4 = dtc_aux54.Text
     If CDbl(Txt_campo44.Text) < 1100 Then var_campoc34 = "2.4" Else var_campoc34 = "2" 'End If
     rs_datos!c_time_entrada_salida4 = var_campoc34
     'D.    4
     '=+SI(B40>0;B9-(B9-1)*((B9-2)/(B9-1))^B39;0)
     If CDbl(Txt_campo34.Text) > 0 Then
        var_campod14 = CDbl(Txt_campo24.Text) - (CDbl(Txt_campo24.Text) - 1) * ((CDbl(Txt_campo24.Text) - 2) / (CDbl(Txt_campo24.Text) - 1)) ^ CDbl(dtc_desc34.Text)
     Else
        var_campod14 = 0
     End If
     rs_datos!d_num_paradas_probables4 = Round(var_campod14, 2)
     'E.    4
        '=+B10*2/B41
        var_campoe14 = CDbl(Txt_aux24.Text) * 2 / CDbl(dtc_desc44.Text)
     rs_datos!e_tiempo_recorrido4 = Round(var_campoe14, 2)
        '=+B53*B47/2
        var_campoe24 = CDbl(var_campod14) * CDbl(dtc_aux44.Text) / 2
     rs_datos!e_tiempo_asc_desaceleracion4 = Round(var_campoe24, 2)
        '=+B53*B48*1.1
        var_campoe34 = CDbl(var_campod14) * CDbl(dtc_aux54.Text) * 1.1
     rs_datos!e_tiempo_apertura_cierre4 = Round(var_campoe34, 2)
        '=+B39*B49*1.1
        var_campoe44 = CDbl(dtc_desc34.Text) * CDbl(var_campoc34) * 1.1
     rs_datos!e_tiempo_entrada_salida4 = Round(var_campoe44, 2)
     If CDbl(Txt_campo34.Text) > 0 Then
        rs_datos!e_tiempo_total4 = CDbl(var_campoe14) + CDbl(var_campoe24) + CDbl(var_campoe34) + CDbl(var_campoe44)
     Else
        rs_datos!e_tiempo_total4 = "0"
     End If
     'F.    4
     If rs_datos!e_tiempo_total4 > 0 Then
        rs_datos!f_tiempo_recorrido4 = Round(CDbl(var_campoe14) / rs_datos!e_tiempo_total4, 2)
        rs_datos!f_time_asc_desaceleracion4 = Round(CDbl(var_campoe24) / rs_datos!e_tiempo_total4, 2)
        rs_datos!f_time_apertura_cierre4 = Round(CDbl(var_campoe34) / rs_datos!e_tiempo_total4, 2)
        rs_datos!f_time_entrada_salida4 = Round(CDbl(var_campoe44) / rs_datos!e_tiempo_total4, 2)
     Else
        rs_datos!f_tiempo_recorrido4 = 0
        rs_datos!f_time_asc_desaceleracion4 = 0
        rs_datos!f_time_apertura_cierre4 = 0
        rs_datos!f_time_entrada_salida4 = 0
     End If
     'G.    4
        '=300*B39/B59
        var_campog14 = 300 * CDbl(dtc_desc34.Text) / rs_datos!e_tiempo_total
     rs_datos!g_capacidad_tiempo_cti4 = Round(var_campog14, 2)
        '=+SI(B40>0;B69*B40;0)
        If Txt_campo34.Text > 0 Then
           var_campog24 = CDbl(var_campog14) * CDbl(Txt_campo34.Text)
        Else
           var_campog24 = 0
        End If
     rs_datos!g_capacidad_total_arreglo4 = Round(var_campog24, 2)
    End If
    'TOOLTIPTEXT 4
     dtc_desc54.ToolTipText = dtc_desc54.Text
     dtc_desc64.ToolTipText = dtc_desc64.Text
     dtc_desc74.ToolTipText = dtc_desc74.Text
     dtc_desc84.ToolTipText = dtc_desc84.Text
     dtc_desc94.ToolTipText = dtc_desc94.Text
     dtc_desc04.ToolTipText = dtc_desc04.Text
     dtc_desc14.ToolTipText = dtc_desc14.Text

     'H.
     lbl_campoh11.Caption = Val(Txt_campo31.Text) + Val(IIf(Txt_campo32.Text = "", "0", Txt_campo32.Text)) + Val(IIf(Txt_campo33.Text = "", "0", Txt_campo33.Text)) + Val(IIf(Txt_campo34.Text = "", "0", Txt_campo34.Text))
     rs_datos!h_nro_total_equipos = lbl_campoh11.Caption
     rs_datos!h_nro_total_equipos_parametro = lbl_campoh11.Caption
     lbl_campoh13.Caption = "CORRECTO"
     rs_datos!h_nro_total_equipos_result = "CORRECTO"
        
        '=+SUMA(B70:E70)
     'lbl_campoh41.Caption = CDbl(var_campog21) + CDbl(var_campog22) + CDbl(var_campog23) + CDbl(lbl_campog24.Caption)
     lbl_campoh41.Caption = CDbl(var_campog21) + CDbl(IIf(var_campog22 = "", "0", var_campog22)) + CDbl(IIf(var_campog23 = "", "0", var_campog23))
     rs_datos!h_capacidad_trafico = Round(lbl_campoh41.Caption, 2)
     rs_datos!h_capacidad_trafico_parametro = lbl_campoh42.Caption
     '   lbl_campoh13.Caption = "CORRECTO"
     If CDbl(lbl_campoh41.Caption) > CDbl(lbl_campoh42.Caption) Then
        rs_datos!h_capacidad_trafico_result = "CORRECTO"
     Else
        rs_datos!h_capacidad_trafico_result = "INCORRECTO"
     End If
     lbl_campoh43.Caption = rs_datos!h_capacidad_trafico_result
        
        '=3600*(B53+C53+D53+E53)/(B59+C59+D59+E59)*D75/B75/B72
        'lbl_campoh21.Caption = 3600 * (CDbl(var_campod11) + CDbl(var_campod12) + CDbl(var_campod13) + CDbl(lbl_campod14.Caption)) / (rs_datos!e_tiempo_total + rs_datos!e_tiempo_total2 + rs_datos!e_tiempo_total3 + rs_datos!e_tiempo_total4) * CDbl(lbl_campoh42.Caption) / CDbl(lbl_campoh41.Caption) / CDbl(lbl_campoh11.Caption)
        lbl_campoh21.Caption = 3600 * (CDbl(var_campod11) + CDbl(IIf(var_campod12 = "", "0", var_campod12)) + CDbl(IIf(var_campod13 = "", "0", var_campod13))) / (rs_datos!e_tiempo_total + IIf(IsNull(rs_datos!e_tiempo_total2), "0", rs_datos!e_tiempo_total2) + IIf(IsNull(rs_datos!e_tiempo_total3), "0", rs_datos!e_tiempo_total3)) * CDbl(lbl_campoh42.Caption) / CDbl(lbl_campoh41.Caption) / CDbl(lbl_campoh11.Caption)
     rs_datos!h_partidas_por_hora = Round(lbl_campoh21.Caption, 2)
     rs_datos!h_partida_por_hora_parametro = Round(lbl_campoh21.Caption, 2)
        lbl_campoh23.Caption = "CORRECTO"
     rs_datos!h_partida_por_hora_result = "CORRECTO"
     
        '=+(B59*B40+C59*C40+D59*D40+E59*E40)/B72/B72
        'lbl_campoh31.Caption = (rs_datos!e_tiempo_total * CDbl(Txt_campo31.Text) + rs_datos!e_tiempo_total2 * CDbl(Txt_campo32.Text) + rs_datos!e_tiempo_total3 * CDbl(Txt_campo33.Text) + rs_datos!e_tiempo_total4 * CDbl(Txt_campo34.Text)) / CDbl(lbl_campoh11.Caption) / CDbl(lbl_campoh11.Caption)
        lbl_campoh31.Caption = (rs_datos!e_tiempo_total * CDbl(Txt_campo31.Text) + IIf(IsNull(rs_datos!e_tiempo_total2), "0", rs_datos!e_tiempo_total2) * CDbl(IIf(Txt_campo32.Text = "", "0", Txt_campo32.Text)) + IIf(IsNull(rs_datos!e_tiempo_total3), "0", rs_datos!e_tiempo_total3) * CDbl(IIf(Txt_campo33.Text = "", "0", Txt_campo33.Text))) / CDbl(lbl_campoh11.Caption) / CDbl(lbl_campoh11.Caption)
     rs_datos!h_intervalo_trafico = Round(lbl_campoh31.Caption, 2)
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "Select intervalo_trafico_parametro as Codigo from ac_bienes_equipo_intervalos_trafico where edif_tipo = '" & dtc_aux3.Text & "' and nro_total_equipos = " & CDbl(lbl_campoh11.Caption) & "   ", db, adOpenStatic
        'rs_aux2.Open "Select intervalo_trafico_parametro as Codigo from ac_bienes_equipo_intervalos_trafico where edif_tipo = '" & dtc_aux3.Text & "' and nro_total_equipos = '3' ", db, adOpenStatic
        If Not rs_aux2.EOF Then
            VAR_AUX = rs_aux2!Codigo
        End If
     rs_datos!h_intervalo_trafico_parametro = VAR_AUX
     If CDbl(lbl_campoh31.Caption) < VAR_AUX Then
        rs_datos!h_intervalo_trafico_result = "CORRECTO"
     Else
        rs_datos!h_intervalo_trafico_result = "INCORRECTO"
     End If
     
     'hora_registro
     If parametro = "DNMOD" Then
        rs_datos!proceso_codigo = "TEC"
        rs_datos!subproceso_codigo = "TEC-05"
        rs_datos!etapa_codigo = "TEC-05-03"
        rs_datos!clasif_codigo = "TEC"
        rs_datos!doc_codigo = "R-313"
        rs_datos!doc_numero = Txt_campo1.Text
        rs_datos!poa_codigo = "3.2.7"
        'REVISAR !!! JQA 2014_07_08
        rs_datos!archivo_respaldo = "TEC_R313-" + Trim(Ado_datos.Recordset!trafico_codigo) + ".PDF"
     Else
        rs_datos!proceso_codigo = "COM"
        rs_datos!subproceso_codigo = "COM-01"
        rs_datos!etapa_codigo = "COM-01-03"
        rs_datos!clasif_codigo = "COM"
        rs_datos!doc_codigo = "R-221"
        rs_datos!doc_numero = Txt_campo1.Text
        rs_datos!poa_codigo = "3.1.1"
        'REVISAR !!! JQA 2014_07_08
        rs_datos!archivo_respaldo = "COM_R221-" + Trim(Ado_datos.Recordset!trafico_codigo) + ".PDF"
     End If
     rs_datos!archivo_respaldo_cargado = "N"
     If rs_datos!estado_codigo_verif <> "APR" Then
        rs_datos!estado_codigo_verif = "REG"
     End If
     
     rs_datos!fecha_registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     rs_datos.Update 'adAffectAll 'Batch
'    If Ado_datos.Recordset!estado_codigo = "REG" Then
'        Call OptFilGral1_Click
'    Else
'        Call OptFilGral2_Click
'    End If
     'rs_datos.MoveLast
'     mbDataChanged = False

     Fra_datos.Enabled = False
     Frame2.Enabled = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     FraNavega.Enabled = True
     dg_datos.Enabled = True
     'dtc_desc1.BackColor = &HFFFFC0
     VAR_SW = ""
'     dtc_codigo9.Enabled = True
'     BtnAprobar.Visible = True
'     BtnModificar.Visible = True
  End If
'  dtc_desc1.Visible = True
'  lbl_aux1.Visible = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub valida_campos()
  'A.
  If (Txt_campo21.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (Txt_aux21.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo31.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (Txt_campo31.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo41.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo51.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_campo41.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_campo7.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  'B.
  If dtc_codigo011.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_campo14.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo61.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo71.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo81.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo10.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo91.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo01.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_campo12.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
End Sub


Private Sub BtnImprimir_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR01.ReportFileName = App.Path & "\Reportes\comercial\R221_ar_calculo_trafico.rpt"
    CR01.WindowShowPrintSetupBtn = True
    CR01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
      'CR01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
      'CR01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
    'Call CREAVISTAF11          'JQA JUN-2008
    CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
    CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
    CR01.StoredProcParam(3) = Me.Ado_datos.Recordset!edif_codigo
    iResult = CR01.PrintReport
    If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnModificar_Click()
'MODIFICAR WWWWWWWWWWWWWWWWWWWWWWWWW
  On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    Fra_datos.Enabled = True
    Frame2.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    FraNavega.Enabled = False
    dg_datos.Enabled = False
    VAR_SW = "MOD"
    'DTPfecha1.Value = Date
    Txt_campo21.SetFocus
'    BtnVer.Visible = True
    'dtc_codigo9.Enabled = False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub BtnVer_Click()
    GlUnidad = Ado_datos.Recordset!unidad_codigo
    GlSolicitud = Ado_datos.Recordset!solicitud_codigo
    GlEdificio = Ado_datos.Recordset!edif_codigo
    'AV_EDIF_VS_BIENES_VENTANUEVA       '
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    'rs_datos10.Open "Select * from AV_EDIF_VS_BIENES_VENTANUEVA where edif_codigo = '" & GlEdificio & "' ", db, adOpenStatic
    rs_datos10.Open "Select * from av_bienes_eqp_caracteristicas_y_venta where edif_codigo = '" & GlEdificio & "' ", db, adOpenStatic
    Set Ado_detalle2.Recordset = rs_datos10
    Set DtGLista.DataSource = Ado_detalle2.Recordset
        
    FrmDetalle.Visible = True
    
'  If dtc_aux41.Text <> "" Then
'    'ARREGLO 1
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc11 = dtc_aux41.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc21 = dtc_aux51.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc31 = IIf(IsNull(Ado_datos.Recordset!c_time_entrada_salida), 0, Ado_datos.Recordset!c_time_entrada_salida)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campod11 = IIf(IsNull(Ado_datos.Recordset!d_num_paradas_probables), 0, Ado_datos.Recordset!d_num_paradas_probables)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe11 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_recorrido), 0, Ado_datos.Recordset!e_tiempo_recorrido)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe21 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_asc_desaceleracion), 0, Ado_datos.Recordset!e_tiempo_asc_desaceleracion)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe31 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_apertura_cierre), 0, Ado_datos.Recordset!e_tiempo_apertura_cierre)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe41 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_entrada_salida), 0, Ado_datos.Recordset!e_tiempo_entrada_salida)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof11 = IIf(IsNull(Ado_datos.Recordset!f_tiempo_recorrido), 0, Ado_datos.Recordset!f_tiempo_recorrido)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof21 = IIf(IsNull(Ado_datos.Recordset!f_time_asc_desaceleracion), 0, Ado_datos.Recordset!f_time_asc_desaceleracion)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof31 = IIf(IsNull(Ado_datos.Recordset!f_time_apertura_cierre), 0, Ado_datos.Recordset!f_time_apertura_cierre)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof41 = IIf(IsNull(Ado_datos.Recordset!f_time_entrada_salida), 0, Ado_datos.Recordset!f_time_entrada_salida)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog11 = IIf(IsNull(Ado_datos.Recordset!g_capacidad_tiempo_cti), 0, Ado_datos.Recordset!g_capacidad_tiempo_cti)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog21 = IIf(IsNull(Ado_datos.Recordset!g_capacidad_total_arreglo), 0, Ado_datos.Recordset!g_capacidad_total_arreglo)
'
'    'ARREGLO 2
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc12 = dtc_aux42.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc22 = dtc_aux52.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc32 = IIf(IsNull(Ado_datos.Recordset!c_time_entrada_salida2), 0, Ado_datos.Recordset!c_time_entrada_salida2)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campod12 = IIf(IsNull(Ado_datos.Recordset!d_num_paradas_probables2), 0, Ado_datos.Recordset!d_num_paradas_probables2)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe12 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_recorrido2), 0, Ado_datos.Recordset!e_tiempo_recorrido2)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe22 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_asc_desaceleracion2), 0, Ado_datos.Recordset!e_tiempo_asc_desaceleracion2)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe32 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_apertura_cierre2), 0, Ado_datos.Recordset!e_tiempo_apertura_cierre2)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe42 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_entrada_salida2), 0, Ado_datos.Recordset!e_tiempo_entrada_salida2)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof12 = IIf(IsNull(Ado_datos.Recordset!f_tiempo_recorrido2), 0, Ado_datos.Recordset!f_tiempo_recorrido2)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof22 = IIf(IsNull(Ado_datos.Recordset!f_time_asc_desaceleracion2), 0, Ado_datos.Recordset!f_time_asc_desaceleracion2)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof32 = IIf(IsNull(Ado_datos.Recordset!f_time_apertura_cierre2), 0, Ado_datos.Recordset!f_time_apertura_cierre2)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof42 = IIf(IsNull(Ado_datos.Recordset!f_time_entrada_salida2), 0, Ado_datos.Recordset!f_time_entrada_salida2)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog12 = IIf(IsNull(Ado_datos.Recordset!g_capacidad_tiempo_cti2), 0, Ado_datos.Recordset!g_capacidad_tiempo_cti2)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog22 = IIf(IsNull(Ado_datos.Recordset!g_capacidad_total_arreglo2), 0, Ado_datos.Recordset!g_capacidad_total_arreglo2)
'
'    'ARREGLO 3
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc13 = dtc_aux43.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc23 = dtc_aux53.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc33 = IIf(IsNull(Ado_datos.Recordset!c_time_entrada_salida3), 0, Ado_datos.Recordset!c_time_entrada_salida3)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campod13 = IIf(IsNull(Ado_datos.Recordset!d_num_paradas_probables3), 0, Ado_datos.Recordset!d_num_paradas_probables3)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe13 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_recorrido3), 0, Ado_datos.Recordset!e_tiempo_recorrido3)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe23 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_asc_desaceleracion3), 0, Ado_datos.Recordset!e_tiempo_asc_desaceleracion3)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe33 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_apertura_cierre3), 0, Ado_datos.Recordset!e_tiempo_apertura_cierre3)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe43 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_entrada_salida3), 0, Ado_datos.Recordset!e_tiempo_entrada_salida3)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof13 = IIf(IsNull(Ado_datos.Recordset!f_tiempo_recorrido3), 0, Ado_datos.Recordset!f_tiempo_recorrido3)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof23 = IIf(IsNull(Ado_datos.Recordset!f_time_asc_desaceleracion3), 0, Ado_datos.Recordset!f_time_asc_desaceleracion3)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof33 = IIf(IsNull(Ado_datos.Recordset!f_time_apertura_cierre3), 0, Ado_datos.Recordset!f_time_apertura_cierre3)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof43 = IIf(IsNull(Ado_datos.Recordset!f_time_entrada_salida3), 0, Ado_datos.Recordset!f_time_entrada_salida3)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog13 = IIf(IsNull(Ado_datos.Recordset!g_capacidad_tiempo_cti3), 0, Ado_datos.Recordset!g_capacidad_tiempo_cti3)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog23 = IIf(IsNull(Ado_datos.Recordset!g_capacidad_total_arreglo3), 0, Ado_datos.Recordset!g_capacidad_total_arreglo3)
'
'    'ARREGLO 4
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc14 = dtc_aux44.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc24 = dtc_aux54.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc34 = IIf(IsNull(Ado_datos.Recordset!c_time_entrada_salida4), 0, Ado_datos.Recordset!c_time_entrada_salida4)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campod14 = IIf(IsNull(Ado_datos.Recordset!d_num_paradas_probables4), 0, Ado_datos.Recordset!d_num_paradas_probables4)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe14 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_recorrido4), 0, Ado_datos.Recordset!e_tiempo_recorrido4)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe24 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_asc_desaceleracion4), 0, Ado_datos.Recordset!e_tiempo_asc_desaceleracion4)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe34 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_apertura_cierre4), 0, Ado_datos.Recordset!e_tiempo_apertura_cierre4)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe44 = IIf(IsNull(Ado_datos.Recordset!e_tiempo_entrada_salida4), 0, Ado_datos.Recordset!e_tiempo_entrada_salida4)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof14 = IIf(IsNull(Ado_datos.Recordset!f_tiempo_recorrido4), 0, Ado_datos.Recordset!f_tiempo_recorrido4)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof24 = IIf(IsNull(Ado_datos.Recordset!f_time_asc_desaceleracion4), 0, Ado_datos.Recordset!f_time_asc_desaceleracion4)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof34 = IIf(IsNull(Ado_datos.Recordset!f_time_apertura_cierre4), 0, Ado_datos.Recordset!f_time_apertura_cierre4)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof44 = IIf(IsNull(Ado_datos.Recordset!f_time_entrada_salida4), 0, Ado_datos.Recordset!f_time_entrada_salida4)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog14 = IIf(IsNull(Ado_datos.Recordset!g_capacidad_tiempo_cti4), 0, Ado_datos.Recordset!g_capacidad_tiempo_cti4)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog24 = IIf(IsNull(Ado_datos.Recordset!g_capacidad_total_arreglo4), 0, Ado_datos.Recordset!g_capacidad_total_arreglo4)
'
'    aw_p_ao_solicitud_calculo_trafico_det.Show vbModal
'    'parametro = txt_campo1.Text
'  Else
'    MsgBox "No existen dados calculados, verifique el registro de ... " + Label19.Caption
'  End If
'  Call ABRIR_TABLAS_AUX
''    If Ado_datos.Recordset!estado_codigo = "REG" Then
''        Call OptFilGral1_Click
''    Else
''        Call OptFilGral2_Click
''    End If
End Sub

Private Sub dtc_aux31_Click(Area As Integer)
    dtc_codigo31.BoundText = dtc_aux31.BoundText
    dtc_desc31.BoundText = dtc_aux31.BoundText
End Sub

Private Sub dtc_aux32_Click(Area As Integer)
    dtc_codigo32.BoundText = dtc_aux32.BoundText
    dtc_desc32.BoundText = dtc_aux32.BoundText
End Sub

Private Sub dtc_aux33_Click(Area As Integer)
    dtc_codigo33.BoundText = dtc_aux33.BoundText
    dtc_desc33.BoundText = dtc_aux33.BoundText
End Sub

Private Sub dtc_aux34_Click(Area As Integer)
    dtc_codigo34.BoundText = dtc_aux34.BoundText
    dtc_desc34.BoundText = dtc_aux34.BoundText
End Sub

Private Sub dtc_aux41_Click(Area As Integer)
    dtc_desc41.BoundText = dtc_aux41.BoundText
    dtc_codigo41.BoundText = dtc_aux41.BoundText
    dtc_valor41.BoundText = dtc_aux41.BoundText
End Sub

Private Sub dtc_aux42_Click(Area As Integer)
    dtc_desc42.BoundText = dtc_aux42.BoundText
    dtc_codigo42.BoundText = dtc_aux42.BoundText
End Sub

Private Sub dtc_aux43_Click(Area As Integer)
    dtc_desc43.BoundText = dtc_aux43.BoundText
    dtc_codigo43.BoundText = dtc_aux43.BoundText
End Sub

Private Sub dtc_aux44_Click(Area As Integer)
    dtc_desc44.BoundText = dtc_aux44.BoundText
    dtc_codigo44.BoundText = dtc_aux44.BoundText
End Sub

Private Sub dtc_aux51_Click(Area As Integer)
    dtc_desc51.BoundText = dtc_aux51.BoundText
    dtc_codigo51.BoundText = dtc_aux51.BoundText
End Sub

Private Sub dtc_aux52_Click(Area As Integer)
    dtc_desc52.BoundText = dtc_aux52.BoundText
    dtc_codigo52.BoundText = dtc_aux52.BoundText
End Sub

Private Sub dtc_aux53_Click(Area As Integer)
    dtc_desc53.BoundText = dtc_aux53.BoundText
    dtc_codigo53.BoundText = dtc_aux53.BoundText
End Sub

Private Sub dtc_aux54_Click(Area As Integer)
    dtc_desc54.BoundText = dtc_aux54.BoundText
    dtc_codigo54.BoundText = dtc_aux54.BoundText
End Sub

Private Sub dtc_codigo04_Click(Area As Integer)
    dtc_desc04.BoundText = dtc_codigo04.BoundText
End Sub

Private Sub dtc_codigo11_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub dtc_codigo12_Click(Area As Integer)
    dtc_desc12.BoundText = dtc_codigo12.BoundText
End Sub

Private Sub dtc_codigo13_Click(Area As Integer)
    dtc_desc13.BoundText = dtc_codigo13.BoundText
End Sub

Private Sub dtc_codigo14_Click(Area As Integer)
    dtc_desc14.BoundText = dtc_codigo14.BoundText
End Sub

'Private Sub dtc_codigo24_Click(Area As Integer)
'    dtc_desc24.BoundText = dtc_codigo24.BoundText
'End Sub

Private Sub dtc_codigo34_Click(Area As Integer)
    dtc_desc34.BoundText = dtc_codigo34.BoundText
    dtc_aux34.BoundText = dtc_codigo34.BoundText
End Sub

Private Sub dtc_codigo44_Click(Area As Integer)
    dtc_desc44.BoundText = dtc_codigo44.BoundText
    dtc_aux44.BoundText = dtc_codigo44.BoundText
End Sub

Private Sub dtc_codigo54_Click(Area As Integer)
    dtc_desc54.BoundText = dtc_codigo54.BoundText
    dtc_aux54.BoundText = dtc_codigo54.BoundText
End Sub

Private Sub dtc_codigo64_Click(Area As Integer)
    dtc_desc64.BoundText = dtc_codigo64.BoundText
End Sub

Private Sub dtc_codigo74_Click(Area As Integer)
    dtc_desc74.BoundText = dtc_codigo74.BoundText
End Sub

Private Sub dtc_codigo84_Click(Area As Integer)
    dtc_desc84.BoundText = dtc_codigo84.BoundText
End Sub

Private Sub dtc_codigo94_Click(Area As Integer)
    dtc_desc94.BoundText = dtc_codigo94.BoundText
End Sub

Private Sub dtc_desc04_Click(Area As Integer)
    dtc_codigo04.BoundText = dtc_codigo04.BoundText
    dtc_desc04.ToolTipText = dtc_desc04.Text
End Sub

Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
    dtc_desc11.ToolTipText = dtc_desc11.Text
End Sub

Private Sub dtc_desc12_Click(Area As Integer)
    dtc_codigo12.BoundText = dtc_desc12.BoundText
    dtc_desc12.ToolTipText = dtc_desc12.Text
End Sub

Private Sub dtc_desc13_Click(Area As Integer)
    dtc_codigo13.BoundText = dtc_desc13.BoundText
    dtc_desc13.ToolTipText = dtc_desc13.Text
End Sub

Private Sub dtc_desc14_Click(Area As Integer)
    dtc_codigo14.BoundText = dtc_desc14.BoundText
    dtc_desc14.ToolTipText = dtc_desc14.Text
End Sub

'Private Sub dtc_desc24_Click(Area As Integer)
'    dtc_codigo24.BoundText = dtc_desc24.BoundText
'End Sub

Private Sub dtc_desc31_LostFocus()
'    If CDbl(dtc_desc31.Text) = 0 Or dtc_desc31.Text = "" Then
'        MsgBox "Vuelva a registrar el: " + lbl_campo3.Caption
'        Txt_campo31 = Round(CDbl(lbl_campoh42.Caption) / 1, 0)
'    Else
'        Txt_campo31 = Round(CDbl(lbl_campoh42.Caption) / CDbl(dtc_desc31.Text), 0)
'    End If
    
'    If CDbl(Txt_campo31.Text) > 0 Then
'        '=+SI(B40>0;B9-(B9-1)*((B9-2)/(B9-1))^B39;0)
'        var_campod11 = CDbl(Txt_campo21.Text) - (CDbl(Txt_campo21.Text) - 1) * ((CDbl(Txt_campo21.Text) - 2) / (CDbl(Txt_campo21.Text) - 1)) ^ CDbl(dtc_desc31.Text)
'    Else
'        var_campod11 = 0
'    End If
End Sub

Private Sub dtc_desc32_LostFocus()
'    If CDbl(dtc_desc32.Text) = 0 Or dtc_desc32.Text = "" Then
'        MsgBox "Vuelva a registrar el: " + lbl_campo3.Caption
'        Txt_campo32 = Round(CDbl(lbl_campoh42.Caption) / 1, 0)
'    Else
'        Txt_campo32 = Round(CDbl(lbl_campoh42.Caption) / CDbl(dtc_desc32.Text), 0)
'    End If
'    If CDbl(Txt_campo32.Text) > 0 Then
'        '=+SI(B40>0;B9-(B9-1)*((B9-2)/(B9-1))^B39;0)
'        var_campod12 = Round(CDbl(Txt_campo22.Text) - (CDbl(Txt_campo22.Text) - 1) * ((CDbl(Txt_campo22.Text) - 2) / (CDbl(Txt_campo22.Text) - 1)) ^ CDbl(dtc_desc32.Text), 2)
'    Else
'        var_campod12 = 0
'    End If
End Sub

Private Sub dtc_desc33_LostFocus()
'    If CDbl(dtc_desc33.Text) = 0 Or dtc_desc33.Text = "" Then
'        MsgBox "Vuelva a registrar el: " + lbl_campo3.Caption
'        Txt_campo33 = Round(CDbl(lbl_campoh42.Caption) / 1, 0)
'    Else
'        Txt_campo33 = Round(CDbl(lbl_campoh42.Caption) / CDbl(dtc_desc33.Text), 0)
'    End If
'    If CDbl(Txt_campo33.Text) > 0 Then
'        '=+SI(B40>0;B9-(B9-1)*((B9-2)/(B9-1))^B39;0)
'        var_campod13 = CDbl(Txt_campo23.Text) - (CDbl(Txt_campo23.Text) - 1) * ((CDbl(Txt_campo23.Text) - 2) / (CDbl(Txt_campo23.Text) - 1)) ^ CDbl(dtc_desc33.Text)
'    Else
'        var_campod13 = 0
'    End If
End Sub

Private Sub dtc_desc34_Click(Area As Integer)
    dtc_codigo34.BoundText = dtc_desc34.BoundText
    dtc_aux34.BoundText = dtc_desc34.BoundText
End Sub

Private Sub dtc_desc34_LostFocus()
'    If CDbl(Txt_campo34.Text) > 0 Then
'        '=+SI(B40>0;B9-(B9-1)*((B9-2)/(B9-1))^B39;0)
'        var_campod14 = CDbl(Txt_campo24.Text) - (CDbl(Txt_campo24.Text) - 1) * ((CDbl(Txt_campo24.Text) - 2) / (CDbl(Txt_campo24.Text) - 1)) ^ CDbl(dtc_desc34.Text)
'    Else
'        var_campod14 = 0
'    End If
End Sub

Private Sub dtc_desc41_LostFocus()
    '=+B10*2/B41
    var_campoe11 = CDbl(Txt_aux21.Text) * 2 / CDbl(dtc_valor41.Text)
End Sub

Private Sub dtc_desc42_LostFocus()
    var_campoe12 = CDbl(Txt_aux22.Text) * 2 / CDbl(dtc_desc42.Text)
End Sub

Private Sub dtc_desc43_LostFocus()
    var_campoe13 = CDbl(Txt_aux23.Text) * 2 / CDbl(dtc_desc43.Text)
End Sub

Private Sub dtc_desc44_Click(Area As Integer)
    dtc_codigo44.BoundText = dtc_codigo44.BoundText
    dtc_aux44.BoundText = dtc_desc44.BoundText
End Sub

Private Sub dtc_desc54_Click(Area As Integer)
    dtc_codigo54.BoundText = dtc_codigo54.BoundText
    dtc_aux54.BoundText = dtc_desc54.BoundText
    dtc_desc54.ToolTipText = dtc_desc54.Text
End Sub

Private Sub dtc_desc61_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl_ttt.Left = dtc_desc61.Left
    lbl_ttt.Top = dtc_desc61.Top
    lbl_ttt.Caption = dtc_desc61.Text
End Sub

Private Sub dtc_desc64_Click(Area As Integer)
    dtc_codigo64.BoundText = dtc_codigo64.BoundText
    dtc_desc64.ToolTipText = dtc_desc64.Text
End Sub

Private Sub dtc_desc74_Click(Area As Integer)
    dtc_codigo74.BoundText = dtc_codigo74.BoundText
    dtc_desc74.ToolTipText = dtc_desc74.Text
End Sub

Private Sub dtc_desc84_Click(Area As Integer)
    dtc_codigo84.BoundText = dtc_codigo84.BoundText
    dtc_desc84.ToolTipText = dtc_desc84.Text
End Sub

Private Sub dtc_desc94_Click(Area As Integer)
    dtc_codigo94.BoundText = dtc_codigo94.BoundText
    dtc_desc93.ToolTipText = dtc_desc93.Text
End Sub

Private Sub dtc_valor41_Click(Area As Integer)
    dtc_codigo41.BoundText = dtc_valor41.BoundText
    dtc_aux41.BoundText = dtc_valor41.BoundText
    dtc_desc41.BoundText = dtc_valor41.BoundText
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    'parametro = txt_campo1.Text
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        usuario2 = rs_aux3!beneficiario_codigo
        VAR_DA = rs_aux3!da_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.2"
    End If
    VAR_UORIGEN = Aux
    Select Case VAR_DA
        Case "1.8"    'Cochabamba
            Aux = "DCOMB"
            VAR_DPTO = "3"
        Case "1.7"    'Santa Cruz
            Aux = "DCOMS"
            VAR_DPTO = "7"
        Case "1.2"    'La Paz - Comercial
            Aux = "DVTA"
            VAR_DPTO = "2"
        Case "1.9"    ' Chuquisaca
            Aux = "DCOMC"
            VAR_DPTO = "1"
        Case "1.3"    'La Paz - Modernizacion
            Aux = "DNMOD"
            VAR_DPTO = "2"
        Case "0"    ' TODO
            Aux = "DVTA"
            VAR_DPTO = "0"
     End Select
    
    parametro = Aux
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    
    Fra_datos.Enabled = False
    Frame2.Enabled = False
    dg_datos.Enabled = True
    If Ado_datos.Recordset.RecordCount > 0 Then
        If Ado_datos.Recordset!estado_codigo = "REG" Then
            BtnModificar.Visible = True
            BtnAprobar.Visible = True
         Else
            BtnModificar.Visible = False
            BtnAprobar.Visible = False
         End If
'     If Ado_datos.Recordset!estado_codigo = "APR" Then
'        BtnBuscar.Visible = False
'        BtnGrabar.Visible = False
'        BtnAprobar.Visible = False
'     Else
''        Call BtnModificar_Click
'        BtnGrabar.Visible = True
'        If Ado_datos.Recordset!estado_codigo_verif = "APR" Then
'            'BtnBuscar.Visible = False
'            BtnAprobar.Visible = False
'        Else
'            'BtnBuscar.Visible = True
'            BtnAprobar.Visible = True
'        End If
'     End If
   End If
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
        
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from gc_edificaciones order by edif_denominacion", db, adOpenStatic
    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
    'A.
    Set rs_datos21 = New ADODB.Recordset
    If rs_datos21.State = 1 Then rs_datos21.Close
    rs_datos21.Open "Select * from ac_bienes_equipo_recorrido ", db, adOpenStatic
    Set Ado_datos21.Recordset = rs_datos21
    'dtc_desc21.BoundText = dtc_codigo21.BoundText
    
    Set rs_datos22 = New ADODB.Recordset
    If rs_datos22.State = 1 Then rs_datos22.Close
    rs_datos22.Open "Select * from ac_bienes_equipo_recorrido ", db, adOpenStatic
    Set Ado_datos22.Recordset = rs_datos22
    'dtc_desc22.BoundText = dtc_codigo22.BoundText
    
    Set rs_datos23 = New ADODB.Recordset
    If rs_datos23.State = 1 Then rs_datos23.Close
    rs_datos23.Open "Select * from ac_bienes_equipo_recorrido ", db, adOpenStatic
    Set Ado_datos23.Recordset = rs_datos23
    'dtc_desc23.BoundText = dtc_codigo23.BoundText
    
    Set rs_datos24 = New ADODB.Recordset
    If rs_datos24.State = 1 Then rs_datos24.Close
    rs_datos24.Open "Select * from ac_bienes_equipo_recorrido ", db, adOpenStatic
    Set Ado_datos24.Recordset = rs_datos24
    'dtc_desc24.BoundText = dtc_codigo24.BoundText
    
    Set rs_datos31 = New ADODB.Recordset
    If rs_datos31.State = 1 Then rs_datos31.Close
    rs_datos31.Open "Select * from ac_bienes_equipo_nro_pasajeros ", db, adOpenStatic
    Set Ado_datos31.Recordset = rs_datos31
    dtc_desc31.BoundText = dtc_codigo31.BoundText
    
    Set rs_datos32 = New ADODB.Recordset
    If rs_datos32.State = 1 Then rs_datos32.Close
    rs_datos32.Open "Select * from ac_bienes_equipo_nro_pasajeros ", db, adOpenStatic
    Set Ado_datos32.Recordset = rs_datos32
    dtc_desc32.BoundText = dtc_codigo32.BoundText
    
    Set rs_datos33 = New ADODB.Recordset
    If rs_datos33.State = 1 Then rs_datos33.Close
    rs_datos33.Open "Select * from ac_bienes_equipo_nro_pasajeros ", db, adOpenStatic
    Set Ado_datos33.Recordset = rs_datos33
    dtc_desc33.BoundText = dtc_codigo33.BoundText
    
    Set rs_datos34 = New ADODB.Recordset
    If rs_datos34.State = 1 Then rs_datos34.Close
    rs_datos34.Open "Select * from ac_bienes_equipo_nro_pasajeros ", db, adOpenStatic
    Set Ado_datos34.Recordset = rs_datos34
    dtc_desc34.BoundText = dtc_codigo34.BoundText
        
    Set rs_datos41 = New ADODB.Recordset
    If rs_datos41.State = 1 Then rs_datos41.Close
    rs_datos41.Open "Select * from ac_bienes_equipo_velocidad ", db, adOpenStatic
    Set Ado_datos41.Recordset = rs_datos41
    dtc_desc41.BoundText = dtc_codigo41.BoundText
    
    Set rs_datos42 = New ADODB.Recordset
    If rs_datos42.State = 1 Then rs_datos42.Close
    rs_datos42.Open "Select * from ac_bienes_equipo_velocidad ", db, adOpenStatic
    Set Ado_datos42.Recordset = rs_datos42
    dtc_desc42.BoundText = dtc_codigo42.BoundText
    
    Set rs_datos43 = New ADODB.Recordset
    If rs_datos43.State = 1 Then rs_datos43.Close
    rs_datos43.Open "Select * from ac_bienes_equipo_velocidad ", db, adOpenStatic
    Set Ado_datos43.Recordset = rs_datos43
    dtc_desc43.BoundText = dtc_codigo43.BoundText
        
    Set rs_datos44 = New ADODB.Recordset
    If rs_datos44.State = 1 Then rs_datos44.Close
    rs_datos44.Open "Select * from ac_bienes_equipo_velocidad ", db, adOpenStatic
    Set Ado_datos44.Recordset = rs_datos44
    dtc_desc44.BoundText = dtc_codigo44.BoundText
    
    Set rs_datos51 = New ADODB.Recordset
    If rs_datos51.State = 1 Then rs_datos51.Close
    rs_datos51.Open "Select * from ac_bienes_equipo_tipo_puerta_piso ", db, adOpenStatic
    Set Ado_datos51.Recordset = rs_datos51
    dtc_desc51.BoundText = dtc_codigo51.BoundText
    
    Set rs_datos52 = New ADODB.Recordset
    If rs_datos52.State = 1 Then rs_datos52.Close
    rs_datos52.Open "Select * from ac_bienes_equipo_tipo_puerta_piso ", db, adOpenStatic
    Set Ado_datos52.Recordset = rs_datos52
    dtc_desc52.BoundText = dtc_codigo52.BoundText
    
    Set rs_datos53 = New ADODB.Recordset
    If rs_datos53.State = 1 Then rs_datos53.Close
    rs_datos53.Open "Select * from ac_bienes_equipo_tipo_puerta_piso ", db, adOpenStatic
    Set Ado_datos53.Recordset = rs_datos53
    dtc_desc53.BoundText = dtc_codigo53.BoundText

    Set rs_datos54 = New ADODB.Recordset
    If rs_datos54.State = 1 Then rs_datos54.Close
    rs_datos54.Open "Select * from ac_bienes_equipo_tipo_puerta_piso ", db, adOpenStatic
    Set Ado_datos54.Recordset = rs_datos54
    dtc_desc54.BoundText = dtc_codigo54.BoundText
    
    'B.
    Set rs_datos61 = New ADODB.Recordset
    If rs_datos61.State = 1 Then rs_datos61.Close
    rs_datos61.Open "Select * from ac_bienes_equipo_cabina_estetica ", db, adOpenStatic
    'rs_datos4.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos61.Recordset = rs_datos61
    dtc_desc61.BoundText = dtc_codigo61.BoundText
    
    Set rs_datos62 = New ADODB.Recordset
    If rs_datos62.State = 1 Then rs_datos62.Close
    rs_datos62.Open "Select * from ac_bienes_equipo_cabina_estetica ", db, adOpenStatic
    'rs_datos5.Open "gp_listar_apr_gc_proceso_nivel1", db, adOpenStatic
    Set Ado_datos62.Recordset = rs_datos62
    dtc_desc62.BoundText = dtc_codigo62.BoundText
    
    Set rs_datos63 = New ADODB.Recordset
    If rs_datos63.State = 1 Then rs_datos63.Close
    rs_datos63.Open "Select * from ac_bienes_equipo_cabina_estetica ", db, adOpenStatic
    'rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
    Set Ado_datos63.Recordset = rs_datos63
    dtc_desc63.BoundText = dtc_codigo63.BoundText
    
    Set rs_datos64 = New ADODB.Recordset
    If rs_datos64.State = 1 Then rs_datos63.Close
    rs_datos64.Open "Select * from ac_bienes_equipo_cabina_estetica ", db, adOpenStatic
    'rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
    Set Ado_datos64.Recordset = rs_datos64
    dtc_desc64.BoundText = dtc_codigo64.BoundText
    
    Set rs_datos71 = New ADODB.Recordset
    If rs_datos71.State = 1 Then rs_datos71.Close
    rs_datos71.Open "Select * from ac_bienes_equipo_tecnologia ", db, adOpenStatic
    'rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
    Set Ado_datos71.Recordset = rs_datos71
    dtc_desc71.BoundText = dtc_codigo71.BoundText
          
    Set rs_datos72 = New ADODB.Recordset
    If rs_datos72.State = 1 Then rs_datos72.Close
    rs_datos72.Open "Select * from ac_bienes_equipo_tecnologia ", db, adOpenStatic
    'rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
    Set Ado_datos72.Recordset = rs_datos72
    dtc_desc72.BoundText = dtc_codigo72.BoundText
    
    Set rs_datos73 = New ADODB.Recordset
    If rs_datos73.State = 1 Then rs_datos73.Close
    rs_datos73.Open "Select * from ac_bienes_equipo_tecnologia ", db, adOpenStatic
    'rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
    Set Ado_datos73.Recordset = rs_datos73
    dtc_desc73.BoundText = dtc_codigo73.BoundText
        
    Set rs_datos74 = New ADODB.Recordset
    If rs_datos74.State = 1 Then rs_datos74.Close
    rs_datos74.Open "Select * from ac_bienes_equipo_tecnologia ", db, adOpenStatic
    'rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
    Set Ado_datos74.Recordset = rs_datos74
    dtc_desc74.BoundText = dtc_codigo74.BoundText
    
    Set rs_datos81 = New ADODB.Recordset
    If rs_datos81.State = 1 Then rs_datos81.Close
    rs_datos81.Open "Select * from ac_bienes_equipo_sistema_puertas ", db, adOpenStatic
    'rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
    Set Ado_datos81.Recordset = rs_datos81
    dtc_desc81.BoundText = dtc_codigo81.BoundText
          
    Set rs_datos82 = New ADODB.Recordset
    If rs_datos82.State = 1 Then rs_datos82.Close
    rs_datos82.Open "Select * from ac_bienes_equipo_sistema_puertas ", db, adOpenStatic
    'rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
    Set Ado_datos82.Recordset = rs_datos82
    dtc_desc82.BoundText = dtc_codigo82.BoundText
    
    Set rs_datos83 = New ADODB.Recordset
    If rs_datos83.State = 1 Then rs_datos83.Close
    rs_datos83.Open "Select * from ac_bienes_equipo_sistema_puertas ", db, adOpenStatic
    'rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
    Set Ado_datos83.Recordset = rs_datos83
    dtc_desc83.BoundText = dtc_codigo83.BoundText
        
    Set rs_datos84 = New ADODB.Recordset
    If rs_datos84.State = 1 Then rs_datos84.Close
    rs_datos84.Open "Select * from ac_bienes_equipo_sistema_puertas ", db, adOpenStatic
    'rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
    Set Ado_datos84.Recordset = rs_datos84
    dtc_desc84.BoundText = dtc_codigo84.BoundText
    
    Set rs_datos91 = New ADODB.Recordset
    If rs_datos91.State = 1 Then rs_datos91.Close
    rs_datos91.Open "Select * from ac_bienes_equipo_condicion_ventas ", db, adOpenStatic
    'rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
    Set Ado_datos91.Recordset = rs_datos91
    dtc_desc91.BoundText = dtc_codigo91.BoundText
          
    Set rs_datos92 = New ADODB.Recordset
    If rs_datos92.State = 1 Then rs_datos92.Close
    rs_datos92.Open "Select * from ac_bienes_equipo_condicion_ventas ", db, adOpenStatic
    'rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
    Set Ado_datos92.Recordset = rs_datos92
    dtc_desc92.BoundText = dtc_codigo92.BoundText
    
    Set rs_datos93 = New ADODB.Recordset
    If rs_datos93.State = 1 Then rs_datos93.Close
    rs_datos93.Open "Select * from ac_bienes_equipo_condicion_ventas ", db, adOpenStatic
    'rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
    Set Ado_datos93.Recordset = rs_datos93
    dtc_desc93.BoundText = dtc_codigo93.BoundText
           
    Set rs_datos94 = New ADODB.Recordset
    If rs_datos94.State = 1 Then rs_datos94.Close
    rs_datos94.Open "Select * from ac_bienes_equipo_condicion_ventas ", db, adOpenStatic
    'rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
    Set Ado_datos94.Recordset = rs_datos94
    dtc_desc94.BoundText = dtc_codigo94.BoundText
    
    Set rs_datos01 = New ADODB.Recordset
    If rs_datos01.State = 1 Then rs_datos01.Close
    rs_datos01.Open "Select * from ac_bienes_equipo_condicion_cabina ", db, adOpenStatic
    'rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
    Set Ado_datos01.Recordset = rs_datos01
    dtc_desc01.BoundText = dtc_codigo01.BoundText
          
    Set rs_datos02 = New ADODB.Recordset
    If rs_datos02.State = 1 Then rs_datos02.Close
    rs_datos02.Open "Select * from ac_bienes_equipo_condicion_cabina ", db, adOpenStatic
    'rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
    Set Ado_datos02.Recordset = rs_datos02
    dtc_desc02.BoundText = dtc_codigo02.BoundText
    
    Set rs_datos03 = New ADODB.Recordset
    If rs_datos03.State = 1 Then rs_datos03.Close
    rs_datos03.Open "Select * from ac_bienes_equipo_condicion_cabina ", db, adOpenStatic
    'rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
    Set Ado_datos03.Recordset = rs_datos03
    dtc_desc03.BoundText = dtc_codigo03.BoundText
    
    Set rs_datos04 = New ADODB.Recordset
    If rs_datos04.State = 1 Then rs_datos04.Close
    rs_datos04.Open "Select * from ac_bienes_equipo_condicion_cabina ", db, adOpenStatic
    'rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
    Set Ado_datos04.Recordset = rs_datos04
    dtc_desc04.BoundText = dtc_codigo04.BoundText
    
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    rs_datos11.Open "Select * from ac_bienes_equipo_ctrl_maquina ", db, adOpenStatic
    'rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
          
    Set rs_datos12 = New ADODB.Recordset
    If rs_datos12.State = 1 Then rs_datos12.Close
    rs_datos12.Open "Select * from ac_bienes_equipo_ctrl_maquina ", db, adOpenStatic
    'rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
    Set Ado_datos12.Recordset = rs_datos12
    dtc_desc12.BoundText = dtc_codigo12.BoundText
    
    Set rs_datos13 = New ADODB.Recordset
    If rs_datos13.State = 1 Then rs_datos13.Close
    rs_datos13.Open "Select * from ac_bienes_equipo_ctrl_maquina ", db, adOpenStatic
    'rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
    Set Ado_datos13.Recordset = rs_datos13
    dtc_desc13.BoundText = dtc_codigo13.BoundText
    
    Set rs_datos14 = New ADODB.Recordset
    If rs_datos14.State = 1 Then rs_datos14.Close
    rs_datos14.Open "Select * from ac_bienes_equipo_ctrl_maquina ", db, adOpenStatic
    'rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
    Set Ado_datos14.Recordset = rs_datos14
    dtc_desc14.BoundText = dtc_codigo14.BoundText
    
End Sub

Private Sub Maximo_Numerador()
'  TxtCrr.Text = "1"
'  Set RsTmp = New ADODB.Recordset
''  Set rst_ben = New ADODB.Recordset
''  rst_ben.Open "Select max(trafico_codigo) + 1 as Codigo from ao_solicitud_ctrl_trafico ", DB, adOpenStatic
''  Set AdoTip_ben.Recordset = rst_ben
'  RsTmp.Open "Select max(trafico_codigo) + 1 as Codigo from ao_solicitud_ctrl_trafico ", db, adOpenStatic
'  'Set RsTmp = DbConex.Execute("Select max(trafico_codigo) + 1 as Codigo from ao_solicitud_ctrl_trafico ;")
'  If Not RsTmp.EOF Then
'     TxtCrr.Text = RsTmp!Codigo
'  End If
End Sub

Private Sub Carga_Beneficiario()
'  Set rstbeneficiario = New ADODB.Recordset
'  If rstbeneficiario.State = 1 Then rstbeneficiario.Close
'  sql = "SELECT ges_gestion as gestion,unidad_codigo as Unid_Ejec,solicitud_codigo as Codigo,trafico_codigo,estado_codigo,edif_codigo,trafico_num_paradas,trafico_recorrido," _
'  & " trafico_nro_equipos,vel_equipo_codigo,tipo_puerta,trafico_ancho_puerta,cabina_codigo," _
'  & " tecnologia_codigo , sist_puerta, condicion_ventas " _
'  & " From ao_solicitud_ctrl_trafico WHERE estado_codigo = 'REG'"
''  SQL = "Select ges_gestion,unidad_codigo,solicitud_codigo,trafico_codigo from ao_solicitud_ctrl_trafico order by unidad_codigo,solicitud_codigo,trafico_codigo"
'  rstbeneficiario.Open sql, db, adOpenKeyset, adLockOptimistic, adCmdText
'  Set Ado_datos.Recordset = rstbeneficiario
'  'Ado_datos.ConnectionString = sConex
'  'Ado_datos.RecordSource = SQL
'  'Ado_datos.Refresh
'
'  dg_datos.Columns(0).Width = 800 'maxWidth
'  dg_datos.Columns(1).Width = 1556
'  dg_datos.Columns(2).Width = 1556
'  dg_datos.Columns(4).Alignment = dbgRight
''  dg_datos.Columns(2).Alignment = dbgRight
''  dg_datos.Columns(3).Alignment = dbgRight
''  dg_datos.Columns(4).Alignment = dbgCenter
''  dg_datos.Columns(2).NumberFormat = ("###0.00")
''  dg_datos.Columns(3).NumberFormat = ("###0.00")
'
'  'LblReg.Caption = "Total Registros --> " & Ado_datos.Recordset.RecordCount
End Sub

Function Llena_Combos()
'  CmbReco.Clear
'  sql = " SELECT recorrido_descripcion From ac_bienes_equipo_recorrido; "
'  If RsTmp.State = 1 Then RsTmp.Close
'  RsTmp.Open sql, db, adOpenStatic
'  If Not RsTmp.EOF Then
'     While Not (RsTmp.EOF)
'           CmbReco.AddItem RsTmp!recorrido_descripcion
'         RsTmp.MoveNext
'     Wend
'  End If
''---
'  CmbNroPasaj.Clear
'  sql = " SELECT pasajeros_descripcion From ac_bienes_equipo_nro_pasajeros; "
'  If RsTmp.State = 1 Then RsTmp.Close
'  RsTmp.Open sql, db, adOpenStatic
'  If Not RsTmp.EOF Then
'     While Not (RsTmp.EOF)
'           CmbNroPasaj.AddItem RsTmp!pasajeros_descripcion
'         RsTmp.MoveNext
'     Wend
'  End If
''---
''  CmbVelEq.Clear
''  SQL = " SELECT vel_equipo_descripcion From ac_bienes_equipo_velocidad WHERE vel_equipo_codigo = " & nCod & "; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open SQL, DB, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbVelEq.AddItem RsTmp!pasajeros_descripcion
''         RsTmp.MoveNext
''     Wend
''  End If
End Function

Function Llena_Clientes1()
'  CmbCodCli1.Clear
'  CmbCliente.Clear
'  Call ABRE_CONECCION
'  Set RsTmp = DbConex.Execute("select * from CLIENTES order by nomBRECLI ;")
'  If Not RsTmp.EOF Then
'     While Not (RsTmp.EOF)
'           CmbCodCli1.AddItem RsTmp!CodCli
'           CmbCliente.AddItem RsTmp!nombrecli
'         RsTmp.MoveNext
'     Wend
'  End If
'  Call CERRAR_CONECCION
End Function

Private Sub CmbCliente_Click()
' If CmbCliente.ListIndex = -1 Then Exit Sub
' CmbCodCli1.ListIndex = CmbCliente.ListIndex
End Sub

Private Sub dg_datos_Click()
'  MsgBox "sss"
'   Call Llena_Varios
'  txtDescrip = dg_datos.Columns(1).Text
End Sub

'Private Sub dg_datos_KeyDown(KeyCode As Integer, Shift As Integer)
'  Call Llena_Varios
''  txtDescrip = dg_datos.Columns(1).Text
'End Sub
'Function Llena_Varios()
''  If RsTmp.State = 1 Then RsTmp.Close
''  'If DB.State = adStateOpen Then DB.Close
''  sql = " SELECT unidad_descripcion FROM gc_unidad_ejecutora " & _
''        "  WHERE unidad_codigo = '" & TxtUEjec & "';"
''  RsTmp.Open sql, db, adCmdText 'adOpenStatic
''  If Not RsTmp.EOF Then
''     txtDescrip.Text = RsTmp!unidad_descripcion
''  End If
'''--
''  sql = " SELECT edif_denominacion FROM gc_edificaciones " & _
''              "  WHERE edif_codigo = '" & Txtedif & "';"
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     txtDesEdif.Text = RsTmp!edif_denominacion
''  End If
'''-------
''  CmbVelEq.Clear
''  sql = " SELECT vel_equipo_descripcion From ac_bienes_equipo_velocidad WHERE vel_equipo_codigo = " & TxtCodVel & "; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbVelEq.AddItem RsTmp!vel_equipo_descripcion
''         RsTmp.MoveNext
''     Wend
''     CmbVelEq.ListIndex = 0
''  End If
'''-------
''  CmbTipoPuerta.Clear
''  sql = " SELECT tipo_puerta_descripcion From ac_bienes_equipo_tipo_puerta_piso WHERE tipo_puerta = " & Txttip & "; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbTipoPuerta.AddItem RsTmp!tipo_puerta_descripcion
''         RsTmp.MoveNext
''     Wend
''     CmbTipoPuerta.ListIndex = 0
''  End If
'''-------cabina_codigo
''  CmbEstat.Clear
''  sql = " SELECT cabina_descripcion From ac_bienes_equipo_cabina_estetica WHERE cabina_codigo = '" & TxtEst & "'; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbEstat.AddItem RsTmp!cabina_descripcion
''         RsTmp.MoveNext
''     Wend
''     CmbEstat.ListIndex = 0
''  End If
'''-------
''  CmbTecno.Clear
''  sql = " SELECT tecnologia_descripcion From ac_bienes_equipo_tecnologia WHERE tecnologia_codigo = '" & TxtTecno & "'; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbTecno.AddItem RsTmp!tecnologia_descripcion
''         RsTmp.MoveNext
''     Wend
''     CmbTecno.ListIndex = 0
''  End If
''        'FALTA sist_puerta
'''-------
''  CmbCondVenta.Clear
''  sql = " SELECT condicion_ventas_descripcion From ac_bienes_equipo_condicion_ventas WHERE condicion_ventas = '" & TxtCondVenta & "'; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbCondVenta.AddItem RsTmp!condicion_ventas_descripcion
''         RsTmp.MoveNext
''     Wend
''     CmbCondVenta.ListIndex = 0
''  End If
''
'End Function

'Private Sub dtc_aux1_Click(Area As Integer)
'    dtc_desc1.BoundText = dtc_aux1.BoundText
'    dtc_codigo1.BoundText = dtc_aux1.BoundText
'End Sub
'
Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
'    dtc_aux1.BoundText = dtc_codigo1.BoundText
End Sub

'Private Sub dtc_codigo21_Click(Area As Integer)
'    dtc_desc21.BoundText = dtc_codigo21.BoundText
'End Sub

'Private Sub dtc_codigo22_Click(Area As Integer)
'    dtc_desc22.BoundText = dtc_codigo22.BoundText
'End Sub

'Private Sub dtc_codigo23_Click(Area As Integer)
'    dtc_desc23.BoundText = dtc_codigo23.BoundText
'End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo31_Click(Area As Integer)
    dtc_desc31.BoundText = dtc_codigo31.BoundText
    dtc_aux31.BoundText = dtc_codigo31.BoundText
End Sub

Private Sub dtc_codigo32_Click(Area As Integer)
    dtc_desc32.BoundText = dtc_codigo32.BoundText
    dtc_aux32.BoundText = dtc_codigo32.BoundText
End Sub

Private Sub dtc_codigo33_Click(Area As Integer)
    dtc_desc33.BoundText = dtc_codigo33.BoundText
    dtc_aux33.BoundText = dtc_codigo33.BoundText
End Sub

Private Sub dtc_codigo41_Click(Area As Integer)
    dtc_desc41.BoundText = dtc_codigo41.BoundText
    dtc_aux41.BoundText = dtc_codigo41.BoundText
    dtc_valor41.BoundText = dtc_desc41.BoundText
End Sub

Private Sub dtc_codigo42_Click(Area As Integer)
    dtc_desc42.BoundText = dtc_codigo42.BoundText
    dtc_aux42.BoundText = dtc_codigo42.BoundText
End Sub

Private Sub dtc_codigo43_Click(Area As Integer)
    dtc_desc43.BoundText = dtc_codigo43.BoundText
    dtc_aux43.BoundText = dtc_codigo43.BoundText
End Sub

Private Sub dtc_codigo51_Click(Area As Integer)
    dtc_desc51.BoundText = dtc_codigo51.BoundText
    dtc_aux51.BoundText = dtc_codigo51.BoundText
End Sub

Private Sub dtc_codigo52_Click(Area As Integer)
    dtc_desc52.BoundText = dtc_codigo52.BoundText
    dtc_aux52.BoundText = dtc_codigo52.BoundText
End Sub

Private Sub dtc_codigo53_Click(Area As Integer)
    dtc_desc53.BoundText = dtc_codigo53.BoundText
    dtc_aux53.BoundText = dtc_codigo53.BoundText
End Sub

Private Sub dtc_codigo61_Click(Area As Integer)
    dtc_desc61.BoundText = dtc_codigo61.BoundText
End Sub

Private Sub dtc_codigo62_Click(Area As Integer)
    dtc_desc62.BoundText = dtc_codigo62.BoundText
End Sub

Private Sub dtc_codigo63_Click(Area As Integer)
    dtc_desc63.BoundText = dtc_codigo63.BoundText
End Sub

Private Sub dtc_codigo71_Click(Area As Integer)
    dtc_desc71.BoundText = dtc_codigo71.BoundText
End Sub

Private Sub dtc_codigo72_Click(Area As Integer)
    dtc_desc72.BoundText = dtc_codigo72.BoundText
End Sub

Private Sub dtc_codigo73_Click(Area As Integer)
    dtc_desc73.BoundText = dtc_codigo73.BoundText
End Sub

Private Sub dtc_codigo81_Click(Area As Integer)
    dtc_desc81.BoundText = dtc_codigo81.BoundText
End Sub

Private Sub dtc_codigo82_Click(Area As Integer)
    dtc_desc82.BoundText = dtc_codigo82.BoundText
End Sub

Private Sub dtc_codigo83_Click(Area As Integer)
    dtc_desc83.BoundText = dtc_codigo83.BoundText
End Sub

Private Sub dtc_codigo91_Click(Area As Integer)
    dtc_desc91.BoundText = dtc_codigo91.BoundText
End Sub

Private Sub dtc_codigo92_Click(Area As Integer)
    dtc_desc92.BoundText = dtc_codigo92.BoundText
End Sub

Private Sub dtc_codigo93_Click(Area As Integer)
    dtc_desc93.BoundText = dtc_codigo93.BoundText
End Sub

Private Sub dtc_codigo01_Click(Area As Integer)
    dtc_desc01.BoundText = dtc_codigo01.BoundText
End Sub

Private Sub dtc_codigo02_Click(Area As Integer)
    dtc_desc02.BoundText = dtc_codigo02.BoundText
End Sub

Private Sub dtc_codigo03_Click(Area As Integer)
    dtc_desc03.BoundText = dtc_codigo03.BoundText
End Sub

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

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
'    dtc_aux1.BoundText = dtc_desc1.BoundText
'    Call pnivel1(dtc_codigo1.BoundText)
'    dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub

'Private Sub pnivel1(codigo1 As String)
''   Dim strConsultaF As String
''   strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
'
'   Set dtc_codigo10.RowSource = Nothing
''   Set dtc_codigo10.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_codigo10.ReFill
'   dtc_codigo10.BoundText = Empty
'
'   Set dtc_desc10.RowSource = Nothing
'   'Set dtc_desc10.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_desc10.ReFill
'   dtc_desc10.BoundText = Empty
'End Sub

'Private Sub pnivel11(codigo1 As String)
'   Dim strConsultaF As String
'   'strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
'   strConsultaF = "Select * from gv_personal_contratado where unidad_codigo = '" & codigo1 & "' order by beneficiario_denominacion"
'
'   Set dtc_codigo11.RowSource = Nothing
'   Set dtc_codigo11.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_codigo11.ReFill
'   dtc_codigo11.BoundText = Empty
'
'   Set dtc_desc11.RowSource = Nothing
'   Set dtc_desc11.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_desc11.ReFill
'   dtc_desc11.BoundText = Empty
'End Sub

'Private Sub dtc_desc1_LostFocus()
''    dtc_codigo5.Text = dtc_aux1.Text
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
''    Call pnivel5(dtc_codigo5.BoundText)
''    dtc_desc6.Enabled = True
'End Sub

'Private Sub dtc_desc21_Click(Area As Integer)
'    dtc_codigo21.BoundText = dtc_desc21.BoundText
'End Sub

'Private Sub dtc_desc22_Click(Area As Integer)
'    dtc_codigo22.BoundText = dtc_desc22.BoundText
'End Sub

'Private Sub dtc_desc23_Click(Area As Integer)
'    dtc_codigo23.BoundText = dtc_desc23.BoundText
'End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub

'Private Sub dtc_desc3_LostFocus()
'    dtc_codigo4.Text = dtc_aux3.Text
'    Txt_descripcion.Text = "SOLICITUD DE COTIZACION - " + dtc_desc3.Text
'    dtc_desc4.BoundText = dtc_codigo4.BoundText
'End Sub

Private Sub dtc_desc31_Click(Area As Integer)
    dtc_codigo31.BoundText = dtc_desc31.BoundText
    dtc_aux31.BoundText = dtc_desc31.BoundText
End Sub

Private Sub dtc_desc32_Click(Area As Integer)
    dtc_codigo32.BoundText = dtc_desc32.BoundText
    dtc_aux32.BoundText = dtc_desc32.BoundText
'    Call pnivel5(dtc_codigo5.BoundText)
'    dtc_desc6.Enabled = True
End Sub

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

Private Sub dtc_desc33_Click(Area As Integer)
    dtc_codigo33.BoundText = dtc_desc33.BoundText
    dtc_aux33.BoundText = dtc_desc33.BoundText
'    Call pnivel6(dtc_codigo6.BoundText)
'    dtc_desc7.Enabled = True
End Sub

'Private Sub pnivel6(codigo6 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from gc_proceso_nivel3 where subproceso_codigo = '" & codigo6 & "'"
'
'   Set dtc_codigo7.RowSource = Nothing
'   Set dtc_codigo7.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo7.RowSource = db.Execute("EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
'   dtc_codigo7.ReFill
'   dtc_codigo7.BoundText = Empty
'
'   Set dtc_desc7.RowSource = Nothing
'   Set dtc_desc7.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo7.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
'   dtc_desc7.ReFill
'   dtc_desc7.BoundText = Empty
'End Sub

Private Sub dtc_desc41_Click(Area As Integer)
    dtc_codigo41.BoundText = dtc_desc41.BoundText
    dtc_aux41.BoundText = dtc_desc41.BoundText
    dtc_valor41.BoundText = dtc_desc41.BoundText
End Sub

Private Sub dtc_desc42_Click(Area As Integer)
    dtc_codigo42.BoundText = dtc_desc42.BoundText
    dtc_aux42.BoundText = dtc_desc42.BoundText
'    Call pnivel8(dtc_codigo8.BoundText)
'    'dtc_desc9.Enabled = True
'    dtc_codigo9.Enabled = True
End Sub

'Private Sub pnivel8(codigo8 As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_documentos_respaldo where clasif_codigo = '" & codigo8 & "'"
'
'   Set dtc_codigo9.RowSource = Nothing
'   Set dtc_codigo9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo9.ReFill
'   dtc_codigo9.BoundText = Empty
'
'   Set dtc_desc9.RowSource = Nothing
'   Set dtc_desc9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc9.ReFill
'   dtc_desc9.BoundText = Empty
'End Sub

Private Sub dtc_desc43_Click(Area As Integer)
    dtc_codigo43.BoundText = dtc_codigo43.BoundText
    dtc_aux43.BoundText = dtc_desc43.BoundText
End Sub

Private Sub dtc_desc51_Click(Area As Integer)
    dtc_codigo51.BoundText = dtc_desc51.BoundText
    dtc_aux51.BoundText = dtc_desc51.BoundText
    dtc_desc51.ToolTipText = dtc_desc51.Text
End Sub

Private Sub dtc_desc52_Click(Area As Integer)
    dtc_codigo52.BoundText = dtc_desc52.BoundText
    dtc_aux52.BoundText = dtc_desc52.BoundText
    dtc_desc52.ToolTipText = dtc_desc52.Text
End Sub

Private Sub dtc_desc53_Click(Area As Integer)
    dtc_codigo53.BoundText = dtc_codigo53.BoundText
    dtc_aux53.BoundText = dtc_desc53.BoundText
    dtc_desc53.ToolTipText = dtc_desc53.Text
End Sub

Private Sub dtc_desc61_Click(Area As Integer)
    dtc_codigo61.BoundText = dtc_desc61.BoundText
    dtc_desc61.ToolTipText = dtc_desc61.Text
End Sub

Private Sub dtc_desc62_Click(Area As Integer)
    dtc_codigo62.BoundText = dtc_desc62.BoundText
    dtc_desc62.ToolTipText = dtc_desc62.Text
End Sub

Private Sub dtc_desc63_Click(Area As Integer)
    dtc_codigo63.BoundText = dtc_codigo63.BoundText
    dtc_desc63.ToolTipText = dtc_desc63.Text
End Sub

Private Sub dtc_desc71_Click(Area As Integer)
    dtc_codigo71.BoundText = dtc_desc71.BoundText
    dtc_desc71.ToolTipText = dtc_desc71.Text
End Sub

Private Sub dtc_desc72_Click(Area As Integer)
    dtc_codigo72.BoundText = dtc_desc72.BoundText
    dtc_desc72.ToolTipText = dtc_desc72.Text
End Sub

Private Sub dtc_desc73_Click(Area As Integer)
    dtc_codigo73.BoundText = dtc_codigo73.BoundText
    dtc_desc73.ToolTipText = dtc_desc73.Text
End Sub

Private Sub dtc_desc81_Click(Area As Integer)
    dtc_codigo81.BoundText = dtc_desc81.BoundText
    dtc_desc81.ToolTipText = dtc_desc81.Text
End Sub

Private Sub dtc_desc82_Click(Area As Integer)
    dtc_codigo82.BoundText = dtc_desc82.BoundText
    dtc_desc82.ToolTipText = dtc_desc82.Text
End Sub

Private Sub dtc_desc83_Click(Area As Integer)
    dtc_codigo83.BoundText = dtc_codigo83.BoundText
    dtc_desc83.ToolTipText = dtc_desc83.Text
End Sub

Private Sub dtc_desc91_Click(Area As Integer)
    dtc_codigo91.BoundText = dtc_desc91.BoundText
    dtc_desc91.ToolTipText = dtc_desc91.Text
End Sub

Private Sub dtc_desc92_Click(Area As Integer)
    dtc_codigo92.BoundText = dtc_desc92.BoundText
    dtc_desc92.ToolTipText = dtc_desc92.Text
End Sub

Private Sub dtc_desc93_Click(Area As Integer)
    dtc_codigo93.BoundText = dtc_codigo93.BoundText
    dtc_desc93.ToolTipText = dtc_desc93.Text
End Sub

Private Sub dtc_desc01_Click(Area As Integer)
    dtc_codigo01.BoundText = dtc_desc01.BoundText
    dtc_desc01.ToolTipText = dtc_desc01.Text
End Sub

Private Sub dtc_desc02_Click(Area As Integer)
    dtc_codigo02.BoundText = dtc_desc02.BoundText
    dtc_desc02.ToolTipText = dtc_desc02.Text
End Sub

Private Sub dtc_desc03_Click(Area As Integer)
    dtc_codigo03.BoundText = dtc_codigo03.BoundText
    dtc_desc03.ToolTipText = dtc_desc03.Text
End Sub

Private Sub OptFilGral1_Click()
    Set rs_aux6 = New ADODB.Recordset
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux6.RecordCount > 0 Then
        usuario2 = rs_aux6!beneficiario_codigo
        VAR_DA = rs_aux6!da_codigo
        VAR_DPTO = rs_aux6!depto_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.2"
        VAR_DPTO = "2"
    End If
    Set rs_datos = New Recordset
     If rs_datos.State = 1 Then rs_datos.Close
     Select Case VAR_DA
        Case "1.8"    'Cochabamba
            queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '4' ))) "
        Case "1.7"    'Santa Cruz
            If glusuario = "CURDININEA" Then        'SCZ
                queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG' AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' or left(edif_codigo,1) = '9' or left(edif_codigo,1) = '3'  ) )) "
            Else
                queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG' AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' ) )) "
            End If
            
        Case "1.2"    'La Paz - Comercial
            If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "GSOLIZ" Or glusuario = "ASANTIVAÑEZ" Then           'LPZ
                queryinicial = "select * From ao_solicitud_calculo_trafico WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMC')) "
            Else
                'queryinicial = "select * From ao_solicitud WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '1' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' or left(edif_codigo,1) = '9' )))  "
                queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '1' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' or left(edif_codigo,1) = '9'  ) )) "
            End If
        Case "1.3"    'La Paz - Modernizacion
            If glusuario = "ADMIN" Or glusuario = "JSAAVEDRA" Or glusuario = "CCOLODRO" Then
                queryinicial = "select * From ao_solicitud_calculo_trafico WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DNMOD') )"
            Else
                queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "'  AND left(edif_codigo,1) = '" & VAR_DPTO & "' ))"      'AND beneficiario_codigo_resp2 = '" & usuario2 & "'
            End If
        Case "1.9"    ' Chuquisaca
            queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG' AND unidad_codigo = '" & VAR_UORIGEN & "'  AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '5' or left(edif_codigo,1) = '6') )) "
        Case "1.4"    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From ao_solicitud_calculo_trafico WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC')) "
                Else
                    queryinicial = "select * From ao_solicitud_calculo_trafico WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC')) "
                End If
            End If
        Case Else    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From ao_solicitud_calculo_trafico WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC')) "
                Else
                    queryinicial = "select * From ao_solicitud_calculo_trafico WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC')) "
                End If
            End If
     End Select
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    'queryinicial = "Select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & parametro & "' and estado_codigo = 'REG' "
'    If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "GSOLIZ" Then
'        queryinicial = "Select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & parametro & "' and estado_codigo = 'REG' "
'    Else
'        queryinicial = "Select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & parametro & "' and estado_codigo = 'REG' AND beneficiario_codigo_resp = '" & usuario2 & "' "
'    End If
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    Set rs_aux6 = New ADODB.Recordset
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux6.RecordCount > 0 Then
        usuario2 = rs_aux6!beneficiario_codigo
        VAR_DA = rs_aux6!da_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.2"
    End If
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DA
        Case "1.8"    'Cochabamba
            queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '4' ))) "
        Case "1.7"    'Santa Cruz
            If glusuario = "CURDININEA" Then        'SCZ
                queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' or left(edif_codigo,1) = '9' or left(edif_codigo,1) = '3' )))  "
            Else
                queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' )))  "
            End If
            
        Case "1.2"    'La Paz - Comercial
            If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "GSOLIZ" Or glusuario = "ASANTIVAÑEZ" Then           'LPZ
                queryinicial = "select * From ao_solicitud_calculo_trafico WHERE (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMC') "
            Else
                queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '1' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' or left(edif_codigo,1) = '9' )))  "
                'queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '1' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' or left(edif_codigo,1) = '9'  ) )) "
            End If
        Case "1.3"    'La Paz - Modernizacion
            If glusuario = "ADMIN" Or glusuario = "JSAAVEDRA" Or glusuario = "CCOLODRO" Then
                queryinicial = "select * From ao_solicitud_calculo_trafico WHERE (unidad_codigo = 'DNMOD') "
            Else
                queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' )))"      'AND beneficiario_codigo_resp2 = '" & usuario2 & "'
            End If
        Case "1.9"    ' Chuquisaca
            queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '5' or left(edif_codigo,1) = '6' )))  "
        Case "1.4"    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From ao_solicitud_calculo_trafico WHERE (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC') "
                    'queryinicial = "select * From ao_solicitud WHERE estado_codigo = 'REG'  "
                Else
                    queryinicial = "select * From ao_solicitud_calculo_trafico WHERE (unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC') "
                End If
            End If
        Case Else    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From ao_solicitud_calculo_trafico WHERE (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC') "
                    'queryinicial = "select * From ao_solicitud WHERE estado_codigo = 'REG'  "
                Else
                    queryinicial = "select * From ao_solicitud_calculo_trafico WHERE (unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC') "
                End If
            End If
     End Select
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    'queryinicial = "Select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & parametro & "'  "
'    'queryinicial = "Select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & txt_codigo.Text & " "   'and estado_codigo = 'REG' "  '+ parametro
'    If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "BINFANTE" Or glusuario = "AURBINA" Or glusuario = "GSOLIZ" Then
'        queryinicial = "Select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & parametro & "'  "
'    Else
'        queryinicial = "Select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & parametro & "' AND beneficiario_codigo_resp = '" & usuario2 & "' "
'    End If
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
'    parametro = "estado_codigo" + " = " + "'REG'"
End Sub

'Private Sub Img_03_Click()
' If AdoPermiso.Recordset!ARCHIVO = "Cargar_Archivo" Then
'    MsgBox "No Existe el Archivo asociado al Registro, debe Cargarlo ...", vbExclamation, "Advertencia"
' Else
'   If GlServidor = "SRVPRO" Then
'      If AdoPermiso.Recordset!TipoPermiso = "VC" Then
'        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(AdoPermiso.Recordset!solicitud_codigo) & "\VACACIONES\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      Else
'        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(AdoPermiso.Recordset!solicitud_codigo) & "\LICENCIAS\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      End If
'   Else
'      If AdoPermiso.Recordset!TipoPermiso = "VC" Then
'        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(AdoPermiso.Recordset!solicitud_codigo) & "\VACACIONES\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      Else
'        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(AdoPermiso.Recordset!solicitud_codigo) & "\LICENCIAS\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      End If
'   End If
' End If
'
'End Sub

'Private Sub Img_CTO_Click()
' If Ado_Memo.Recordset!ARCHIVO = "Cargar_Archivo" Then
'    MsgBox "No Existe el Archivo Asociado al Contrato, debe Cargarlo ...", vbExclamation, "Advertencia"
' Else
'    'If GlServidor <> GlMaquina Then      ' "-" Then
'    If GlServidor = "SRVPRO" Then
'        'e = ShellExecute(Img_CTO, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_Memo.Recordset!solicitud_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
'        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_Memo.Recordset!solicitud_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'    Else
'        'e = ShellExecute(Img_CTO, "open", App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_Memo.Recordset!solicitud_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
'        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_Memo.Recordset!solicitud_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'    End If
' End If
'End Sub

'Private Sub Img_CV_Click()
''    Dim e As Long
'  If swnuevo <> "X" Then
'    If Ado_datos.Recordset!ARCHIVO_HOJAVIDA = "Cargar_Archivo" Then
'      NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\"
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "C_V"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\"
'         ' e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario.AdoMovilidad.Recordset!solicitud_codigo) & "\FINIQUITO\" & Trim(Ado_Auxiliar.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      Else
'         e = NombreCarpeta
'      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
'    Else
'      'MsgBox ""
'      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'          NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\"
'          Frmexporta.DirDestino.Path = NombreCarpeta
'          GlArch = "C_V"
'          'If GlServidor <> GlMaquina Then      ' "-" Then
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\"
'          Else
'            e = NombreCarpeta
'          End If
'          Frmexporta.DirDestino2.Path = e
'          Frmexporta.Show vbModal
'      End If
'    End If
'  End If
'  If GlServidor = "SRVPRO" Then
'        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\" & Trim(Ado_datos.Recordset!ARCHIVO_VAC), vbNullString, vbNullString, vbNormalFocus)
'  Else
'        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\" & Trim(Ado_datos.Recordset!ARCHIVO_VAC), vbNullString, vbNullString, vbNormalFocus)
'  End If
'End Sub
'
'Private Sub Img_Foto_Click()
'  If swnuevo <> "X" Then
'    If Ado_datos.Recordset!ARCHIVO_FOTO = "Cargar_Archivo" Then
'      NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "FOT"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
'      Else
'         e = NombreCarpeta
'      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
'    Else
'      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'          NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
'          Frmexporta.DirDestino.Path = NombreCarpeta
'          GlArch = "FOT"
'          'If GlServidor <> GlMaquina Then      ' "-" Then
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
'          Else
'            e = NombreCarpeta
'          End If
'          Frmexporta.DirDestino2.Path = e
'          Frmexporta.Show vbModal
'      End If
'    End If
'
'    Dim ARCH_FOTO As String
'    If GlServidor = "SRVPRO" Then
'        ARCH_FOTO = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("solicitud_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    Else
'        ARCH_FOTO = App.Path + "\" & Trim(GLCarpeta2) & "\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("solicitud_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    End If
'    If Guardar_Imagen(db, "Select Foto From Gc_beneficiario Where solicitud_codigo= '" & Ado_datos.Recordset("solicitud_codigo") & "' ", "Foto", ARCH_FOTO) Then
'        MsgBox "Se cargo la Imagen Correctamente !!"
'    Else
'        MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
'    End If
'  End If
'End Sub

'Private Sub SSTab1_DblClick()
'    If SSTab1.Tab = 0 Then
'    End If
'End Sub


Private Sub Form_Unload(Cancel As Integer)
  If glPersNew = "P" Then
  End If
  glPersNew = "N"
   
'   If (rstbeneficiario.State = adStateClosed) Then rstbeneficiario.Close
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub

Private Sub Txt_campo41_LostFocus()
    If Txt_campo41.Text < 1100 Then
        var_campoc31 = "2.4"
    Else
        var_campoc31 = "2"
    End If
End Sub

Private Sub Txt_campo42_LostFocus()
    If Txt_campo42.Text < 1100 Then
        var_campoc32 = "2.4"
    Else
        var_campoc32 = "2"
    End If
End Sub

Private Sub Txt_campo43_LostFocus()
    If Txt_campo43.Text < 1100 Then
        var_campoc33 = "2.4"
    Else
        var_campoc33 = "2"
    End If
End Sub

Private Sub correl_bien()
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "Select * from fc_partida_gasto where par_codigo = '43340'   ", db, adOpenDynamic ', adOpenKeyset ', adOpenStatic
    If Not rs_aux2.EOF Then
         VAR_COD2 = rs_aux2!correlativo + 1
         db.Execute "UPDATE fc_partida_gasto SET correlativo = '" & Val(VAR_COD2) & "' where par_codigo = '43340' "
         'rs_aux2!correlativo = rs_aux2!correlativo + 1  'Val(VAR_COD2)
         'rs_aux2.Update
    End If
End Sub

Private Sub Txt_campo44_LostFocus()
    If Txt_campo44.Text < 1100 Then
        var_campoc34 = "2.4"
    Else
        var_campoc34 = "2"
    End If
End Sub

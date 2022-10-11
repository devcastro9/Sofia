VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_almacen_salida_zonas 
   BackColor       =   &H00000000&
   Caption         =   "Almacen - Salida de Insumos"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   5115
      Left            =   120
      Picture         =   "frm_almacen_salida_zonas.frx":0000
      ScaleHeight     =   5055
      ScaleWidth      =   1035
      TabIndex        =   38
      Top             =   4560
      Width           =   1095
      Begin VB.CommandButton BtnAnlDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Anular"
         Height          =   640
         Left            =   120
         Picture         =   "frm_almacen_salida_zonas.frx":6C032
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Cambia el Horario a NO LABORABLE (Anula horario)"
         Top             =   1560
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   120
         Picture         =   "frm_almacen_salida_zonas.frx":6C474
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   840
         Width           =   765
      End
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   120
         Picture         =   "frm_almacen_salida_zonas.frx":6C8B6
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Adiciona Detalle"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000018&
         Caption         =   "Cronogr."
         Height          =   640
         Left            =   120
         Picture         =   "frm_almacen_salida_zonas.frx":6CCF8
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   2280
         Visible         =   0   'False
         Width           =   765
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "frm_almacen_salida_zonas.frx":6E47A
      ScaleHeight     =   960
      ScaleWidth      =   15240
      TabIndex        =   26
      Top             =   120
      Width           =   15300
      Begin VB.CommandButton BtnImprimir2 
         BackColor       =   &H00808000&
         Caption         =   "R-224"
         Height          =   720
         Left            =   6840
         Picture         =   "frm_almacen_salida_zonas.frx":DA4AC
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   720
         Left            =   5160
         Picture         =   "frm_almacen_salida_zonas.frx":DAA69
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "frm_almacen_salida_zonas.frx":DAEAB
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "frm_almacen_salida_zonas.frx":DB0B5
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
         Caption         =   "R-302"
         Height          =   720
         Left            =   4320
         Picture         =   "frm_almacen_salida_zonas.frx":DB66D
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   6000
         Picture         =   "frm_almacen_salida_zonas.frx":DBC2A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "frm_almacen_salida_zonas.frx":DBE34
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "frm_almacen_salida_zonas.frx":DCAFE
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "frm_almacen_salida_zonas.frx":DD0DE
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "frm_almacen_salida_zonas.frx":DD702
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COTIZA"
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
         Left            =   10515
         TabIndex        =   37
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_almacen_salida_zonas.frx":DD90C
      ScaleHeight     =   915
      ScaleWidth      =   15240
      TabIndex        =   22
      Top             =   120
      Width           =   15300
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "frm_almacen_salida_zonas.frx":14993E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "frm_almacen_salida_zonas.frx":149B48
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOLICITUD DE COTIZACIÓN"
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
         Left            =   8460
         TabIndex        =   25
         Top             =   300
         Width           =   4185
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00000000&
      Caption         =   "ASIGNACION DE EDIFICIOS"
      ForeColor       =   &H00FFFFC0&
      Height          =   5175
      Left            =   1320
      TabIndex        =   9
      Top             =   4560
      Width           =   14085
      Begin MSDataGridLib.DataGrid dg_det1 
         Height          =   4815
         Left            =   195
         TabIndex        =   10
         Top             =   240
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   8493
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "zona_edif_orden"
            Caption         =   "Nro"
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
            DataField       =   "edif_codigo"
            Caption         =   "Cod_Edificio"
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
            DataField       =   "edif_descripcion"
            Caption         =   "Nombre_Edificio"
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
            DataField       =   "edif_tipo"
            Caption         =   "Tipo_Edificio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "nro_total_horas"
            Caption         =   "Nro.Horas"
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
            DataField       =   "estado_activo"
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
         BeginProperty Column06 
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Tec.Mantenimiento"
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
            DataField       =   "beneficiario_codigo_resp2"
            Caption         =   "Tec.Emergencias"
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
         BeginProperty Column08 
            DataField       =   ""
            Caption         =   "Cobrador"
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
            DataField       =   "observaciones"
            Caption         =   "Observaciones"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
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
               Object.Visible         =   -1  'True
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   4320
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1470.047
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   2505.26
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000040&
      Height          =   3360
      Left            =   5880
      TabIndex        =   5
      Top             =   1080
      Width           =   9540
      Begin VB.TextBox Txt_campo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         DataField       =   "zpiloto_descripcion"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   7245
      End
      Begin VB.TextBox txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         DataField       =   "zpiloto_codigo"
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
         Height          =   300
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         DataField       =   "fmes_nro_hrs_habiles"
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
         ForeColor       =   &H00FFFF80&
         Height          =   290
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   7
         Top             =   480
         Width           =   1170
      End
      Begin VB.TextBox Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         DataField       =   "estado_codigo"
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
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2760
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "fecha_registro"
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
         Height          =   315
         Left            =   2040
         TabIndex        =   0
         Top             =   2760
         Width           =   1695
         _ExtentX        =   2990
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   87228419
         CurrentDate     =   41678
         MaxDate         =   109939
         MinDate         =   36526
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         DataField       =   "munic_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7920
         TabIndex        =   18
         Top             =   2160
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "munic_codigo"
         BoundColumn     =   "munic_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo9 
         DataField       =   "prov_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7920
         TabIndex        =   19
         Top             =   1560
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "prov_codigo"
         BoundColumn     =   "prov_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc9 
         DataField       =   "prov_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2040
         TabIndex        =   20
         Top             =   1560
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "prov_descripcion"
         BoundColumn     =   "prov_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         DataField       =   "munic_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2040
         TabIndex        =   21
         Top             =   2160
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "munic_descripcion"
         BoundColumn     =   "munic_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad / Lugar"
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
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable Zona"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   1740
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   6
         Left            =   7200
         TabIndex        =   15
         Top             =   2760
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Registro"
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
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   1380
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Compra"
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
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1425
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Zona Piloto (Ruta)"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1005
         Width           =   1605
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFC0&
      Height          =   3360
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5655
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
         Left            =   3120
         TabIndex        =   4
         Top             =   2955
         Width           =   915
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
         TabIndex        =   3
         Top             =   2955
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   2880
         Width           =   5385
         _ExtentX        =   9499
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
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   2610
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   4604
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "zpiloto_codigo"
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
         BeginProperty Column01 
            DataField       =   "fecha_registro"
            Caption         =   "fecha_registro"
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
            DataField       =   "zpiloto_descripcion"
            Caption         =   "Zona.Piloto"
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
            DataField       =   "prov_codigo"
            Caption         =   "Provincia"
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
            DataField       =   "munic_codigo"
            Caption         =   "Municipio"
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
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   2775.118
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
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
   Begin MSAdodcLib.Adodc Ado_datos2 
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   4320
      Top             =   9840
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
   Begin MSAdodcLib.Adodc Ado_datos51 
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
   Begin MSAdodcLib.Adodc Ado_datos61 
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
   Begin MSAdodcLib.Adodc Ado_datos31 
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
   Begin Crystal.CrystalReport CR01 
      Left            =   4320
      Top             =   10200
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
   Begin MSAdodcLib.Adodc Ado_detalle1 
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
   Begin MSAdodcLib.Adodc Ado_datos21 
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
End
Attribute VB_Name = "frm_almacen_salida_zonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset

Dim rsNada As New ADODB.Recordset

'Dim rs_datos21 As New ADODB.Recordset
'Dim rs_datos31 As New ADODB.Recordset
'Dim rs_datos41 As New ADODB.Recordset
'Dim rs_datos51 As New ADODB.Recordset
'Dim rs_datos61 As New ADODB.Recordset

'Dim rstbeneficiario As New ADODB.Recordset
'Dim rst_ben As New ADODB.Recordset
'Dim RsTmp As New ADODB.Recordset
'Dim rs_aux1 As New ADODB.Recordset
'Dim rs_aux2 As New ADODB.Recordset
'Dim rs_aux3 As New ADODB.Recordset
'Dim rs_aux4 As New ADODB.Recordset
Dim rs_det1 As New ADODB.Recordset

'Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim queryinicial As String

'OTROS
Dim VAR_MOD, VAR_MOD1, VAR_MOD2 As String
Dim SQL_FOR As String
Dim sql As String
'Dim swnuevo As String
Dim imag2 As Long

Dim sino As String
Dim NombreCarpeta, e As String
Dim parametro As String
Dim var_titulo As String
Dim var_cod, VAR_GES As String
Dim VAR_VAL, VAR_ARCH, VAR_ARCH2 As String
Dim VAR_SW As String

Dim VAR_AUX, VAR_CONT2 As Double

Dim var_campoc31, var_campoc32, var_campoc33, var_campoc34 As Double
Dim var_campod11, var_campod12, var_campod13, var_campod14 As Double
Dim var_campoe11, var_campoe12, var_campoe13, var_campoe14 As Double
Dim var_campoe21, var_campoe22, var_campoe23, var_campoe24 As Double
Dim var_campoe31, var_campoe32, var_campoe33, var_campoe34 As Double
Dim var_campoe41, var_campoe42, var_campoe43, var_campoe44 As Double
Dim var_campog11, var_campog12, var_campog13, var_campog14 As Double
Dim var_campog21, var_campog22, var_campog23, var_campog24 As Double

Dim mvBookMark, marca1 As Variant
Dim mbDataChanged As Boolean

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
     '<-- Inicio                Identificación del Cliente                Fin -->
     If VAR_SW <> "MOD" Then
'        Select Case dtc_codigo2.Text
'            Case "1"
'            Case "2"
'            Case "3"
'                Call ABRIR_TABLA_DET3
'            Case "4"
'
'        End Select
'        Call ABRIR_TABLA_AUX2
        If Ado_datos.Recordset.RecordCount > 0 Then
            Call ABRIR_TABLA_DET
'            If lbl_texto1.Caption <> "" And lbl_texto1.Caption <> "0" Then
'                lbl_texto2.Caption = UCase(MonthName(Val(lbl_texto1.Caption)))
'            End If
            'mes2 = MonthName(Month(DTPFec_Inicio.Value))
        End If
    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det1.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
End Sub

Private Sub BtnAddDetalle_Click()
  marca1 = Ado_datos.Recordset.BookMark
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    Fra_datos.Enabled = False

'    aw_p_ao_solicitud_cotiza_detalle.txt_codigo.Caption = Me.txt_codigo     ' Nro. Negociacion (Cod.solicitud)
'    aw_p_ao_solicitud_cotiza_detalle.Txt_campo1.Caption = Me.dtc_codigo1.Text       ' Codigo Unidad
'    aw_p_ao_solicitud_cotiza_detalle.Txt_descripcion.Caption = Me.dtc_desc1.Text    ' Descripcion Unidad
'    aw_p_ao_solicitud_cotiza_detalle.Txt_Correl.Caption = Me.txt_codigo1.Text       ' Nro. Cotización
'    aw_p_ao_solicitud_cotiza_detalle.Txt_campo2.Caption = Me.dtc_codigo3.Text       ' Codigo Edificio
'    Ado_detalle1.Recordset.AddNew
'    aw_p_ao_solicitud_cotiza_detalle.Show vbModal

    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
  If Ado_datos.Recordset!estado_codigo = "REG" Then
'    Call OptFilGral1_Click
  Else
    Call OptFilGral2_Click
  End If
  'Call ABRIR_TABLA_DET
  Ado_datos.Recordset.Move marca1 - 1
End Sub

Private Sub BtnAnlDetalle_Click()
   If Ado_detalle1.Recordset("estado_activo") = "REG" Then
      sino = MsgBox("Está Seguro de cambiar a HORARIO NO LABORABLE ? (Este ya no será considerado en el Cronograma) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        Ado_detalle1.Recordset!estado_activo = "ANL"
        Ado_detalle1.Recordset!observaciones = "HORARIO NO LABORABLE"
        Ado_detalle1.Recordset.Update
      End If
   Else
        MsgBox "No se puede ANULAR, el registro ya fue APROBADO o ya fue ANULADO anteriormente ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAprobar_Click()
'  On Error GoTo UpdateErr
'   Set rs_aux2 = New ADODB.Recordset
'   rs_aux2.Open "Select * from ao_solicitud_costos where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
'   If rs_aux2.RecordCount > 0 Then
'        VAR_CONT2 = rs_aux2.RecordCount
'   End If
'   'If rs_datos!estado_codigo = "REG" And Ado_datos.Recordset!correl_edificacion > 0 Then
'   If rs_datos!estado_codigo = "REG" And VAR_CONT2 > 0 Then
'      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'
''        Select Case dtc_codigo2.Text
''            Case "1"
''            Case "2"
''            Case "3"
'                Set rs_aux1 = New ADODB.Recordset
'                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'                SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "    "
'                rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'                If rs_aux1.RecordCount > 0 Then
'                    MsgBox "Una Cotización anterior ya fue Aprobada, el Registro Actual se adicionará al que fue aprobado anteriormente ..."
'                    '    var_cod = 0
'                    '    Exit Sub
'                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + Ado_datos.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + Ado_datos.Recordset!cotiza_precio_total_dol
'                Else
'                    'CREA VENTA CABECERA
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
'                    rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
'                    If Not rs_aux2.EOF Then
'                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                    End If
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    rs_aux2.Open "Select beneficiario_codigo as Codigo from ao_solicitud where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
'                    If Not rs_aux2.EOF Then
'                        VAR_AUX = rs_aux2!Codigo
'                    End If
'                    rs_aux1.AddNew
'                    'var_cod = rs_aux1.RecordCount + 1
'                    rs_aux1!ges_gestion = Year(Date)
'                    rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
'                    rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
'                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'                    rs_aux1!venta_codigo = var_cod
'                    rs_aux1!beneficiario_codigo = VAR_AUX
'                    rs_aux1!venta_monto_total_bs = Ado_datos.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_monto_total_dol = Ado_datos.Recordset!cotiza_precio_total_dol
'                    rs_aux1!venta_monto_cobrado_bs = 0
'                    rs_aux1!venta_monto_cobrado_dol = 0
'                    rs_aux1!venta_saldo_p_cobrar_bs = Ado_datos.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_saldo_p_cobrar_dol = Ado_datos.Recordset!cotiza_precio_total_dol
'                    rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
'                    rs_aux1!estado_codigo = "REG"
'                    rs_aux1!fecha_registro = Date
'                    rs_aux1!usr_codigo = glusuario
'                    rs_aux1.Update
''                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
'                End If
'                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
''            Case "4"
''        End Select
'        'GRABA VENTA DETALLE
'        If var_cod = "" Then
'            var_cod = rs_aux1!venta_codigo
'        End If
'        Set rs_aux3 = New ADODB.Recordset
'        If rs_aux3.State = 1 Then rs_aux3.Close
'        'rs_aux3.Open "Select * from ao_ventas_detalle where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
'        rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & Year(Date) & "'   ", db, adOpenKeyset, adLockOptimistic
'        'If rs_aux3.RecordCount > 0 Then
'            'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'        'Else
'            VAR_AUX = rs_aux3.RecordCount + 1
'            rs_aux3.AddNew
'            rs_aux3!ges_gestion = Year(Date)
'            rs_aux3!venta_codigo = var_cod
'            rs_aux3!venta_codigo_det = VAR_AUX
'            rs_aux3!bien_codigo = Ado_datos.Recordset!bien_codigo
'            rs_aux3!venta_det_cantidad = Ado_datos.Recordset!cotiza_cantidad
'            rs_aux3!venta_precio_unitario_bs = 0
'            rs_aux3!venta_descuento_bs = 0
'            rs_aux3!venta_precio_total_bs = 0
'            rs_aux3!venta_precio_unitario_dol = 0
'            rs_aux3!venta_descuento_dol = 0
'            rs_aux3!venta_precio_total_dol = 0
''            rs_aux3!concepto_venta = dtc_desc21.Text + " - " + Ado_datos.Recordset!bien_codigo
'            'ok
'            rs_aux3!grupo_codigo = "40000"
'            rs_aux3!subgrupo_codigo = "43000"
'            rs_aux3!par_codigo = "43340"
'            'ok
'            rs_aux3!tipo_descuento = 0
'            rs_aux3!almacen_codigo = 0
'            rs_aux3!modelo_codigo1 = Ado_datos.Recordset!modelo_codigo
'            rs_aux3!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h
'            rs_aux3!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x
'            rs_aux3!modelo_elegido = "N"
'            rs_aux3!modelo_elegido_h = "N"
'            rs_aux3!modelo_elegido_x = "N"
'            'rs_aux3!estado_codigo = "REG"
'            rs_aux3!fecha_registro = Date
'            rs_aux3!usr_codigo = glusuario
'            rs_aux3.Update
'        'End If
'        'INI GRABA ALMACEN DETALLE (EN LA ENTREGA EN OBRA)
''        Set rs_aux4 = New ADODB.Recordset
''        If rs_aux4.State = 1 Then rs_aux4.Close
''        rs_aux4.Open "Select * from ao_almacen_detalle where almacen_codigo = 0 and bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "'   ", db, adOpenKeyset, adLockOptimistic
''        If rs_aux4.RecordCount = 0 Then
''            'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
''            rs_aux4.AddNew
''            rs_aux4!almacen_codigo = 0
''            rs_aux4!bien_codigo = Ado_datos.Recordset!bien_codigo
''            rs_aux4!grupo_codigo = "40000"
''            rs_aux4!subgrupo_codigo = "43000"
''            rs_aux4!par_codigo = "43340"
''            rs_aux4!stock_ingreso = 1
''            rs_aux4!stock_salida = 0
''            rs_aux4!stock_actual = 1
''            rs_aux4!estado_codigo = "REG"
''            rs_aux4!usr_codigo = GlUsuario
''            rs_aux4!fecha_registro = Date
''            rs_aux4.Update
''        End If
'        'R-222 "COTIZACION DE EQUIPOS PARA EL CLIENTE"
'        Set rs_aux2 = New ADODB.Recordset
'        If rs_aux2.State = 1 Then rs_aux2.Close
'        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos.Recordset!doc_codigo & "'  "
'        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'            rs_datos!doc_numero = rs_aux2!correl_doc
'            'Txt_campo1.Caption = rs_aux2!correl_doc
'            rs_aux2.Update
'        End If
'        'rs_datos!doc_numero = Txt_campo1.Caption
'        'REVISAR !!! JQA 2014_07_08
'        'VAR_ARCH = RTrim(RTrim(rs_datos!doc_codigo) + "-") + LTrim(Str(rs_datos!doc_numero))
'        VAR_ARCH = "COM_" + RTrim(RTrim(rs_datos!doc_codigo) + "-") + LTrim(Str(rs_datos!doc_numero))
'        rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
'        rs_datos!archivo_respaldo_cargado = "N"
'        'R-224 "PROPUESTA DE COTIZACION DE EQUIPOS PARA EL CLIENTE"
'        Set rs_aux2 = New ADODB.Recordset
'        If rs_aux2.State = 1 Then rs_aux2.Close
'        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos.Recordset!doc_codigo2 & "'  "
'        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'            rs_datos!doc_numero2 = rs_aux2!correl_doc
'            rs_aux2.Update
'        End If
'        VAR_ARCH2 = "COM_" + RTrim(RTrim(rs_datos!doc_codigo2) + "-") + LTrim(Str(rs_datos!doc_numero2))
'        rs_datos!archivo_respaldo2 = VAR_ARCH2 + ".PDF"
'        rs_datos!archivo_respaldo_cargado2 = "N"
'
'        rs_datos!estado_codigo = "APR"
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

End Sub

Private Sub BtnBuscar_Click()
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
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelUpdate
        Call ABRIR_TABLA
        rs_datos.MoveFirst
        'mbDataChanged = False
        Fra_datos.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        VAR_SW = ""
    End If

End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
     rs_datos!fmes_fecha_registro = DTPfecha1.Value
'     rs_datos!beneficiario_codigo_resp = dtc_codigo4.Text
     rs_datos!observaciones = Txt_campo2.Text


'     'rs_datos!Foto = Date
'     'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
'     'rs_datos!archivo_foto_cargado = "N"
'     'hora_registro
'     'rs_datos!hora_aprueba = ""
'     rs_datos!proceso_codigo = "COM"
'     rs_datos!subproceso_codigo = "COM-01"
'     rs_datos!etapa_codigo = "COM-01-04"
'     rs_datos!clasif_codigo = "COM"
'     rs_datos!doc_codigo = "R-222"
'     rs_datos!doc_numero = "0"  'txt_campo1.Text
'     rs_datos!clasif_codigo2 = "COM"
'     rs_datos!doc_codigo2 = "R-224"
'     rs_datos!doc_numero2 = "0"  'txt_campo1.Text
'     rs_datos!poa_codigo = "3.1.1"

     rs_datos!fecha_registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     rs_datos.Update    'Batch 'adAffectAll
'     db.Execute "Update to_cronograma_diario Set beneficiario_codigo_resp = " & dtc_codigo4.Text & " Where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'   "
     Call OptFilGral2_Click
     rs_datos.MoveFirst
'     mbDataChanged = False

     Fra_datos.Enabled = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
     'dtc_desc1.BackColor = &HFFFFC0
     VAR_SW = ""
'     dtc_codigo9.Enabled = True

  End If
'  dtc_desc1.Visible = True
'  lbl_aux1.Visible = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub valida_campos()
  'Valida compos para editables
'  If (dtc_codigo1.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (dtc_codigo3.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If (dtc_codigo2 = "") Then
    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If (Txt_campo2.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If

End Sub


Private Sub BtnImprimir_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim IResult As Integer
    'Dim co As New ADODB.Command
    CR01.ReportFileName = App.Path & "\Reportes\tecnico\tr_R-302_cronograma_mensual.rpt"
    CR01.WindowShowPrintSetupBtn = True
    CR01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
    Select Case Me.Ado_datos.Recordset!unidad_codigo_tec
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
      End Select
      'Cmb_Mes.Text = "ENERO"
      CR01.Formulas(0) = "titulo = '" & var_titulo & "' "
      CR01.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
'      CR01.Formulas(2) = "periodo = '" & lbl_texto2 & "' "
      'CR01.Formulas(2) = "periodo = '" & Cmb_Mes & "' "

'    cr01.StoredProcParam(0) = "2015"    'Me.Ado_datos.Recordset!ges_gestion
'    cr01.StoredProcParam(1) = "DNMAN"   'Me.Ado_datos.Recordset!unidad_codigo_tec
'    cr01.StoredProcParam(2) = 0     'Me.Ado_datos.Recordset!zpiloto_codigo
'    cr01.StoredProcParam(3) = 1     'Me.Ado_datos.Recordset!fmes_correl

    CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo_tec
    CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!zpiloto_codigo
    CR01.StoredProcParam(3) = Me.Ado_datos.Recordset!fmes_correl

    IResult = CR01.PrintReport
    If IResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnImprimir2_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim IResult As Integer
    'Dim co As New ADODB.Command
    CR01.ReportFileName = App.Path & "\Reportes\comercial\R-224_ar_cotiza_venta_cliente.rpt"
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
    CR01.StoredProcParam(4) = Me.Ado_datos.Recordset!cotiza_codigo
    IResult = CR01.PrintReport
    If IResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnModDetalle_Click()
  marca1 = Ado_datos.Recordset.BookMark
  If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
    swnuevo = 2
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    Fra_datos.Enabled = False

'    Select Case dtc_codigo2.Text
'        Case "1"
'        Case "2"
'        Case "3"
        'Call ABRIR_TABLA_DET

        aw_p_ao_solicitud_cotiza_detalle.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo") ' Nro. Negociacion (Cod.solicitud)
        aw_p_ao_solicitud_cotiza_detalle.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")    ' Codigo Unidad
'        aw_p_ao_solicitud_cotiza_detalle.Txt_descripcion.Caption = Me.dtc_desc1.Text                        ' Descripcion Unidad
        aw_p_ao_solicitud_cotiza_detalle.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("cotiza_codigo")    ' Nro. Cotización
        aw_p_ao_solicitud_cotiza_detalle.Txt_campo2.Caption = Me.Ado_detalle1.Recordset("edif_codigo")      ' Codigo Edificio

        aw_p_ao_solicitud_cotiza_detalle.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("codigo_costo")     ' Codigo Costo
        aw_p_ao_solicitud_cotiza_detalle.dtc_desc1.BoundText = aw_p_ao_solicitud_cotiza_detalle.dtc_codigo1.BoundText
        aw_p_ao_solicitud_cotiza_detalle.dtc_aux1.BoundText = aw_p_ao_solicitud_cotiza_detalle.dtc_codigo1.BoundText
        aw_p_ao_solicitud_cotiza_detalle.dtc_aux2.BoundText = aw_p_ao_solicitud_cotiza_detalle.dtc_codigo1.BoundText

        aw_p_ao_solicitud_cotiza_detalle.Txt_campo3.Text = Me.Ado_detalle1.Recordset("costo_porcentaje")    ' % Costo

'        If txt_monto1.Text = "0" Or txt_monto1.Text = "" Then
'            aw_p_ao_solicitud_cotiza_detalle.txt_monto01.Caption = "0"                  ' Monto Modelo1(ME)
'        Else
'            aw_p_ao_solicitud_cotiza_detalle.txt_monto01.Caption = Me.txt_monto1.Text   ' Monto Modelo1(ME)
'        End If
'        If txt_monto1.Text = "0" Or txt_monto1.Text = "" Then
'            aw_p_ao_solicitud_cotiza_detalle.txt_monto02.Caption = "0"                  ' Monto Modelo2(ME)
'        Else
'            aw_p_ao_solicitud_cotiza_detalle.txt_monto02.Caption = Me.txt_monto5.Text   ' Monto Modelo2(ME)
'        End If
'        If txt_monto1.Text = "0" Or txt_monto1.Text = "" Then
'            aw_p_ao_solicitud_cotiza_detalle.txt_monto03.Caption = "0"                  ' Monto Modelo3(ME)
'        Else
'            aw_p_ao_solicitud_cotiza_detalle.txt_monto03.Caption = Me.txt_monto9.Text   ' Monto Modelo3(ME)
'        End If
'
        aw_p_ao_solicitud_cotiza_detalle.Txt_campo4.Text = Me.Ado_detalle1.Recordset("costo_observaciones") ' Observaciones


        aw_p_ao_solicitud_cotiza_detalle.Show vbModal
'        Case "4"
'
'    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    'If Ado_datos.Recordset!estado_codigo = "REG" Then
        Fra_datos.Enabled = True
         fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        VAR_SW = "MOD"
    '    BtnVer.Visible = True
    'Else
    '  MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
    'End If
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub BtnVer_Click()
    'ARREGLO 1
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc11 = dtc_aux41.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc21 = dtc_aux51.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc31 = IIf(IsNull(Ado_datos.Recordset!trafico_c_time_entrada_salida), 0, Ado_datos.Recordset!trafico_c_time_entrada_salida)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campod11 = IIf(IsNull(Ado_datos.Recordset!trafico_d_num_paradas_probables), 0, Ado_datos.Recordset!trafico_d_num_paradas_probables)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe11 = IIf(IsNull(Ado_datos.Recordset!trafico_e_tiempo_recorrido), 0, Ado_datos.Recordset!trafico_e_tiempo_recorrido)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe21 = IIf(IsNull(Ado_datos.Recordset!trafico_e_tiempo_asc_desaceleracion), 0, Ado_datos.Recordset!trafico_e_tiempo_asc_desaceleracion)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe31 = IIf(IsNull(Ado_datos.Recordset!trafico_e_tiempo_apertura_cierre), 0, Ado_datos.Recordset!trafico_e_tiempo_apertura_cierre)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe41 = IIf(IsNull(Ado_datos.Recordset!trafico_e_tiempo_entrada_salida), 0, Ado_datos.Recordset!trafico_e_tiempo_entrada_salida)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof11 = IIf(IsNull(Ado_datos.Recordset!trafico_f_tiempo_recorrido), 0, Ado_datos.Recordset!trafico_f_tiempo_recorrido)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof21 = IIf(IsNull(Ado_datos.Recordset!trafico_f_time_asc_desaceleracion), 0, Ado_datos.Recordset!trafico_f_time_asc_desaceleracion)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof31 = IIf(IsNull(Ado_datos.Recordset!trafico_f_time_apertura_cierre), 0, Ado_datos.Recordset!trafico_f_time_apertura_cierre)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof41 = IIf(IsNull(Ado_datos.Recordset!trafico_f_time_entrada_salida), 0, Ado_datos.Recordset!trafico_f_time_entrada_salida)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog11 = IIf(IsNull(Ado_datos.Recordset!trafico_g_capacidad_tiempo_cti), 0, Ado_datos.Recordset!trafico_g_capacidad_tiempo_cti)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog21 = IIf(IsNull(Ado_datos.Recordset!trafico_g_capacidad_total_arreglo), 0, Ado_datos.Recordset!trafico_g_capacidad_total_arreglo)

End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3A_Click(Area As Integer)
'    dtc_desc3A.BoundText = dtc_codigo3A.BoundText
End Sub

Private Sub dtc_codigo9_Click(Area As Integer)
    dtc_desc9.BoundText = dtc_codigo9.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

'Private Sub dtc_codigo4_Click(Area As Integer)
'    dtc_desc4.BoundText = dtc_codigo4.BoundText
'End Sub

Private Sub dtc_desc3A_Click(Area As Integer)
'    dtc_codigo3A.BoundText = dtc_desc3A.BoundText
End Sub

Private Sub dtc_desc9_Click(Area As Integer)
    dtc_codigo9.BoundText = dtc_desc9.BoundText
    Call pnivel9(dtc_codigo9.BoundText)
    dtc_desc2.Enabled = True
End Sub

Private Sub pnivel9(codigo9 As String)
   Dim strConsultaF As String

   strConsultaF = "select * from gc_municipio where prov_codigo = '" & codigo9 & "'"
   Set dtc_codigo2.RowSource = Nothing
   Set dtc_codigo2.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo2.ReFill
   dtc_codigo2.BoundText = Empty

   Set dtc_desc2.RowSource = Nothing
   Set dtc_desc2.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc2.ReFill
   dtc_desc2.BoundText = Empty
End Sub

'Private Sub dtc_desc4_Click(Area As Integer)
'    dtc_codigo4.BoundText = dtc_desc4.BoundText
'End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    'Fra_Gestion.Visible = True
'    VAR_GES = Cmb_gestion.Text
    parametro = Aux
    Call ABRIR_TABLAS_AUX
    Call OptFilGral2_Click

    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    'lbl_aux1.Visible = False
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
   'If Not Ado_datos.Recordset.EOF Then
            'SSTab1.Tab = 0
            'SSTab1.TabEnabled(0) = True
            ''SSTab1.TabEnabled(1) = False
            'SSTab1.TabVisible(1) = False
   'End If
End Sub

Private Sub ABRIR_TABLAS_AUX()
'    'gc_unidad_ejecutora
'    Set rs_datos1 = New ADODB.Recordset
'    If rs_datos1.State = 1 Then rs_datos1.Close
'    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
'    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora ", db, adOpenStatic
'    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
'
'    'tc_zonas_piloto
'    Set rs_datos3 = New ADODB.Recordset
'    If rs_datos3.State = 1 Then rs_datos3.Close
'    rs_datos3.Open "Select * from tc_zonas_piloto order by zpiloto_descripcion ", db, adOpenStatic
'    Set Ado_datos3.Recordset = rs_datos3
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
'
'    'Beneficiario Funcionario CGI (Vendedor, Cobrador, Adm, etc.)
'    Set rs_datos4 = New ADODB.Recordset
'    If rs_datos4.State = 1 Then rs_datos4.Close
'    rs_datos4.Open "rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
'    Set Ado_datos4.Recordset = rs_datos4
'    dtc_desc4.BoundText = dtc_codigo4.BoundText

    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from gc_municipio where region_codigo = 'SI' order by munic_descripcion", db, adOpenStatic
    'rs_datos2.Open "gp_listar_gc_municipio", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    'gc_provincia
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_provincia  ", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_desc9.BoundText = dtc_codigo9.BoundText
End Sub

'Private Sub dtc_codigo1_Click(Area As Integer)
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
'End Sub
'
'Private Sub dtc_codigo3_Click(Area As Integer)
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
'End Sub

'Private Sub dtc_desc1_Click(Area As Integer)
'    dtc_codigo1.BoundText = dtc_desc1.BoundText
''    Call pnivel1(dtc_codigo1.BoundText)
''    dtc_desc10.Enabled = True
'End Sub

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

'Private Sub dtc_desc1_LostFocus()
''    dtc_codigo5.Text = dtc_aux1.Text
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
''    Call pnivel5(dtc_codigo5.BoundText)
''    dtc_desc6.Enabled = True
'End Sub
'
'Private Sub dtc_desc3_Click(Area As Integer)
'    dtc_codigo3.BoundText = dtc_desc3.BoundText
'End Sub

'Private Sub OptFilGral1_Click()
'    '===== Proceso para filtrado general de datos (registros no aprobados)
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    'queryinicial = "select * From to_cronograma_mensual WHERE estado_codigo = 'REG' AND unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '" & VAR_GES & "' "
'    queryinicial = "select * From to_cronograma_mensual WHERE estado_codigo = 'REG' AND unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '2015' "
'    'queryinicial = "Select * from ao_solicitud where " + parametro
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    Set dg_datos.DataSource = Ado_datos.Recordset
'End Sub

Private Sub OptFilGral2_Click()
    '===== Proceso para filtrado general de datos (todos los registros)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "select * From av_ventas_cabecera "
    'queryinicial = "Select * from to_cronograma_mensual where  unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '" & VAR_GES & "' "
    queryinicial = "Select * from tc_zonas_piloto  "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub ABRIR_TABLA()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "Select * from ao_solicitud_cotiza_venta where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset

'    dtc_desc31.BoundText = dtc_codigo31.BoundText
'    dtc_desc32.BoundText = dtc_codigo31.BoundText
'    dtc_desc33.BoundText = dtc_codigo31.BoundText
'    dtc_desc34.BoundText = dtc_codigo31.BoundText
'
'    dtc_desc41.BoundText = dtc_codigo41.BoundText
'    dtc_desc42.BoundText = dtc_codigo41.BoundText
'    dtc_desc43.BoundText = dtc_codigo41.BoundText
'    dtc_desc44.BoundText = dtc_codigo41.BoundText
'
'    dtc_desc51.BoundText = dtc_codigo51.BoundText
'    dtc_desc52.BoundText = dtc_codigo51.BoundText
'    dtc_desc53.BoundText = dtc_codigo51.BoundText
'    dtc_desc54.BoundText = dtc_codigo51.BoundText
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

Private Sub ABRIR_TABLA_DET()
'    Set rs_det1 = New ADODB.Recordset
'    If rs_det1.State = 1 Then rs_det1.Close
'    rs_det1.Open "select * from to_cronograma_diario where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    Set Ado_detalle1.Recordset = rs_det1
'    Set dg_det1.DataSource = Ado_detalle1.Recordset
End Sub


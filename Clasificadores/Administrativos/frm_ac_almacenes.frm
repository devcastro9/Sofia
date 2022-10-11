VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_ac_almacenes 
   BackColor       =   &H00000000&
   Caption         =   "Clasificadores - Almacenes Físicos"
   ClientHeight    =   9495
   ClientLeft      =   495
   ClientTop       =   1905
   ClientWidth     =   15120
   Icon            =   "frm_ac_almacenes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "frm_ac_almacenes.frx":0A02
      ScaleHeight     =   960
      ScaleWidth      =   14880
      TabIndex        =   47
      Top             =   120
      Width           =   14940
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "frm_ac_almacenes.frx":6CA34
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   6000
         Picture         =   "frm_ac_almacenes.frx":6CC3E
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "frm_ac_almacenes.frx":6CE48
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "frm_ac_almacenes.frx":6D46C
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "frm_ac_almacenes.frx":6DA4C
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4320
         Picture         =   "frm_ac_almacenes.frx":6E716
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "frm_ac_almacenes.frx":6ECD3
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "frm_ac_almacenes.frx":6F28B
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   720
         Left            =   5160
         Picture         =   "frm_ac_almacenes.frx":6F495
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   120
         Visible         =   0   'False
         Width           =   765
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
         Left            =   10200
         TabIndex        =   57
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_ac_almacenes.frx":6FE97
      ScaleHeight     =   915
      ScaleWidth      =   14880
      TabIndex        =   43
      Top             =   120
      Width           =   14940
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "frm_ac_almacenes.frx":DBEC9
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "frm_ac_almacenes.frx":DC0D3
         Style           =   1  'Graphical
         TabIndex        =   44
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
         Left            =   9840
         TabIndex        =   46
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.PictureBox Fra_aux1 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1260
      Left            =   6480
      ScaleHeight     =   1200
      ScaleWidth      =   8520
      TabIndex        =   36
      Top             =   7560
      Width           =   8580
      Begin VB.CommandButton CmdCancelaDet 
         BackColor       =   &H80000018&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   6720
         MaskColor       =   &H00000000&
         Picture         =   "frm_ac_almacenes.frx":DC2DD
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Cancelar"
         Top             =   360
         Width           =   765
      End
      Begin VB.CommandButton CmdGrabaDet 
         BackColor       =   &H80000018&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   7575
         Picture         =   "frm_ac_almacenes.frx":DC4E7
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   360
         Width           =   765
      End
      Begin VB.TextBox Txt_descripcion11 
         DataField       =   "calle_denominacion"
         Height          =   645
         Left            =   2160
         TabIndex        =   38
         Text            =   "-"
         Top             =   360
         Width           =   4335
      End
      Begin VB.ComboBox dtc_codigo11 
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Text            =   "CALLE"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lbl_enlace11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Vía de Acceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   2070
      End
      Begin VB.Label lbl_descripcion11 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación Via de Acceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Left            =   2400
         TabIndex        =   39
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFC0&
      Height          =   6255
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   6255
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   5775
         Width           =   5985
         _ExtentX        =   10557
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
         Caption         =   " <-- Inicio                             Navegar                                  Fin -->"
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
         Bindings        =   "frm_ac_almacenes.frx":DC6F1
         Height          =   5415
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   9551
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
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
            DataField       =   "almacen_codigo"
            Caption         =   "Código"
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
            DataField       =   "almacen_descripcion"
            Caption         =   "Descripción"
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
         BeginProperty Column03 
            DataField       =   "beneficiario_codigo"
            Caption         =   "Responsable"
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
            DataField       =   "edif_codigo"
            Caption         =   "Codigo Edificio"
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
               ColumnWidth     =   4140.284
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox fraDatos 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   6255
      Left            =   6480
      ScaleHeight     =   6195
      ScaleWidth      =   8565
      TabIndex        =   4
      Top             =   1320
      Width           =   8625
      Begin VB.CommandButton BtnAux1 
         BackColor       =   &H00808000&
         Caption         =   "Nueva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7560
         Picture         =   "frm_ac_almacenes.frx":DC709
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   5040
         Width           =   660
      End
      Begin VB.TextBox Txt_descripcion 
         DataField       =   "almacen_descripcion"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   240
         TabIndex        =   62
         Top             =   880
         Width           =   8040
      End
      Begin VB.TextBox txt_campo1 
         DataField       =   "almacen_telefonos"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3000
         TabIndex        =   35
         Top             =   5160
         Width           =   3975
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Localización"
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
         Height          =   1815
         Left            =   60
         TabIndex        =   22
         Top             =   1800
         Width           =   8430
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "frm_ac_almacenes.frx":DCEBA
            DataField       =   "pais_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   23
            Top             =   435
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "pais_codigo"
            BoundColumn     =   "pais_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_aux7 
            Bindings        =   "frm_ac_almacenes.frx":DCED3
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6120
            TabIndex        =   24
            Top             =   1035
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "correl_edif"
            BoundColumn     =   "munic_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo7 
            Bindings        =   "frm_ac_almacenes.frx":DCEEC
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7080
            TabIndex        =   25
            Top             =   1035
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
         Begin MSDataListLib.DataCombo dtc_desc7 
            Bindings        =   "frm_ac_almacenes.frx":DCF05
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4200
            TabIndex        =   26
            Top             =   1300
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "munic_descripcion"
            BoundColumn     =   "munic_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo6 
            Bindings        =   "frm_ac_almacenes.frx":DCF1E
            DataField       =   "prov_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   27
            Top             =   1035
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
         Begin MSDataListLib.DataCombo dtc_desc6 
            Bindings        =   "frm_ac_almacenes.frx":DCF37
            DataField       =   "prov_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   1300
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "prov_descripcion"
            BoundColumn     =   "prov_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "frm_ac_almacenes.frx":DCF50
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7080
            TabIndex        =   29
            Top             =   435
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "depto_codigo"
            BoundColumn     =   "depto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "frm_ac_almacenes.frx":DCF69
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4200
            TabIndex        =   30
            Top             =   600
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "depto_descripcion"
            BoundColumn     =   "depto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "frm_ac_almacenes.frx":DCF82
            DataField       =   "pais_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "pais_descripcion"
            BoundColumn     =   "pais_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label lbl_campo7 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Municipio"
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
            Left            =   4200
            TabIndex        =   65
            Top             =   1035
            Width           =   855
         End
         Begin VB.Label lbl_campo4 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "País"
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
            TabIndex        =   59
            Top             =   320
            Width           =   405
         End
         Begin VB.Label lbl_campo5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Departamento"
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
            Left            =   4200
            TabIndex        =   33
            Top             =   320
            Width           =   1290
         End
         Begin VB.Label lbl_campo6 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Provincia"
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
            TabIndex        =   32
            Top             =   1035
            Width           =   840
         End
      End
      Begin VB.TextBox txt_campo3 
         DataField       =   "almacen_piso_nro"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5760
         TabIndex        =   3
         Top             =   4695
         Width           =   1275
      End
      Begin VB.TextBox txt_campo4 
         BackColor       =   &H00FFFFFF&
         DataField       =   "almacen_depto_nro"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7160
         TabIndex        =   0
         Top             =   4695
         Width           =   1260
      End
      Begin VB.TextBox txt_campo5 
         DataField       =   "almacen_obs"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "-"
         Top             =   5780
         Width           =   8280
      End
      Begin VB.TextBox txt_campo2 
         DataField       =   "almacen_edif_nro"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4320
         TabIndex        =   2
         Top             =   4695
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "frm_ac_almacenes.frx":DCF9B
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2730
         TabIndex        =   9
         Top             =   1380
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "frm_ac_almacenes.frx":DCFB4
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4800
         TabIndex        =   10
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "frm_ac_almacenes.frx":DCFCD
         DataField       =   "zona_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3000
         TabIndex        =   13
         Top             =   3720
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "zona_codigo"
         BoundColumn     =   "zona_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "frm_ac_almacenes.frx":DCFE6
         DataField       =   "zona_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   3960
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "zona_denominacion"
         BoundColumn     =   "zona_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo9 
         Bindings        =   "frm_ac_almacenes.frx":DCFFF
         DataField       =   "calle_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7200
         TabIndex        =   15
         Top             =   3600
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "calle_codigo"
         BoundColumn     =   "calle_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc9 
         Bindings        =   "frm_ac_almacenes.frx":DD018
         DataField       =   "calle_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4260
         TabIndex        =   16
         Top             =   3960
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "calle_denominacion"
         BoundColumn     =   "calle_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "frm_ac_almacenes.frx":DD031
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2940
         TabIndex        =   17
         Top             =   4440
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "edif_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "frm_ac_almacenes.frx":DD04B
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   4695
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux4 
         Bindings        =   "frm_ac_almacenes.frx":DD065
         DataField       =   "pais_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2280
         TabIndex        =   34
         Top             =   5160
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "pais_cod_telefonico"
         BoundColumn     =   "pais_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lbl_campo15 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Observaciones"
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
         TabIndex        =   64
         Top             =   5520
         Width           =   1380
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Responsable del Almacen:"
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
         Left            =   240
         TabIndex        =   63
         Top             =   1440
         Width           =   2445
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Denominación del Almacen"
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
         Left            =   240
         TabIndex        =   61
         Top             =   600
         Width           =   2475
      End
      Begin VB.Label lbl_titulo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Código Almacen"
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
         Left            =   240
         TabIndex        =   60
         Top             =   200
         Width           =   1500
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "REG"
         DataField       =   "almacen_codigo"
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
         Height          =   315
         Left            =   1920
         TabIndex        =   58
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label lbl_calle 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Via de Acceso (Calle, Av, etc.)"
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
         Left            =   4320
         TabIndex        =   42
         Top             =   3705
         Width           =   2685
      End
      Begin VB.Label lbl_zona 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Zona / Barrio"
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
         TabIndex        =   41
         Top             =   3705
         Width           =   1155
      End
      Begin VB.Label lbl_campo14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro. Depto."
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
         Left            =   7260
         TabIndex        =   21
         Top             =   4440
         Width           =   1020
      End
      Begin VB.Label lbl_campo3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro. Piso"
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
         Left            =   5925
         TabIndex        =   20
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label lbl_campo10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Vivienda / Edificación"
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
         TabIndex        =   19
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Estado Registro"
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
         Index           =   9
         Left            =   5280
         TabIndex        =   8
         Top             =   200
         Width           =   1455
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
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
         Height          =   315
         Left            =   6840
         TabIndex        =   7
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label lbl_campo11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Teléfono del Almacen"
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
         TabIndex        =   6
         Top             =   5160
         Width           =   1980
      End
      Begin VB.Label lbl_campo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro. Vivienda"
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
         Left            =   4305
         TabIndex        =   5
         Top             =   4440
         Width           =   1245
      End
   End
   Begin Crystal.CrystalReport CR01 
      Left            =   720
      Top             =   9000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   10800
      Top             =   8400
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
   Begin MSAdodcLib.Adodc Ado_datos7 
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
   Begin MSAdodcLib.Adodc Ado_datos2 
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
   Begin MSAdodcLib.Adodc Ado_datos8 
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   2160
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
   Begin MSAdodcLib.Adodc Ado_datos4 
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
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   4320
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
   Begin MSAdodcLib.Adodc Ado_datos5 
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
End
Attribute VB_Name = "frm_ac_almacenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mantenimiento de Beneficiarios
Option Explicit
Dim rs_datos As New ADODB.Recordset
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

Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

'OTROS
Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim SQL_FOR As String
Dim RUTA1 As String
Dim VAR_COD2 As Double
'Dim SW As Boolean
'Dim CORREL, accion As Integer
'Dim swnuevo As Boolean
Dim CodBenef As String
Dim sino As String       ', NombreCarpeta, e

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  If Ado_datos.Recordset.EOF Or Ado_datos.Recordset.BOF Then
'      BtnModificar.Enabled = False
'     ' BtnEliminar.Enabled = False
'      'TxtTipo.Text = Empty
'      txtCodigo.Text = Empty
'      Text1.Text = Empty
'      Text2.Text = Empty
'      Text3.Text = Empty
'      txtDenominacion.Text = Empty
'      Exit Sub
'  End If
  If Ado_datos.Recordset.RecordCount > 0 Then
'    Select Case Ado_datos.Recordset.EditMode
'      Case adEditInProgress
'        Frame2.Enabled = False            'Verif. Nombre Proveedor JQA NOV-2009
'
'      Case adEditNone
'      Case adEditDelete
'      Case adEditAdd
'        Frame2.Enabled = True            'Verif. Nombre Proveedor JQA NOV-2009
'    End Select

'    If VAR_SW = "ADD" Or VAR_SW = "MOD" Then
'        txt_campo4.Visible = True
'        txt_codigo.Visible = False
'    Else
'        txt_campo4.Visible = False
'        txt_codigo.Visible = True
'    End If
    Ado_datos.Caption = CStr(Ado_datos.Recordset.AbsolutePosition) & " de " & CStr(Ado_datos.Recordset.RecordCount)
  End If
End Sub
   
Private Sub BtnAux1_Click()
    'Validacion 1
    If dtc_codigo8 = "" Or dtc_codigo8 = "0" Then
        MsgBox "Debe registrar: " + lbl_zona.Caption, vbCritical + vbExclamation, "Validación de datos"
        VAR_VAL = "ERR"
        Exit Sub
    End If
    fraDatos.Enabled = False
    Fra_aux1.Visible = True
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

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
     If VAR_SW = "ADD" Then
       var_cod = txt_codigo
       Set rs_aux1 = New ADODB.Recordset
       SQL_FOR = "select * from ac_almacenes where almacen_codigo = '" & var_cod & "'  "
       rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic        ', adCmdText
       If rs_aux1.RecordCount > 0 Then
'                SW = True
                MsgBox " CODIGO DUPLICADO"
                Txt_descripcion.SetFocus
                Exit Sub
       End If
       rs_datos!estado_codigo = "REG"
       rs_datos!CORREL = 0
        'rs_datos!ARCHIVO_Foto = txt_codigo.Caption + ".JPG"
        'rs_datos!archivo_foto_cargado = "N"
        'rs_datos!ges_gestion = Year(Date)
        'db.Execute "Update gc_municipio Set correl_edif = CAST('" & dtc_aux2.Text & "' AS INT) + 1 Where munic_codigo= '" & dtc_codigo2.Text & "' "
     End If
     rs_datos!almacen_descripcion = Txt_descripcion.Text
     
     rs_datos!beneficiario_codigo = dtc_codigo1.Text
     rs_datos!pais_codigo = IIf(dtc_codigo4.Text = "", "BOL", dtc_codigo4.Text)
     rs_datos!depto_codigo = dtc_codigo5.Text
     rs_datos!prov_codigo = dtc_codigo6.Text
     rs_datos!munic_codigo = dtc_codigo7.Text
     rs_datos!zona_codigo = IIf(dtc_codigo8.Text = "", "0", dtc_codigo8.Text)
     rs_datos!calle_codigo = IIf(dtc_codigo9.Text = "", "0", dtc_codigo9.Text)
     rs_datos!edif_codigo = IIf(dtc_codigo10.Text = "", "10101-0", dtc_codigo10.Text)

     rs_datos!almacen_edif_nro = IIf(Txt_campo2.Text = "", "0", Txt_campo2.Text)
     rs_datos!almacen_piso_nro = IIf(txt_campo3.Text = "", "0", txt_campo3.Text)
     rs_datos!almacen_depto_nro = IIf(txt_campo4.Text = "", "0", txt_campo4.Text)
     rs_datos!almacen_telefonos = txt_campo1.Text
     rs_datos!almacen_obs = txt_campo5.Text
     
'     If rs_datos!ARCHIVO_Foto = ".JPG" Or rs_datos!ARCHIVO_Foto = "" Then
'        rs_datos!ARCHIVO_Foto = txt_codigo.Caption + ".JPG"
'     End If
     
     rs_datos!fecha_registro = Date
     rs_datos!usr_codigo = glusuario
     rs_datos.UpdateBatch adAffectAll
    
     Call ABRIR_TABLA
     rs_datos.MoveLast
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     fraDatos.Enabled = False
     dg_datos.Enabled = True
     txt_campo4.Enabled = True

  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
    
End Sub

Private Sub valida_campos()
  If (txt_campo4.Text = "") Then
    MsgBox "Debe registrar el " + lbl_titulo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If txt_campo2.Text = "" Then
'    MsgBox "Debe registrar la " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If txt_campo3.Text = "" Then
'    MsgBox "Debe registrar la " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If dtc_codigo2.Text = "" Then
'    MsgBox "Debe registrar la " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If dtc_codigo3.Text = "" Then
'    MsgBox "Debe registrar la " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
End Sub
 
Private Sub BtnAñadir_Click()
  On Error GoTo AddErr
    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    rs_datos.AddNew
    'lblStatus.Caption = "Agregar registro"
    fraDatos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "ADD"
'    txt_campo4.SetFocus
'    BtnVer.Visible = False
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If rs_datos!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
                  
         rs_datos!estado_codigo = "APR"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnEliminar_Click()
   If ExisteBenef(Ado_datos.Recordset!beneficiario_codigo) Then MsgBox "No se puede ANULAR un Beneficiario que ya fue utilizado ...", vbInformation + vbOKOnly, "Atención": Exit Sub
   If ExisteBenef2(Ado_datos.Recordset!beneficiario_codigo) Then MsgBox "No se puede ANULAR un Beneficiario que ya fue utilizado ...", vbInformation + vbOKOnly, "Atención": Exit Sub
   sino = MsgBox("Está Seguro de ANULAR el Registro?", vbYesNo + vbQuestion, "Atención")
   If Ado_datos.Recordset("estado_codigo") = "APR" Then
      If sino = vbYes Then
        Ado_datos.Recordset("estado_codigo") = "ERR"
        Ado_datos.Recordset("fecha_registro") = Date
        Ado_datos.Recordset("usr_codigo") = glusuario
        Ado_datos.Recordset.Update  'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR un registro Elaborado (REG) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelUpdate
        Call ABRIR_TABLAS_AUX
        Call ABRIR_TABLA
        rs_datos.MoveFirst
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        fraDatos.Enabled = False
        dg_datos.Enabled = True
'        txt_campo4.Enabled = True
    End If
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
  If Ado_datos.Recordset!estado_codigo = "REG" Then
'  lblStatus.Caption = "Modificar registro"
    fraDatos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
'    txt_campo4.Enabled = False
    VAR_SW = "MOD"
'    BtnVer.Visible = True
  Else
     MsgBox "No se puede MODIFICAR un registro Aprobado (APR) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
  End If
  Exit Sub

EditErr:
  MsgBox Err.Description

End Sub

Private Sub BtnVer_Click()
   On Error GoTo QError
    If Ado_datos.Recordset!ARCHIVO_Foto = "Cargar_Archivo" Then
      NombreCarpeta = App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "FOT"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(ado_datos.Recordset!iniciales) & "-" & Trim(ado_datos.Recordset!codigo_beneficiario) & "\"
'      Else
         e = NombreCarpeta
'      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
    Else
      'MsgBox ""
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "FOT"
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(ado_datos.Recordset!iniciales) & "-" & Trim(ado_datos.Recordset!codigo_beneficiario) & "\"
'          Else
            e = NombreCarpeta
'          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
      End If
    End If

    Dim ARCH_FOTO As String
'    If GlServidor = "SRVPRO" Then
'        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(ado_datos.Recordset!iniciales) + "-" + Trim(ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(ado_datos.Recordset!ARCHIVO_FOTO)
'    Else
        ARCH_FOTO = App.Path + "\PERSONAL\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_Foto)
'    End If
    'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + ado_datos.Recordset!codigo_beneficiario + "\" + ado_datos.Recordset("codigo_beneficiario") + "-FOTO.JPG"
    CodBenef = Ado_datos.Recordset!codigo_beneficiario
    If Guardar_Imagen(db, "Select Foto From fc_beneficiario Where codigo_beneficiario= '" & CodBenef & "' ", "Foto", ARCH_FOTO) Then
        MsgBox "Se cargo la Imagen Correctamente !!"
        Exit Sub
    Else
        MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    End If
QError:
    ' Manejo de errores
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
'    db.RollbackTrans
    Screen.MousePointer = vbDefault
End Sub

Private Sub BtnImprimir_Click()
  Dim iResult As Integer
     CR01.WindowShowPrintSetupBtn = True
     CR01.WindowShowRefreshBtn = True
     CR01.ReportFileName = App.Path & "\REPORTES\clasificadores\gr_beneficiario_Persona.rpt"
  iResult = CR01.PrintReport
  If iResult <> 0 Then
      MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If

CR01.WindowState = crptMaximized
End Sub

Private Sub BtnSalir_Click()
'  If glPersNew = "CMP" Then
'    frmComprasDirectas.DtcProv = Ado_datos.Recordset("codigo_beneficiario")
'    frmComprasDirectas.cboListaProv2 = Ado_datos.Recordset("denominacion_Beneficiario")
'    frmComprasDirectas.txtProv = Ado_datos.Recordset("codigo_beneficiario")
'    frmComprasDirectas.txtProveedor = Ado_datos.Recordset("denominacion_Beneficiario")
'  End If
  Unload Me
End Sub

Private Sub CmdCancelaDet_Click()
    fraDatos.Enabled = True
    Fra_aux1.Visible = False
End Sub

Private Sub CmdGrabaDet_Click()
  'Validacion
  If Txt_descripcion11.Text = "" Then
    MsgBox "Debe registrar la " + lbl_descripcion11.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo11.Text = "" Then
    MsgBox "Debe registrar: " + lbl_enlace11.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo8 = "" Or dtc_codigo8 = "0" Then
    MsgBox "Debe registrar: " + lbl_zona.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  'INI Graba Calle
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "Select max(calle_codigo) as Codigo from gc_calles where zona_codigo = " & dtc_codigo8.Text & "    ", db, adOpenStatic
    If rs_aux2.RecordCount > 0 Then
        VAR_COD2 = Round(CDbl(rs_aux2!Codigo) + 1, 0)
    Else
        VAR_COD2 = 1
    End If
    db.Execute "insert into gc_calles(zona_codigo, calle_codigo, calle_denominacion, calle_tipo, correl, estado_codigo, fecha_registro, usr_codigo)" & _
    "values ('" & dtc_codigo8.Text & "', " & VAR_COD2 & ", '" & Txt_descripcion11.Text & "', '" & dtc_codigo11.Text & "', '0', 'APR', '" & Date & "', '" & glusuario & "') "
    
   'FIN Graba Calle
    'Guarda en el Padre, en el campo ctrl de correlativos para codigos que se generan
    db.Execute "Update gc_zonas Set correl = " & VAR_COD2 & " Where zona_codigo= '" & dtc_codigo8.Text & "' "
    'gc_calles
    Call pnivel6(dtc_codigo8.BoundText)
    dtc_desc9.Enabled = True
    
    dtc_codigo9.Text = VAR_COD2
    dtc_desc9.BoundText = dtc_codigo9.BoundText
    
    fraDatos.Enabled = True
    Fra_aux1.Visible = False

End Sub

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_aux4.BoundText
    dtc_codigo4.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_codigo7_Click(Area As Integer)
    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc7.BoundText
    Call pnivel5(dtc_codigo7.BoundText)
    dtc_desc8.Enabled = True
    Call pnivel7(dtc_codigo7.BoundText)
    dtc_desc10.Enabled = True
End Sub
   
Private Sub pnivel5(codigo7 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_zonas where munic_codigo = '" & codigo7 & "'"
   Set dtc_codigo8.RowSource = Nothing
   Set dtc_codigo8.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo8.ReFill
   dtc_codigo8.BoundText = Empty
   
   Set dtc_desc8.RowSource = Nothing
   Set dtc_desc8.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc8.ReFill
   dtc_desc8.BoundText = Empty
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
    Call pnivel6(dtc_codigo8.BoundText)
    dtc_desc9.Enabled = True
End Sub

Private Sub pnivel6(codigo8 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_calles where zona_codigo = '" & codigo8 & "'"
   Set dtc_codigo9.RowSource = Nothing
   Set dtc_codigo9.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo9.ReFill
   dtc_codigo9.BoundText = Empty
   
   Set dtc_desc9.RowSource = Nothing
   Set dtc_desc9.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc9.ReFill
   dtc_desc9.BoundText = Empty
End Sub

Private Sub dtc_desc9_Click(Area As Integer)
    dtc_codigo9.BoundText = dtc_desc9.BoundText
End Sub

Private Sub pnivel7(codigo9 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_edificaciones where munic_codigo = '" & codigo9 & "'"
   Set dtc_codigo10.RowSource = Nothing
   Set dtc_codigo10.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo10.ReFill
   dtc_codigo10.BoundText = Empty
   
   Set dtc_desc10.RowSource = Nothing
   Set dtc_desc10.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc10.ReFill
   dtc_desc10.BoundText = Empty
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLAS_AUX
    Call ABRIR_TABLA
    'txt_codigo.Enabled = True
'    mbDataChanged = False
    fraDatos.Enabled = False
    dg_datos.Enabled = True
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
    Fra_aux1.Visible = False
End Sub

Private Sub ABRIR_TABLA()
   Set rs_datos = New ADODB.Recordset
   If rs_datos.State = 1 Then rs_datos.Close
   queryinicial = "select * from ac_almacenes "     'WHERE beneficiario_codigo <> ' ' and beneficiario_codigo <> '0' and tipoben_codigo > 19 "
   'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
   rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
   rs_datos.Sort = "almacen_descripcion"
   Set Ado_datos.Recordset = rs_datos
End Sub

Private Sub Form_Unload(Cancel As Integer)
'  If glPersNew = "P" Then
'     FrmVentas.DtcNIT = Ado_datos.Recordset("codigo_beneficiario")
'     FrmVentas.DtcdesNIT = Ado_datos.Recordset("denominacion_Beneficiario")
'  End If
'
'  glPersNew = "N"
   
   If (rs_datos.State = adStateClosed) Then rs_datos.Close
   'Set rs_datos = Nothing
End Sub

Private Sub ABRIR_TABLAS_AUX()
    ' gc_beneficiario (Personal CGI)
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "select * from rv_unidad_vs_responsable ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText

    'gc_pais
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from gc_pais  where estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    'gc_Departamento  '<>
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from gc_departamento order by depto_descripcion", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    'gc_provincia
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from gc_provincia ", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    'gc_municipio
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from gc_municipio ", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
    'gc_zonas
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_zonas ", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    'gc_calles
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_calles ", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_desc9.BoundText = dtc_codigo9.BoundText
    'gc_edificaciones
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from gc_edificaciones ", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText
    
End Sub

Private Sub txt_campo1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function ExisteBenef(CodBenef As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE beneficiario_codigo_resp = '" & CodBenef & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteBenef = rs!Cuantos > 0
End Function

Private Function ExisteBenef2(CodBenef As String) As Boolean
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE beneficiario_codigo = '" & CodBenef & "'"
    rs2.Open GlSqlAux, db, adOpenStatic
    ExisteBenef2 = rs2!Cuantos > 0
End Function


Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
    dtc_aux4.BoundText = dtc_desc4.BoundText
    Call pnivel2(dtc_codigo4.BoundText)
    dtc_desc5.Enabled = True
End Sub
   
Private Sub pnivel2(codigo4 As String)
   Dim strConsultaF As String
     
   strConsultaF = "select * from gc_departamento where pais_codigo = '" & codigo4 & "'"
   Set dtc_codigo5.RowSource = Nothing
   Set dtc_codigo5.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo5.ReFill
   dtc_codigo5.BoundText = Empty
   
   Set dtc_desc5.RowSource = Nothing
   Set dtc_desc5.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc5.ReFill
   dtc_desc5.BoundText = Empty

End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    Call pnivel3(dtc_codigo5.BoundText)
    dtc_desc6.Enabled = True
End Sub
   
Private Sub pnivel3(codigo5 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_provincia where depto_codigo = '" & codigo5 & "'"
   Set dtc_codigo6.RowSource = Nothing
   Set dtc_codigo6.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo6.ReFill
   dtc_codigo6.BoundText = Empty
   
   Set dtc_desc6.RowSource = Nothing
   Set dtc_desc6.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc6.ReFill
   dtc_desc6.BoundText = Empty
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
    Call pnivel4(dtc_codigo6.BoundText)
    dtc_desc7.Enabled = True
End Sub
   
Private Sub pnivel4(codigo6 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_municipio where prov_codigo = '" & codigo6 & "'"
   Set dtc_codigo7.RowSource = Nothing
   Set dtc_codigo7.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo7.ReFill
   dtc_codigo7.BoundText = Empty
   
   Set dtc_desc7.RowSource = Nothing
   Set dtc_desc7.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc7.ReFill
   dtc_desc7.BoundText = Empty
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

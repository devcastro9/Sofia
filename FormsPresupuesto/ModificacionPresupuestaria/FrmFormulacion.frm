VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmFormulacion 
   Caption         =   "Formulación Presupuestaria"
   ClientHeight    =   8970
   ClientLeft      =   2010
   ClientTop       =   915
   ClientWidth     =   12090
   Icon            =   "FrmFormulacion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   12090
   Begin TabDlg.SSTab sstab1 
      Height          =   7815
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13785
      _Version        =   393216
      TabOrientation  =   2
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   0
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "FORMULACION"
      TabPicture(0)   =   "FrmFormulacion.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "TRANSACCIONES"
      TabPicture(1)   =   "FrmFormulacion.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   7680
         Left            =   380
         TabIndex        =   1
         Top             =   60
         Width           =   11320
         Begin VB.PictureBox fraOpciones 
            BackColor       =   &H00404040&
            Height          =   1020
            Left            =   120
            Picture         =   "FrmFormulacion.frx":047A
            ScaleHeight     =   960
            ScaleWidth      =   11040
            TabIndex        =   155
            Top             =   120
            Width           =   11100
            Begin VB.CommandButton BtnAprobar 
               BackColor       =   &H00808000&
               Caption         =   "Aprobar"
               Height          =   720
               Left            =   2640
               Picture         =   "FrmFormulacion.frx":6C4AC
               Style           =   1  'Graphical
               TabIndex        =   156
               ToolTipText     =   "Aprueba Registro"
               Top             =   120
               Width           =   765
            End
            Begin VB.CommandButton BtnVer 
               BackColor       =   &H00808000&
               Caption         =   "Digitaliza"
               Height          =   720
               Left            =   5160
               Picture         =   "FrmFormulacion.frx":6C6B6
               Style           =   1  'Graphical
               TabIndex        =   164
               ToolTipText     =   "Guarda en Archivo Digital"
               Top             =   120
               Width           =   765
            End
            Begin VB.CommandButton BtnDesAprobar 
               BackColor       =   &H00808000&
               Caption         =   "Desapro."
               Height          =   720
               Left            =   2640
               Picture         =   "FrmFormulacion.frx":6CAF8
               Style           =   1  'Graphical
               TabIndex        =   163
               Top             =   120
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.CommandButton BtnBuscar 
               BackColor       =   &H00808000&
               Caption         =   "Buscar"
               Height          =   720
               Left            =   3480
               Picture         =   "FrmFormulacion.frx":6CD02
               Style           =   1  'Graphical
               TabIndex        =   162
               ToolTipText     =   "Busca un Registro"
               Top             =   120
               Width           =   765
            End
            Begin VB.CommandButton BtnImprimir 
               BackColor       =   &H00808000&
               Caption         =   "Imprimir"
               Height          =   720
               Left            =   4320
               Picture         =   "FrmFormulacion.frx":6D2BA
               Style           =   1  'Graphical
               TabIndex        =   161
               ToolTipText     =   "Imprime Formulario"
               Top             =   120
               Width           =   765
            End
            Begin VB.CommandButton BtnSalir 
               BackColor       =   &H00808000&
               Caption         =   "Cerrar"
               Height          =   720
               Left            =   6000
               Picture         =   "FrmFormulacion.frx":6D877
               Style           =   1  'Graphical
               TabIndex        =   160
               ToolTipText     =   "Cerrar Ventana"
               Top             =   120
               Width           =   765
            End
            Begin VB.CommandButton BtnEliminar 
               BackColor       =   &H00808000&
               Caption         =   "Anular"
               Height          =   720
               Left            =   1800
               Picture         =   "FrmFormulacion.frx":6DA81
               Style           =   1  'Graphical
               TabIndex        =   159
               ToolTipText     =   "Anula Registro Activo"
               Top             =   120
               Width           =   765
            End
            Begin VB.CommandButton BtnModificar 
               BackColor       =   &H00808000&
               Caption         =   "Modificar"
               Height          =   720
               Left            =   960
               Picture         =   "FrmFormulacion.frx":6E74B
               Style           =   1  'Graphical
               TabIndex        =   158
               ToolTipText     =   "Modifica Registro Activo"
               Top             =   120
               Width           =   765
            End
            Begin VB.CommandButton BtnAńadir 
               BackColor       =   &H00808000&
               Caption         =   "Nuevo"
               Height          =   720
               Left            =   120
               Picture         =   "FrmFormulacion.frx":6ED2B
               Style           =   1  'Graphical
               TabIndex        =   157
               ToolTipText     =   "Nuevo Registro"
               Top             =   120
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
               Left            =   8280
               TabIndex        =   165
               Top             =   300
               Width           =   1305
            End
         End
         Begin VB.PictureBox FraGrabarCancelar 
            BackColor       =   &H00000000&
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   120
            Picture         =   "FrmFormulacion.frx":6F34F
            ScaleHeight     =   915
            ScaleWidth      =   11040
            TabIndex        =   172
            Top             =   120
            Width           =   11100
            Begin VB.CommandButton BtnCancelar 
               BackColor       =   &H00808000&
               Caption         =   "Cancelar"
               Height          =   675
               Left            =   3600
               MaskColor       =   &H00000000&
               Picture         =   "FrmFormulacion.frx":DB381
               Style           =   1  'Graphical
               TabIndex        =   174
               ToolTipText     =   "Cancelar"
               Top             =   120
               Width           =   765
            End
            Begin VB.CommandButton BtnGrabar 
               BackColor       =   &H00808000&
               Caption         =   "Grabar"
               Height          =   675
               Left            =   1560
               Picture         =   "FrmFormulacion.frx":DB58B
               Style           =   1  'Graphical
               TabIndex        =   173
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
               Left            =   8220
               TabIndex        =   175
               Top             =   300
               Width           =   1305
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2280
            TabIndex        =   166
            Top             =   4425
            Width           =   6615
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Totales :"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   255
               Left            =   240
               TabIndex        =   171
               Top             =   25
               Width           =   975
            End
            Begin VB.Label lblFormulado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "0"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1380
               TabIndex        =   170
               Top             =   25
               Width           =   1240
            End
            Begin VB.Label lblAdiciones 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "0"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2715
               TabIndex        =   169
               Top             =   25
               Width           =   1200
            End
            Begin VB.Label lblModificaciones 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "0"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3980
               TabIndex        =   168
               Top             =   25
               Width           =   1215
            End
            Begin VB.Label lblVigente 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "0"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5280
               TabIndex        =   167
               Top             =   25
               Width           =   1240
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00000000&
            Caption         =   "REGISTRO"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   2655
            Left            =   120
            TabIndex        =   2
            Top             =   4800
            Width           =   11100
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               DataField       =   "fgs_vigente"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adoformulacion"
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
               Left            =   8760
               TabIndex        =   56
               Text            =   "0"
               Top             =   2160
               Width           =   2055
            End
            Begin VB.TextBox Text3 
               Alignment       =   2  'Center
               DataField       =   "fgs_adiciones"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adoformulacion"
               Height          =   285
               Left            =   5880
               TabIndex        =   55
               Text            =   "0"
               Top             =   2160
               Width           =   2055
            End
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               DataField       =   "fgs_modificaciones"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adoformulacion"
               Height          =   285
               Left            =   3000
               TabIndex        =   54
               Text            =   "0"
               Top             =   2160
               Width           =   2055
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               DataField       =   "fgs_formulado"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adoformulacion"
               Height          =   285
               Left            =   120
               TabIndex        =   53
               Text            =   "0"
               Top             =   2160
               Width           =   2055
            End
            Begin MSAdodcLib.Adodc Adofuente 
               Height          =   375
               Left            =   8520
               Top             =   360
               Visible         =   0   'False
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   661
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
               Caption         =   "Fuente"
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
            Begin MSAdodcLib.Adodc adoorganismo 
               Height          =   330
               Left            =   8520
               Top             =   720
               Visible         =   0   'False
               Width           =   1935
               _ExtentX        =   3413
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
               Caption         =   "Organismo"
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
            Begin MSDataListLib.DataCombo dtv_fuente 
               Bindings        =   "FrmFormulacion.frx":DB795
               DataField       =   "fte_codigo"
               DataSource      =   "adoformulacion"
               Height          =   315
               Left            =   2280
               TabIndex        =   3
               Top             =   360
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "fte_codigo"
               BoundColumn     =   "fte_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DTC 
               Bindings        =   "FrmFormulacion.frx":DB7AD
               DataField       =   "fte_codigo"
               DataSource      =   "adoformulacion"
               Height          =   315
               Left            =   3960
               TabIndex        =   4
               Top             =   360
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Fte_descripcion_larga"
               BoundColumn     =   "fte_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSAdodcLib.Adodc adoproyecto 
               Height          =   330
               Left            =   8520
               Top             =   1080
               Visible         =   0   'False
               Width           =   1935
               _ExtentX        =   3413
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
               Caption         =   "proyecto"
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
            Begin MSAdodcLib.Adodc Adopartida 
               Height          =   330
               Left            =   8520
               Top             =   1440
               Visible         =   0   'False
               Width           =   1935
               _ExtentX        =   3413
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
               Caption         =   "partida"
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
            Begin MSDataListLib.DataCombo DataCombo1 
               Bindings        =   "FrmFormulacion.frx":DB7C5
               DataField       =   "org_codigo"
               DataSource      =   "adoformulacion"
               Height          =   315
               Left            =   2280
               TabIndex        =   47
               Top             =   720
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "org_codigo"
               BoundColumn     =   "org_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo2 
               Bindings        =   "FrmFormulacion.frx":DB7E0
               DataField       =   "org_codigo"
               DataSource      =   "adoformulacion"
               Height          =   315
               Left            =   3960
               TabIndex        =   48
               Top             =   720
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "org_descripcion"
               BoundColumn     =   "org_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo4 
               Bindings        =   "FrmFormulacion.frx":DB7FB
               DataField       =   "pro_proyecto"
               DataSource      =   "adoformulacion"
               Height          =   315
               Left            =   3000
               TabIndex        =   49
               Top             =   1080
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "pro_proyecto"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo5 
               Bindings        =   "FrmFormulacion.frx":DB815
               DataField       =   "pro_proyecto"
               DataSource      =   "adoformulacion"
               Height          =   315
               Left            =   4560
               TabIndex        =   50
               Top             =   1080
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Pro_descripcion_larga"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo6 
               Bindings        =   "FrmFormulacion.frx":DB82F
               DataField       =   "par_codigo"
               DataSource      =   "adoformulacion"
               Height          =   315
               Left            =   2280
               TabIndex        =   51
               Top             =   1440
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "par_codigo"
               BoundColumn     =   "par_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo7 
               Bindings        =   "FrmFormulacion.frx":DB848
               DataField       =   "par_codigo"
               DataSource      =   "adoformulacion"
               Height          =   315
               Left            =   3960
               TabIndex        =   52
               Top             =   1440
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Par_descripcion_larga"
               BoundColumn     =   "par_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo8 
               Bindings        =   "FrmFormulacion.frx":DB861
               DataField       =   "pro_proyecto"
               DataSource      =   "adoformulacion"
               Height          =   315
               Left            =   3720
               TabIndex        =   60
               Top             =   1080
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "pro_actividad"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo9 
               Bindings        =   "FrmFormulacion.frx":DB87B
               DataField       =   "pro_proyecto"
               DataSource      =   "adoformulacion"
               Height          =   315
               Left            =   2280
               TabIndex        =   61
               Top             =   1080
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "pro_programa"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin VB.Label Label29 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "="
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF80&
               Height          =   255
               Left            =   8040
               TabIndex        =   59
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label Label28 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF80&
               Height          =   255
               Left            =   5160
               TabIndex        =   58
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF80&
               Height          =   255
               Left            =   2280
               TabIndex        =   57
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "Monto Vigente Bs."
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   8760
               TabIndex        =   46
               Top             =   1900
               Width           =   2055
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "Adiciones o Reducciones Bs."
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   5760
               TabIndex        =   45
               Top             =   1905
               Width           =   2295
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "Traspasos Bs."
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   3000
               TabIndex        =   44
               Top             =   1900
               Width           =   2055
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "Monto Formulado Bs."
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   1900
               Width           =   2055
            End
            Begin VB.Label Label3 
               BackColor       =   &H00000000&
               Caption         =   "Partida del Gasto"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   1440
               Width           =   2055
            End
            Begin VB.Label Label10 
               BackColor       =   &H00000000&
               Caption         =   "Fuente de Financiamiento"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label Label11 
               BackColor       =   &H00000000&
               Caption         =   "Organismo Financiador"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   720
               Width           =   2055
            End
            Begin VB.Label Label12 
               BackColor       =   &H00000000&
               Caption         =   "Categoría Programática"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   5
               Top             =   1080
               Width           =   2055
            End
         End
         Begin MSDataGridLib.DataGrid Dtgformulacion 
            Bindings        =   "FrmFormulacion.frx":DB895
            Height          =   3135
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   11100
            _ExtentX        =   19579
            _ExtentY        =   5530
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777152
            Enabled         =   -1  'True
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   19
            RowDividerStyle =   3
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "FORMULACION PRESUPUESTARIA"
            ColumnCount     =   11
            BeginProperty Column00 
               DataField       =   "fte_codigo"
               Caption         =   "Fte"
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
               DataField       =   "org_codigo"
               Caption         =   "Org"
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
               DataField       =   "pro_programa"
               Caption         =   "Pro"
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
               DataField       =   "pro_proyecto"
               Caption         =   "Pry"
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
               DataField       =   "pro_actividad"
               Caption         =   "Act"
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
               DataField       =   "par_codigo"
               Caption         =   "Partida"
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
               DataField       =   "fgs_formulado"
               Caption         =   "Formulado Bs."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "fgs_adiciones"
               Caption         =   "Add/Red.Bs."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "fgs_modificaciones"
               Caption         =   "Traspasos Bs."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "fgs_vigente"
               Caption         =   "Vigente Bs."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "par_descripcion_larga"
               Caption         =   "      Descripción"
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
                  ColumnWidth     =   510.236
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   540.284
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   464.882
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   480.189
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   464.882
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   750.047
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  ColumnWidth     =   1275.024
               EndProperty
               BeginProperty Column08 
                  Alignment       =   1
                  ColumnWidth     =   1289.764
               EndProperty
               BeginProperty Column09 
                  Alignment       =   1
                  ColumnWidth     =   1335.118
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   2264.882
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc adoformulacion 
            Height          =   330
            Left            =   120
            Top             =   4395
            Width           =   11090
            _ExtentX        =   19553
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
            Caption         =   "    "
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
      Begin Crystal.CrystalReport CryAREAS 
         Left            =   600
         Top             =   -600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin Crystal.CrystalReport Cryempresas 
         Left            =   600
         Top             =   -600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin Crystal.CrystalReport Crypersonal 
         Left            =   600
         Top             =   -600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin Crystal.CrystalReport CryClientes 
         Left            =   600
         Top             =   -600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   7455
         Left            =   -74520
         TabIndex        =   9
         Top             =   120
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   13150
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   8421376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "ADICIONES/REDUCIONES"
         TabPicture(0)   =   "FrmFormulacion.frx":DB8B2
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label38"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblAdiciones2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "dtgAdicion"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "adoAdicion"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "fraprincipalAd"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "fragrabarAd"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Frame1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "TRANSFERENCIAS"
         TabPicture(1)   =   "FrmFormulacion.frx":DB8CE
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label35"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label2"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label1"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label16"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "DataCombo26"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "dtcTipoT"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "dtgTraspaso"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Adotraspaso"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Frame2"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "fraprincipalTr"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "fragrabarTr"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Text5"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "TxtResT"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Text6"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).ControlCount=   14
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            DataField       =   "nro_transaccion"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adotraspaso"
            Height          =   285
            Left            =   480
            TabIndex        =   126
            Text            =   "0"
            Top             =   3360
            Width           =   735
         End
         Begin VB.TextBox TxtResT 
            Alignment       =   2  'Center
            DataField       =   "resolucion"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adotraspaso"
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
            Left            =   6000
            TabIndex        =   125
            Top             =   3360
            Width           =   2175
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
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
            Left            =   8880
            TabIndex        =   124
            Text            =   "0"
            Top             =   3360
            Width           =   1695
         End
         Begin VB.Frame fragrabarTr 
            Height          =   1215
            Left            =   4200
            TabIndex        =   97
            Top             =   6240
            Width           =   3135
            Begin VB.CommandButton BtnCancelarT 
               BackColor       =   &H8000000A&
               Caption         =   "Cancelar"
               Height          =   675
               Left            =   1920
               Picture         =   "FrmFormulacion.frx":DB8EA
               Style           =   1  'Graphical
               TabIndex        =   146
               Top             =   360
               Width           =   765
            End
            Begin VB.CommandButton BtnGrabarT 
               BackColor       =   &H8000000A&
               Caption         =   "Grabar"
               Height          =   675
               Left            =   360
               Picture         =   "FrmFormulacion.frx":DBAF4
               Style           =   1  'Graphical
               TabIndex        =   143
               Top             =   360
               Width           =   765
            End
         End
         Begin VB.Frame fraprincipalTr 
            Height          =   1215
            Left            =   135
            TabIndex        =   96
            Top             =   6120
            Width           =   10935
            Begin VB.CommandButton BtnImprimirC 
               BackColor       =   &H8000000A&
               Caption         =   "Cmpbte."
               Height          =   720
               Left            =   6240
               Picture         =   "FrmFormulacion.frx":DBCFE
               Style           =   1  'Graphical
               TabIndex        =   154
               ToolTipText     =   "Imprime Lista de Personas"
               Top             =   360
               Width           =   765
            End
            Begin VB.CommandButton BtnImprimirD 
               BackColor       =   &H8000000A&
               Caption         =   "Imprimir"
               Height          =   720
               Left            =   7080
               Picture         =   "FrmFormulacion.frx":DC2BB
               Style           =   1  'Graphical
               TabIndex        =   153
               ToolTipText     =   "Imprime Lista de Personas"
               Top             =   360
               Width           =   765
            End
            Begin VB.CommandButton BtnBuscarT 
               BackColor       =   &H8000000A&
               Caption         =   "Buscar"
               Height          =   720
               Left            =   5400
               Picture         =   "FrmFormulacion.frx":DC878
               Style           =   1  'Graphical
               TabIndex        =   150
               ToolTipText     =   "Busca un Registro"
               Top             =   360
               Width           =   765
            End
            Begin VB.CommandButton BtnSalirT 
               BackColor       =   &H8000000A&
               Caption         =   "Cerrar"
               Height          =   720
               Left            =   9600
               Picture         =   "FrmFormulacion.frx":DCE30
               Style           =   1  'Graphical
               TabIndex        =   148
               Top             =   360
               Width           =   765
            End
            Begin VB.CommandButton BtnEliminarT 
               BackColor       =   &H8000000A&
               Caption         =   "Anular"
               Height          =   720
               Left            =   2040
               Picture         =   "FrmFormulacion.frx":DD03A
               Style           =   1  'Graphical
               TabIndex        =   142
               ToolTipText     =   "Anula Registro Activo"
               Top             =   360
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.CommandButton BtnModificarT 
               BackColor       =   &H8000000A&
               Caption         =   "Modificar"
               Height          =   720
               Left            =   1200
               Picture         =   "FrmFormulacion.frx":DDD04
               Style           =   1  'Graphical
               TabIndex        =   140
               ToolTipText     =   "Modifica Registro Activo"
               Top             =   360
               Width           =   765
            End
            Begin VB.CommandButton BtnAńadirT 
               BackColor       =   &H8000000A&
               Caption         =   "Nuevo"
               Height          =   720
               Left            =   360
               Picture         =   "FrmFormulacion.frx":DE2E4
               Style           =   1  'Graphical
               TabIndex        =   138
               ToolTipText     =   "Nuevo Registro"
               Top             =   360
               Width           =   765
            End
         End
         Begin VB.Frame Frame2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   2415
            Left            =   120
            TabIndex        =   84
            Top             =   3720
            Width           =   10935
            Begin VB.TextBox txtmontoDestino 
               Alignment       =   2  'Center
               DataField       =   "trn_monto_destino"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "Adotraspaso"
               Enabled         =   0   'False
               Height          =   285
               Left            =   5880
               TabIndex        =   107
               Text            =   "0"
               Top             =   1800
               Width           =   2055
            End
            Begin VB.TextBox txtmontoOrigenT 
               Alignment       =   2  'Center
               DataField       =   "trn_monto_origen"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "Adotraspaso"
               Enabled         =   0   'False
               Height          =   285
               Left            =   720
               TabIndex        =   85
               Text            =   "0"
               Top             =   1800
               Width           =   2055
            End
            Begin MSDataListLib.DataCombo dtcFteT 
               Bindings        =   "FrmFormulacion.frx":DE908
               DataField       =   "fte_codigo"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   720
               TabIndex        =   86
               Top             =   360
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "fte_codigo"
               BoundColumn     =   "fte_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo12 
               Bindings        =   "FrmFormulacion.frx":DE920
               DataField       =   "fte_codigo"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   1920
               TabIndex        =   87
               Top             =   360
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Fte_descripcion_larga"
               BoundColumn     =   "fte_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DtcOrgT 
               Bindings        =   "FrmFormulacion.frx":DE938
               DataField       =   "org_codigo"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   720
               TabIndex        =   88
               Top             =   720
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "org_codigo"
               BoundColumn     =   "org_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo16 
               Bindings        =   "FrmFormulacion.frx":DE953
               DataField       =   "org_codigo"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   1920
               TabIndex        =   89
               Top             =   720
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "org_descripcion"
               BoundColumn     =   "org_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcPryT 
               Bindings        =   "FrmFormulacion.frx":DE96E
               DataField       =   "pro_proyecto"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   1440
               TabIndex        =   90
               Top             =   1080
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "pro_proyecto"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo19 
               Bindings        =   "FrmFormulacion.frx":DE988
               DataField       =   "pro_proyecto"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   2880
               TabIndex        =   91
               Top             =   1080
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Pro_descripcion_larga"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcParT 
               Bindings        =   "FrmFormulacion.frx":DE9A2
               DataField       =   "par_codigo"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   720
               TabIndex        =   92
               Top             =   1440
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "par_codigo"
               BoundColumn     =   "par_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo22 
               Bindings        =   "FrmFormulacion.frx":DE9BB
               DataField       =   "par_codigo"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   1920
               TabIndex        =   93
               Top             =   1440
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Par_descripcion_larga"
               BoundColumn     =   "par_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcActT 
               Bindings        =   "FrmFormulacion.frx":DE9D4
               DataField       =   "pro_proyecto"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   2160
               TabIndex        =   94
               Top             =   1080
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "pro_actividad"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcProT 
               Bindings        =   "FrmFormulacion.frx":DE9EE
               DataField       =   "pro_proyecto"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   720
               TabIndex        =   95
               Top             =   1080
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "pro_programa"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcFteT_des 
               Bindings        =   "FrmFormulacion.frx":DEA08
               DataField       =   "fte_codigo_des"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   5880
               TabIndex        =   108
               Top             =   360
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "fte_codigo"
               BoundColumn     =   "fte_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo28 
               Bindings        =   "FrmFormulacion.frx":DEA20
               DataField       =   "fte_codigo_des"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   7080
               TabIndex        =   109
               Top             =   360
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Fte_descripcion_larga"
               BoundColumn     =   "fte_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DtcOrgT_des 
               Bindings        =   "FrmFormulacion.frx":DEA38
               DataField       =   "org_codigo_des"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   5880
               TabIndex        =   110
               Top             =   720
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "org_codigo"
               BoundColumn     =   "org_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo30 
               Bindings        =   "FrmFormulacion.frx":DEA53
               DataField       =   "org_codigo_des"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   7080
               TabIndex        =   111
               Top             =   720
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "org_descripcion"
               BoundColumn     =   "org_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcPryT_des 
               Bindings        =   "FrmFormulacion.frx":DEA6E
               DataField       =   "pro_proyecto_des"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   6600
               TabIndex        =   112
               Top             =   1080
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "pro_proyecto"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo32 
               Bindings        =   "FrmFormulacion.frx":DEA88
               DataField       =   "pro_proyecto_des"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   8040
               TabIndex        =   113
               Top             =   1080
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Pro_descripcion_larga"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcParT_des 
               Bindings        =   "FrmFormulacion.frx":DEAA2
               DataField       =   "par_codigo_des"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   5880
               TabIndex        =   114
               Top             =   1440
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "par_codigo"
               BoundColumn     =   "par_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo34 
               Bindings        =   "FrmFormulacion.frx":DEABB
               DataField       =   "par_codigo_des"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   7080
               TabIndex        =   115
               Top             =   1440
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Par_descripcion_larga"
               BoundColumn     =   "par_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcActT_des 
               Bindings        =   "FrmFormulacion.frx":DEAD4
               DataField       =   "pro_proyecto_des"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   7320
               TabIndex        =   116
               Top             =   1080
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "pro_actividad"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcProT_des 
               Bindings        =   "FrmFormulacion.frx":DEAEE
               DataField       =   "pro_proyecto_des"
               DataSource      =   "Adotraspaso"
               Height          =   315
               Left            =   5880
               TabIndex        =   117
               Top             =   1080
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "pro_programa"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   2
               X1              =   0
               X2              =   10935
               Y1              =   2370
               Y2              =   2385
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   2
               X1              =   10920
               X2              =   10920
               Y1              =   120
               Y2              =   2400
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   2
               X1              =   0
               X2              =   0
               Y1              =   120
               Y2              =   2400
            End
            Begin VB.Label Label39 
               Caption         =   "Monto"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   120
               TabIndex        =   123
               Top             =   1800
               Width           =   735
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               Caption         =   "DESTINO"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   7800
               TabIndex        =   119
               Top             =   0
               Width           =   825
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "ORIGEN"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   2400
               TabIndex        =   118
               Top             =   0
               Width           =   675
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   2
               X1              =   0
               X2              =   10935
               Y1              =   120
               Y2              =   135
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   2
               X1              =   5760
               X2              =   5760
               Y1              =   120
               Y2              =   2400
            End
            Begin VB.Label Label17 
               Caption         =   "Partida"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   120
               TabIndex        =   101
               Top             =   1440
               Width           =   735
            End
            Begin VB.Label Label18 
               Caption         =   "Fte"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   120
               TabIndex        =   100
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label26 
               Caption         =   "Org"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   120
               TabIndex        =   99
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label30 
               Caption         =   "Proy"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   120
               TabIndex        =   98
               Top             =   1080
               Width           =   735
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Registro"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   2775
            Left            =   -74880
            TabIndex        =   65
            Top             =   3240
            Width           =   10935
            Begin VB.TextBox txt_monto_total 
               Alignment       =   2  'Center
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               Height          =   285
               Left            =   8760
               TabIndex        =   134
               Text            =   "0"
               Top             =   2280
               Width           =   2055
            End
            Begin VB.TextBox txt_monto_new 
               Alignment       =   2  'Center
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               Height          =   285
               Left            =   5520
               TabIndex        =   133
               Text            =   "0"
               Top             =   2280
               Width           =   2055
            End
            Begin VB.TextBox Text9 
               Alignment       =   2  'Center
               DataField       =   "nro_transaccion"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adoAdicion"
               Height          =   285
               Left            =   1440
               TabIndex        =   80
               Text            =   "0"
               Top             =   375
               Width           =   975
            End
            Begin VB.TextBox txtmontoOrigen 
               Alignment       =   2  'Center
               DataField       =   "trn_monto_origen"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adoAdicion"
               Height          =   285
               Left            =   2280
               TabIndex        =   67
               Text            =   "0"
               Top             =   2280
               Width           =   2055
            End
            Begin VB.TextBox TxtRes 
               Alignment       =   2  'Center
               DataField       =   "resolucion"
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "adoAdicion"
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
               Left            =   8760
               TabIndex        =   66
               Top             =   375
               Width           =   2055
            End
            Begin MSDataListLib.DataCombo dtcFteA 
               Bindings        =   "FrmFormulacion.frx":DEB08
               DataField       =   "fte_codigo"
               DataSource      =   "adoAdicion"
               Height          =   315
               Left            =   2280
               TabIndex        =   68
               Top             =   840
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "fte_codigo"
               BoundColumn     =   "fte_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DtcFteDesA 
               Bindings        =   "FrmFormulacion.frx":DEB20
               DataField       =   "fte_codigo"
               DataSource      =   "adoAdicion"
               Height          =   315
               Left            =   3960
               TabIndex        =   69
               Top             =   840
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Fte_descripcion_larga"
               BoundColumn     =   "fte_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DtcOrgA 
               Bindings        =   "FrmFormulacion.frx":DEB38
               DataField       =   "org_codigo"
               DataSource      =   "adoAdicion"
               Height          =   315
               Left            =   2280
               TabIndex        =   70
               Top             =   1200
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "org_codigo"
               BoundColumn     =   "org_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DtcOrgDesA 
               Bindings        =   "FrmFormulacion.frx":DEB53
               DataField       =   "org_codigo"
               DataSource      =   "adoAdicion"
               Height          =   315
               Left            =   3960
               TabIndex        =   71
               Top             =   1200
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "org_descripcion"
               BoundColumn     =   "org_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcPryA 
               Bindings        =   "FrmFormulacion.frx":DEB6E
               DataField       =   "pro_proyecto"
               DataSource      =   "adoAdicion"
               Height          =   315
               Left            =   3000
               TabIndex        =   72
               Top             =   1560
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "pro_proyecto"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DtcPryDes 
               Bindings        =   "FrmFormulacion.frx":DEB88
               DataField       =   "pro_proyecto"
               DataSource      =   "adoAdicion"
               Height          =   315
               Left            =   4560
               TabIndex        =   73
               Top             =   1560
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Pro_descripcion_larga"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcParA 
               Bindings        =   "FrmFormulacion.frx":DEBA2
               DataField       =   "par_codigo"
               DataSource      =   "adoAdicion"
               Height          =   315
               Left            =   2280
               TabIndex        =   74
               Top             =   1920
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "par_codigo"
               BoundColumn     =   "par_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DtcPasDesA 
               Bindings        =   "FrmFormulacion.frx":DEBBB
               DataField       =   "par_codigo"
               DataSource      =   "adoAdicion"
               Height          =   315
               Left            =   3960
               TabIndex        =   75
               Top             =   1920
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Par_descripcion_larga"
               BoundColumn     =   "par_codigo"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcActA 
               Bindings        =   "FrmFormulacion.frx":DEBD4
               DataField       =   "pro_proyecto"
               DataSource      =   "adoAdicion"
               Height          =   315
               Left            =   3720
               TabIndex        =   76
               Top             =   1560
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "pro_actividad"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcProA 
               Bindings        =   "FrmFormulacion.frx":DEBEE
               DataField       =   "pro_proyecto"
               DataSource      =   "adoAdicion"
               Height          =   315
               Left            =   2280
               TabIndex        =   77
               Top             =   1560
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "pro_programa"
               BoundColumn     =   "pro_proyecto"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcTipoA 
               Bindings        =   "FrmFormulacion.frx":DEC08
               DataField       =   "tipo_transaccion"
               DataSource      =   "adoAdicion"
               Height          =   315
               Left            =   4080
               TabIndex        =   82
               Top             =   375
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "tipo_transaccion"
               BoundColumn     =   "tipo_transaccion"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo dtcTipoDesA 
               Bindings        =   "FrmFormulacion.frx":DEC1E
               DataField       =   "tipo_transaccion"
               DataSource      =   "adoAdicion"
               Height          =   315
               Left            =   4935
               TabIndex        =   83
               Top             =   375
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "nombre_transaccion"
               BoundColumn     =   "tipo_transaccion"
               Text            =   "DataCombo1"
            End
            Begin MSAdodcLib.Adodc AdoTipo 
               Height          =   375
               Left            =   4080
               Top             =   360
               Visible         =   0   'False
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   661
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
               Caption         =   "Tipo"
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
            Begin VB.Label Label41 
               Alignment       =   2  'Center
               Caption         =   "="
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   7920
               TabIndex        =   136
               Top             =   2280
               Width           =   495
            End
            Begin VB.Label Label40 
               Alignment       =   2  'Center
               Caption         =   "+ / -"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   4680
               TabIndex        =   135
               Top             =   2280
               Width           =   495
            End
            Begin VB.Label Label34 
               Caption         =   "Monto Transacción Bs."
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   120
               TabIndex        =   106
               Top             =   2280
               Width           =   2055
            End
            Begin VB.Label Label33 
               Caption         =   "Partida del Gasto"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   120
               TabIndex        =   105
               Top             =   1920
               Width           =   2055
            End
            Begin VB.Label Label15 
               Caption         =   "Fuente de Financiamiento"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   120
               TabIndex        =   104
               Top             =   840
               Width           =   2055
            End
            Begin VB.Label Label13 
               Caption         =   "Organismo Financiador"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   120
               TabIndex        =   103
               Top             =   1200
               Width           =   2055
            End
            Begin VB.Label Label9 
               Caption         =   "Categoría Programática"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   120
               TabIndex        =   102
               Top             =   1560
               Width           =   2055
            End
            Begin VB.Label Label32 
               Caption         =   "Tipo de Registro"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   2760
               TabIndex        =   81
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label31 
               Caption         =   "Numero Registro "
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   120
               TabIndex        =   79
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label14 
               Caption         =   "Nro. Resolución"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   255
               Left            =   7440
               TabIndex        =   78
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame fragrabarAd 
            Height          =   1215
            Left            =   -71160
            TabIndex        =   63
            Top             =   6120
            Width           =   3135
            Begin VB.CommandButton BtnCancelarA 
               BackColor       =   &H8000000A&
               Caption         =   "Cancelar"
               Height          =   675
               Left            =   1920
               Picture         =   "FrmFormulacion.frx":DEC34
               Style           =   1  'Graphical
               TabIndex        =   145
               Top             =   360
               Width           =   765
            End
            Begin VB.CommandButton BtnGrabarA 
               BackColor       =   &H8000000A&
               Caption         =   "Grabar"
               Height          =   675
               Left            =   360
               Picture         =   "FrmFormulacion.frx":DEE3E
               Style           =   1  'Graphical
               TabIndex        =   144
               Top             =   360
               Width           =   765
            End
         End
         Begin VB.Frame fraprincipalAd 
            Height          =   1215
            Left            =   -74865
            TabIndex        =   62
            Top             =   6000
            Width           =   10935
            Begin VB.CommandButton BtnImprimirA 
               BackColor       =   &H8000000A&
               Caption         =   "Cmpbte."
               Height          =   720
               Left            =   6360
               Picture         =   "FrmFormulacion.frx":DF048
               Style           =   1  'Graphical
               TabIndex        =   152
               ToolTipText     =   "Imprime Lista de Personas"
               Top             =   360
               Width           =   765
            End
            Begin VB.CommandButton BtnImprimirB 
               BackColor       =   &H8000000A&
               Caption         =   "Listado"
               Height          =   720
               Left            =   7200
               Picture         =   "FrmFormulacion.frx":DF605
               Style           =   1  'Graphical
               TabIndex        =   151
               ToolTipText     =   "Imprime Lista de Personas"
               Top             =   360
               Width           =   765
            End
            Begin VB.CommandButton BtnBuscarA 
               BackColor       =   &H8000000A&
               Caption         =   "Buscar"
               Height          =   720
               Left            =   5520
               Picture         =   "FrmFormulacion.frx":DFBC2
               Style           =   1  'Graphical
               TabIndex        =   149
               ToolTipText     =   "Busca un Registro"
               Top             =   360
               Width           =   765
            End
            Begin VB.CommandButton BtnSalirA 
               BackColor       =   &H8000000A&
               Caption         =   "Cerrar"
               Height          =   720
               Left            =   9720
               Picture         =   "FrmFormulacion.frx":E017A
               Style           =   1  'Graphical
               TabIndex        =   147
               Top             =   360
               Width           =   765
            End
            Begin VB.CommandButton BtnEliminarA 
               BackColor       =   &H8000000A&
               Caption         =   "Anular"
               Height          =   720
               Left            =   2040
               Picture         =   "FrmFormulacion.frx":E0384
               Style           =   1  'Graphical
               TabIndex        =   141
               ToolTipText     =   "Anula Registro Activo"
               Top             =   360
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.CommandButton BtnModificarA 
               BackColor       =   &H8000000A&
               Caption         =   "Modificar"
               Height          =   720
               Left            =   1200
               Picture         =   "FrmFormulacion.frx":E104E
               Style           =   1  'Graphical
               TabIndex        =   139
               ToolTipText     =   "Modifica Registro Activo"
               Top             =   360
               Width           =   765
            End
            Begin VB.CommandButton BtnAńadirA 
               BackColor       =   &H8000000A&
               Caption         =   "Nuevo"
               Height          =   720
               Left            =   360
               Picture         =   "FrmFormulacion.frx":E162E
               Style           =   1  'Graphical
               TabIndex        =   137
               ToolTipText     =   "Nuevo Registro"
               Top             =   360
               Width           =   765
            End
         End
         Begin VB.TextBox TxtId_ProcesoCd 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            DataField       =   "id_proceso"
            DataSource      =   "AdoDetalleCd"
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
            Left            =   -73800
            TabIndex        =   33
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox TxtEtapaCd 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            DataField       =   "etapa_tramite"
            DataSource      =   "AdoDetalleCd"
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
            Left            =   -73800
            TabIndex        =   32
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox TxtDescripcionCd 
            DataField       =   "descripcion_etapa"
            DataSource      =   "AdoDetalleCd"
            Height          =   285
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   1  'Horizontal
            TabIndex        =   30
            Top             =   2040
            Width           =   5895
         End
         Begin VB.TextBox TxtLugarCd 
            DataField       =   "lugar_etapa"
            DataSource      =   "AdoDetalleCd"
            Height          =   285
            Left            =   -72840
            TabIndex        =   29
            Top             =   2520
            Width           =   3975
         End
         Begin VB.TextBox TxtCiteCd 
            DataField       =   "otrosi_cite_doc"
            DataSource      =   "AdoDetalleCd"
            Height          =   285
            Left            =   -72000
            TabIndex        =   27
            Top             =   3600
            Width           =   1695
         End
         Begin VB.Frame FraPrincipalCd 
            Height          =   1215
            Left            =   -74880
            TabIndex        =   20
            Top             =   4680
            Width           =   6255
            Begin VB.CommandButton CmdAdicionarCd 
               Caption         =   "&Adicionar"
               DownPicture     =   "FrmFormulacion.frx":E1C52
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   1
               Picture         =   "FrmFormulacion.frx":E2094
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   1
               Width           =   975
            End
            Begin VB.CommandButton CmdModificarCd 
               Caption         =   "&Modificar"
               DownPicture     =   "FrmFormulacion.frx":E24D6
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   1440
               Picture         =   "FrmFormulacion.frx":E2918
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton CmdEliminarCd 
               Caption         =   "&Eliminar"
               DownPicture     =   "FrmFormulacion.frx":E2D5A
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   2640
               Picture         =   "FrmFormulacion.frx":E3064
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton CmdBuscarCd 
               Caption         =   "&Buscar"
               DownPicture     =   "FrmFormulacion.frx":E34A6
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   3840
               Picture         =   "FrmFormulacion.frx":E38E8
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton CmdImprimirCd 
               Caption         =   "&Imprimir"
               DownPicture     =   "FrmFormulacion.frx":E3D2A
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   5040
               Picture         =   "FrmFormulacion.frx":E416C
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame6 
            Height          =   1215
            Left            =   -71400
            TabIndex        =   15
            Top             =   5520
            Width           =   3855
            Begin VB.CommandButton Command7 
               Caption         =   "&Buscar"
               DownPicture     =   "FrmFormulacion.frx":E47D6
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   240
               Picture         =   "FrmFormulacion.frx":E4C18
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command8 
               Caption         =   "&Imprimir"
               DownPicture     =   "FrmFormulacion.frx":E505A
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   1440
               Picture         =   "FrmFormulacion.frx":E549C
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command9 
               Caption         =   "&Salir"
               DownPicture     =   "FrmFormulacion.frx":E5B06
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   2640
               Picture         =   "FrmFormulacion.frx":E5F48
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame fragrabarCd 
            Height          =   1215
            Left            =   -73080
            TabIndex        =   12
            Top             =   4680
            Width           =   2655
            Begin VB.CommandButton CmdCancelarCd 
               Caption         =   "&Cancelar"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   1440
               Picture         =   "FrmFormulacion.frx":E638A
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton CmdGrabarCd 
               Caption         =   "&Grabar"
               DragIcon        =   "FrmFormulacion.frx":E67CC
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   120
               Picture         =   "FrmFormulacion.frx":E6C0E
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox TxtAuxId 
            Height          =   285
            Left            =   -72480
            TabIndex        =   11
            Top             =   840
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox TxtAuxProceso 
            BackColor       =   &H00C0E0FF&
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
            Left            =   -72000
            TabIndex        =   10
            Top             =   840
            Width           =   7815
         End
         Begin MSAdodcLib.Adodc adoAdicion 
            Height          =   330
            Left            =   -74850
            Top             =   2895
            Width           =   10920
            _ExtentX        =   19262
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
            BackColor       =   12640511
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
            Caption         =   " <-- Inicio                  Fin -->"
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
         Begin MSDataGridLib.DataGrid DtgCivilFinCd 
            Bindings        =   "FrmFormulacion.frx":E7050
            Height          =   2175
            Left            =   -74640
            TabIndex        =   19
            Top             =   3000
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   3836
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
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
            Caption         =   "DETALLE DE LOS PROCESOS CIVILES"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DtgCivilCd 
            Bindings        =   "FrmFormulacion.frx":E706E
            Height          =   2655
            Left            =   -68520
            TabIndex        =   26
            Top             =   1320
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   4683
            _Version        =   393216
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "id_proceso"
               Caption         =   "Nro.Proceso"
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
               DataField       =   "etapa_tramite"
               Caption         =   "Etapa"
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
               DataField       =   "descripcion_etapa"
               Caption         =   "Descipción de la Etapa"
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
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcTipoDocCd 
            Bindings        =   "FrmFormulacion.frx":E7089
            DataField       =   "id_tipo_doc"
            DataSource      =   "AdoDetalleCd"
            Height          =   315
            Left            =   -72000
            TabIndex        =   28
            Top             =   3000
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "descripcion_documento"
            BoundColumn     =   "id_tipo_doc"
            Text            =   "DataCombo17"
         End
         Begin MSComCtl2.DTPicker DTPFechaCd 
            DataField       =   "fecha_etapa"
            DataSource      =   "AdoDetalleCd"
            Height          =   315
            Left            =   -70440
            TabIndex        =   31
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   57344001
            CurrentDate     =   36775
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Bindings        =   "FrmFormulacion.frx":E70A2
            DataField       =   "id_tipo_doc"
            DataSource      =   "AdoDetalleCd"
            Height          =   315
            Left            =   -72360
            TabIndex        =   34
            Top             =   3000
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "id_tipo_doc"
            BoundColumn     =   "id_tipo_doc"
            Text            =   "DataCombo17"
         End
         Begin MSDataGridLib.DataGrid dtgAdicion 
            Bindings        =   "FrmFormulacion.frx":E70BB
            Height          =   2055
            Left            =   -74880
            TabIndex        =   64
            Top             =   480
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12648447
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   19
            RowDividerStyle =   3
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "ADICIONES o REDUCCIONES PRESUPUESTARIAS"
            ColumnCount     =   10
            BeginProperty Column00 
               DataField       =   "nro_transaccion"
               Caption         =   "Nro."
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
               DataField       =   "tipo_transaccion"
               Caption         =   "Tipo"
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
               DataField       =   "fte_codigo"
               Caption         =   "Fte"
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
               DataField       =   "org_codigo"
               Caption         =   "Org"
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
               DataField       =   "pro_programa"
               Caption         =   "Pro"
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
               DataField       =   "pro_proyecto"
               Caption         =   "Pry"
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
               DataField       =   "pro_actividad"
               Caption         =   "Act"
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
               DataField       =   "par_codigo"
               Caption         =   "Partida"
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
               DataField       =   "trn_monto_origen"
               Caption         =   "Add/Red.Bs."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "par_descripcion_larga"
               Caption         =   "      Descripción"
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
                  ColumnWidth     =   599.811
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   675.213
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   494.929
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   524.976
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   510.236
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   434.835
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   524.976
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column08 
                  Alignment       =   1
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   4229.858
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Adotraspaso 
            Height          =   330
            Left            =   120
            Top             =   2640
            Width           =   10920
            _ExtentX        =   19262
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
            BackColor       =   12640511
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
            Caption         =   " <-- Inicio                  Fin -->"
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
         Begin MSDataGridLib.DataGrid dtgTraspaso 
            Bindings        =   "FrmFormulacion.frx":E70D4
            Height          =   2055
            Left            =   120
            TabIndex        =   120
            Top             =   480
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12648447
            ColumnHeaders   =   -1  'True
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   19
            RowDividerStyle =   3
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "            <-   ORIGEN    -                     I     I                    -    DESTINO ->"
            ColumnCount     =   17
            BeginProperty Column00 
               DataField       =   "nro_transaccion"
               Caption         =   "Nro."
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
               DataField       =   "tipo_transaccion"
               Caption         =   "Tipo"
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
               DataField       =   "fte_codigo"
               Caption         =   "Fte"
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
               DataField       =   "org_codigo"
               Caption         =   "Org"
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
               DataField       =   "pro_programa"
               Caption         =   "Pro"
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
               DataField       =   "pro_proyecto"
               Caption         =   "Pry"
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
               DataField       =   "pro_actividad"
               Caption         =   "Act"
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
               DataField       =   "par_codigo"
               Caption         =   "Partida"
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
               DataField       =   "trn_monto_origen"
               Caption         =   "Monto Origen"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   ""
               Caption         =   ""
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
            BeginProperty Column10 
               DataField       =   "fte_codigo_des"
               Caption         =   "Fte"
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
               DataField       =   "org_codigo_des"
               Caption         =   "Org"
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
               DataField       =   "pro_programa_des"
               Caption         =   "Pro"
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
            BeginProperty Column13 
               DataField       =   "pro_proyecto_des"
               Caption         =   "Pry"
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
            BeginProperty Column14 
               DataField       =   "pro_actividad_des"
               Caption         =   "Act"
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
            BeginProperty Column15 
               DataField       =   "par_codigo_des"
               Caption         =   "Partida"
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
            BeginProperty Column16 
               DataField       =   "trn_monto_destino"
               Caption         =   "Monto Destino"
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
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   540.284
               EndProperty
               BeginProperty Column01 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   585.071
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   404.787
               EndProperty
               BeginProperty Column03 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   494.929
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   434.835
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   404.787
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   390.047
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   854.929
               EndProperty
               BeginProperty Column08 
                  Alignment       =   1
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column09 
                  DividerStyle    =   1
                  ColumnWidth     =   299.906
               EndProperty
               BeginProperty Column10 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   450.142
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   464.882
               EndProperty
               BeginProperty Column12 
                  ColumnWidth     =   434.835
               EndProperty
               BeginProperty Column13 
                  ColumnWidth     =   480.189
               EndProperty
               BeginProperty Column14 
                  ColumnWidth     =   434.835
               EndProperty
               BeginProperty Column15 
                  ColumnWidth     =   750.047
               EndProperty
               BeginProperty Column16 
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtcTipoT 
            Bindings        =   "FrmFormulacion.frx":E70EE
            DataField       =   "tipo_transaccion"
            DataSource      =   "Adotraspaso"
            Height          =   315
            Left            =   1920
            TabIndex        =   127
            Top             =   3360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "tipo_transaccion"
            BoundColumn     =   "tipo_transaccion"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo26 
            Bindings        =   "FrmFormulacion.frx":E7104
            DataField       =   "tipo_transaccion"
            DataSource      =   "Adotraspaso"
            Height          =   315
            Left            =   2760
            TabIndex        =   128
            Top             =   3360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "nombre_transaccion"
            BoundColumn     =   "tipo_transaccion"
            Text            =   "DataCombo1"
         End
         Begin VB.Label Label16 
            Caption         =   "Monto"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   9480
            TabIndex        =   132
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Registro"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   2520
            TabIndex        =   131
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Numero Registro "
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   240
            TabIndex        =   130
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Label Label35 
            Caption         =   "Nro. Resolución"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   6360
            TabIndex        =   129
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Label lblAdiciones2 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   -70560
            TabIndex        =   122
            Top             =   2520
            Width           =   1260
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            Caption         =   "Totales :"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   -72600
            TabIndex        =   121
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label Label25 
            Caption         =   "Proceso"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   -74760
            TabIndex        =   41
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label24 
            Caption         =   "Etapa"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   -74760
            TabIndex        =   40
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de la Etapa"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   225
            Left            =   -72000
            TabIndex        =   39
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label22 
            Caption         =   "Descripción de la Etapa"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   -74760
            TabIndex        =   38
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label21 
            Caption         =   "Lugar del Proceso"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   -74760
            TabIndex        =   37
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label Label20 
            Caption         =   "Tipo de Doc. que se emite o recibe"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   -74760
            TabIndex        =   36
            Top             =   3000
            Width           =   2655
         End
         Begin VB.Label Label19 
            Caption         =   "Cite/Otrosi del Documento Emitido"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   -74760
            TabIndex        =   35
            Top             =   3600
            Width           =   2895
         End
      End
   End
   Begin Crystal.CrystalReport crTraspaso 
      Left            =   120
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Detalle de la Venta de Pliegos"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "FrmFormulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sino, swgraba, gestion, solo_form As String
Dim varNro, varTipo, varRes, varFte, varorg, varpro, varpry, varAct, varpar, varmontoO As String
Dim varFteD, varorgD, varproD, varpryD, varActD, varparD, varmontoD As String
Dim varnroF As Integer
Dim OriDes, varbusca, parametro, tipoT As String
Dim montoTotal, montoTotalA, montoTotalM, montoTotalA2 As Currency
Public CAMPOS As Variant

Dim rsfuente, rsOrganismo, rsproyecto, rspartida As New ADODB.Recordset
Dim rsTipo, rsRepAdd, rsAdicion, rsformulacion As New ADODB.Recordset
Dim rsTraspaso, rsNada, RsCompro As New ADODB.Recordset

Private Sub adoAdicion_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If (Not adoAdicion.Recordset.EOF Or Not adoAdicion.Recordset.BOF) And swgraba <> "A" Then
        txt_monto_total = adoAdicion.Recordset("trn_monto_origen") + Val(txt_monto_new)
    End If
End Sub

Private Sub adoformulacion_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If (Not adoformulacion.Recordset.EOF Or Not adoformulacion.Recordset.BOF) And swgraba <> "A" Then
        'gestion = adoformulacion.Recordset("ges_gestion")
        'Dtgformulacion.Caption = " FORMULACION PRESUPUESTARIA  - " + ((adoformulacion.Recordset("ges_gestion"))
    End If

End Sub

Private Sub BtnImprimirC_Click()
'Imprime Comprobante de Transferencia
'copia detalle de items (bien o servicio)
       Set rsRepAdd = New ADODB.Recordset
       db.Execute "DELETE from po_formulacion_trn_rep"
       If rsRepAdd.State = 1 Then rsRepAdd.Close
       rsRepAdd.Open "select * from po_formulacion_trn_rep ", db, adOpenKeyset, adLockOptimistic
       Set rsAdicion = New ADODB.Recordset
       If rsAdicion.State = 1 Then rsAdicion.Close
       rsAdicion.Open "select * from po_formulacion_trn where nro_transaccion=" & Text6.Text & " and tipo_transaccion='" & dtcTipoT.Text & "'", db, adOpenKeyset, adLockOptimistic
       If rsAdicion.RecordCount > 0 Then
          While Not rsAdicion.EOF
            rsRepAdd.AddNew
            rsRepAdd("nro_transaccion") = rsAdicion("nro_transaccion")
            rsRepAdd("tipo_transaccion") = rsAdicion("tipo_transaccion")
            rsRepAdd("uni_codigo") = rsAdicion("uni_codigo")
            rsRepAdd("pro_programa") = rsAdicion("pro_programa")
            rsRepAdd("pro_proyecto") = rsAdicion("pro_proyecto")
            rsRepAdd("pro_actividad") = rsAdicion("pro_actividad")
            rsRepAdd("fte_codigo") = rsAdicion("fte_codigo")
            rsRepAdd("org_codigo") = rsAdicion("org_codigo")
            rsRepAdd("par_codigo") = rsAdicion("par_codigo")
            rsRepAdd("ent_codigo") = rsAdicion("ent_codigo")
            rsRepAdd("trn_monto_origen") = rsAdicion("trn_monto_origen")
            
            rsRepAdd("uni_codigo_des") = rsAdicion("uni_codigo_des")
            rsRepAdd("pro_programa_des") = rsAdicion("pro_programa_des")
            rsRepAdd("pro_proyecto_des") = rsAdicion("pro_proyecto_des")
            rsRepAdd("pro_actividad_des") = rsAdicion("pro_actividad_des")
            rsRepAdd("fte_codigo_des") = rsAdicion("fte_codigo_des")
            rsRepAdd("org_codigo_des") = rsAdicion("org_codigo_des")
            rsRepAdd("par_codigo_des") = rsAdicion("par_codigo_des")
            rsRepAdd("ent_codigo_des") = rsAdicion("ent_codigo_des")
            rsRepAdd("trn_monto_destino") = rsAdicion("trn_monto_destino")
            
            rsRepAdd("resolucion") = rsAdicion("resolucion")
            rsRepAdd("fecha_transaccion") = IIf(IsNull(rsAdicion("fecha_transaccion")), Date, CDate(rsAdicion("fecha_transaccion")))
'adjudicado = IIf(IsNull(rsauxiliarmax!adjudicado), "N", rsauxiliarmax!adjudicado)
            rsRepAdd.Update
            rsAdicion.MoveNext
          Wend
       End If
'fin copia detalle de items (bien o servicio)
Dim iResult As Variant

'crPliegos.Formulas(0) = "TCompra='" & TxtCompra.Text & "'"
'crPliegos.Formulas(9) = "tfechaLimite='" & CStr(Day(DTPFechaLimite.Value)) & "  de  " & meses(Month(DTPFechaLimite.Value)) & "  de  " & CStr(Year(DTPFechaLimite.Value)) & "'"
    
    crTraspaso.ReportFileName = App.Path & "\Reportes\ComproModPptoT.rpt"
    
    iResult = crTraspaso.PrintReport
    If iResult <> 0 Then
     MsgBox crTraspaso.LastErrorNumber & "   " & crTraspaso.LastErrorString
    End If

End Sub

Private Sub BtnImprimirD_Click()
'Imprime Lista de Traspasos Presupuestarios
''copia detalle de items (bien o servicio)
'       Set rsRepDet = New ADODB.Recordset
'       db.Execute "DELETE from ao_no_objecion_detalle_rep"
'       If rsRepDet.State = 1 Then rsRepDet.Close
'       rsRepDet.Open "select * from ao_no_objecion_detalle_rep ", db, adOpenKeyset, adLockOptimistic
'       Set rsdetalle = New ADODB.Recordset
'       If rsdetalle.State = 1 Then rsdetalle.Close
'       rsdetalle.Open "select * from ao_no_objecion_detalle_D where nro_licitacion='" & TxtCompra.Text & "' ", db, adOpenKeyset, adLockOptimistic
'       If rsdetalle.RecordCount > 0 Then
'          While Not rsdetalle.EOF
'            rsRepDet.AddNew
'            rsRepDet("nro_licitacion") = rsdetalle("nro_licitacion")
'            rsRepDet("nro_licitacion_detalle") = rsdetalle("nro_licitacion_detalle")
'            rsRepDet("ges_gestion") = rsdetalle("ges_gestion")
'            rsRepDet("codGrupo") = rsdetalle("codGrupo")
'            rsRepDet("CodDetalle") = rsdetalle("CodDetalle")
'            rsRepDet("descripcion_bien") = rsdetalle("descripcion_bien")
'            rsRepDet.Update
'            rsdetalle.MoveNext
'          Wend
'       End If
''fin copia detalle de items (bien o servicio)


Dim iResult As Variant

'crTraspaso.Formulas(0) = "TCompra='" & TxtCompra.Text & "'"
'crTraspaso.Formulas(1) = "tproveedor='" & txtproveedor.Text & "'"
'crTraspaso.Formulas(2) = "tgestion='" & txtGestion.Text & "'"
'crTraspaso.Formulas(3) = "tNroPliego='" & txtNroPliego.Text & "'"
'crTraspaso.Formulas(4) = "tsolicitud='" & txtsolicitud.Text & "'"
'crTraspaso.Formulas(5) = "tformulario='" & txtformulario.Text & "'"
'crTraspaso.Formulas(6) = "TRUC='" & TxtRUC.Text & "'"
'crTraspaso.Formulas(7) = "tfecha='" & TxtFecha.Text & "'"
'crTraspaso.Formulas(8) = "fecha ='La Paz, " & meses(Month(Date)) & " " & CStr(Day(Date)) & " del " & CStr(Year(Date)) & "'"
'crTraspaso.Formulas(9) = "tfechaLimite='" & CStr(Day(DTPFechaLimite.Value)) & "  de  " & meses(Month(DTPFechaLimite.Value)) & "  de  " & CStr(Year(DTPFechaLimite.Value)) & "'"
'crTraspaso.Formulas(10) = "Tcarta='" & txtcuenta.Text & "'"
    
    'crTraspaso.ReportFileName = App.Path & "\Sistemas\Reportes\Modificacion PRESUPUESTARIA 2.rpt"
    crTraspaso.ReportFileName = App.Path & "\Reportes\Modificacion PRESUPUESTARIA 2.rpt"
    iResult = crTraspaso.PrintReport
    If iResult <> 0 Then
     MsgBox crTraspaso.LastErrorNumber & "   " & crTraspaso.LastErrorString
    End If

End Sub

Private Sub BtnAńadir_Click()
    MsgBox "No se puede adicionar el formulado, cuando existe una Adición o Transferencia ..."
End Sub

Private Sub BtnAńadirA_Click()
    swgraba = "A"
    adoAdicion.Recordset.AddNew
    fraprincipalAd.Visible = False
    fragrabarAd.Visible = True
    Frame1.Enabled = True
    Text9.Visible = False
    txt_monto_new.Enabled = False
    txt_monto_total.Enabled = False
End Sub

Private Sub BtnAńadirT_Click()
    Adotraspaso.Recordset.AddNew
    FrmOrigenDestino.Show 'vbModal
    swgraba = "A"
    fraprincipalTr.Visible = False
    fragrabarTr.Visible = True
    Frame2.Enabled = True
    
    Text5.Visible = True
    Label16.Visible = True
    Text6.Visible = False
End Sub

Private Sub BtnBuscar_Click()
'On Error GoTo Error:
'    OriDes = "F"
'    varbusca = "FOR"
'    For Each CAMPOS In rsformulacion.Fields
'        FrmBusqueda.CmbCampo.AddItem CAMPOS.Name
'    Next CAMPOS
'    FrmBusqueda.Show
'Exit Sub
'Error:
'    MsgBox "Existe error de sintaxis", vbDefaultButton2, "ERROR"

End Sub

Private Sub BtnBuscarT_Click()
'On Error GoTo Error:
'    varbusca = "TRF"
'    For Each CAMPOS In rsTraspaso.Fields
'        FrmBusqueda.CmbCampo.AddItem CAMPOS.Name
'    Next CAMPOS
'    FrmBusqueda.Show
'Exit Sub
'Error:
'    MsgBox "Existe error de sintaxis", vbDefaultButton2, "ERROR"

End Sub

Private Sub BtnCancelarA_Click()
'    If TxtRes.Text <> "" Then
'        adoAdicion.Recordset.CancelUpdate
'    End If
    parametro = "fv_formulacion_trn.tipo_transaccion" + " = " + "'A'" + " or " + "fv_formulacion_trn.tipo_transaccion" + " = " + "'R'"
    Call abrir_adicion                   'Abrir Adicion o Reducion
    
    fraprincipalAd.Visible = True
    fragrabarAd.Visible = False
    txt_monto_new.Enabled = True
    txt_monto_total.Enabled = True
    Frame1.Enabled = False
    Text9.Visible = True
    Call Objetos_Ad
End Sub

Private Sub BtnCancelarT_Click()
On Error GoTo Error:
    'Adotraspaso.Recordset.CancelUpdate
    parametro = "po_formulacion_trn.tipo_transaccion" + " = " + "'T'" + " or " + "po_formulacion_trn.tipo_transaccion" + " = " + "'F'"
    Call abrir_traspaso
    fraprincipalTr.Visible = True
    fragrabarTr.Visible = False
    Frame2.Enabled = False
    
    Text5.Visible = False
    Label16.Visible = False
    Text6.Visible = True
Exit Sub
Error:
    MsgBox "Error: No se concluyó el proceso ...", vbDefaultButton2, "ERROR"

End Sub

Private Sub BtnImprimirA_Click()
'Imprime Comprobante de Adicion/Reduccion
'copia detalle de items (bien o servicio)
       Set rsRepAdd = New ADODB.Recordset
       db.Execute "DELETE from po_formulacion_trn_rep"
       If rsRepAdd.State = 1 Then rsRepAdd.Close
       rsRepAdd.Open "select * from po_formulacion_trn_rep ", db, adOpenKeyset, adLockOptimistic
       Set rsAdicion = New ADODB.Recordset
       If rsAdicion.State = 1 Then rsAdicion.Close
       rsAdicion.Open "select * from po_formulacion_trn where nro_transaccion=" & Text9.Text & " and tipo_transaccion='" & dtcTipoA.Text & "'", db, adOpenKeyset, adLockOptimistic
       If rsAdicion.RecordCount > 0 Then
          While Not rsAdicion.EOF
            rsRepAdd.AddNew
            rsRepAdd("nro_transaccion") = rsAdicion("nro_transaccion")
            rsRepAdd("tipo_transaccion") = rsAdicion("tipo_transaccion")
            rsRepAdd("uni_codigo") = rsAdicion("uni_codigo")
            rsRepAdd("pro_programa") = rsAdicion("pro_programa")
            rsRepAdd("pro_proyecto") = rsAdicion("pro_proyecto")
            rsRepAdd("pro_actividad") = rsAdicion("pro_actividad")
            rsRepAdd("fte_codigo") = rsAdicion("fte_codigo")
            rsRepAdd("org_codigo") = rsAdicion("org_codigo")
            rsRepAdd("par_codigo") = rsAdicion("par_codigo")
            rsRepAdd("ent_codigo") = rsAdicion("ent_codigo")
            rsRepAdd("trn_monto_origen") = rsAdicion("trn_monto_origen")
            
            rsRepAdd("uni_codigo_des") = rsAdicion("uni_codigo_des")
            rsRepAdd("pro_programa_des") = rsAdicion("pro_programa_des")
            rsRepAdd("pro_proyecto_des") = rsAdicion("pro_proyecto_des")
            rsRepAdd("pro_actividad_des") = rsAdicion("pro_actividad_des")
            rsRepAdd("fte_codigo_des") = rsAdicion("fte_codigo_des")
            rsRepAdd("org_codigo_des") = rsAdicion("org_codigo_des")
            rsRepAdd("par_codigo_des") = rsAdicion("par_codigo_des")
            rsRepAdd("ent_codigo_des") = rsAdicion("ent_codigo_des")
            rsRepAdd("trn_monto_destino") = rsAdicion("trn_monto_destino")
            
            rsRepAdd("resolucion") = rsAdicion("resolucion")
            rsRepAdd("fecha_transaccion") = IIf(IsNull(rsAdicion("fecha_transaccion")), Date, CDate(rsAdicion("fecha_transaccion")))
'adjudicado = IIf(IsNull(rsauxiliarmax!adjudicado), "N", rsauxiliarmax!adjudicado)
            rsRepAdd.Update
            rsAdicion.MoveNext
          Wend
       End If
'fin copia detalle de items (bien o servicio)
Dim iResult As Variant

'crPliegos.Formulas(0) = "TCompra='" & TxtCompra.Text & "'"
'crPliegos.Formulas(9) = "tfechaLimite='" & CStr(Day(DTPFechaLimite.Value)) & "  de  " & meses(Month(DTPFechaLimite.Value)) & "  de  " & CStr(Year(DTPFechaLimite.Value)) & "'"
    
    crTraspaso.ReportFileName = App.Path & "\Reportes\ComproModPpto.rpt"
    'crTraspaso.ReportFileName = "c:\Sistemas\Reportes\ComproModPpto.rpt"
    
    iResult = crTraspaso.PrintReport
    If iResult <> 0 Then
     MsgBox crTraspaso.LastErrorNumber & "   " & crTraspaso.LastErrorString
    End If

End Sub

Private Sub BtnEliminar_Click()
    MsgBox "No se puede Eliminar el formulado, cuando ya existe una Adición o Transferencia ..."
End Sub

Private Sub BtnGrabarA_Click()
    Text9.Visible = True
    txt_monto_new.Enabled = True
    txt_monto_total.Enabled = True
    solo_form = "N"
    'Valida ingreso de datos
    If dtcTipoA <> "" Then
        varTipo = dtcTipoA
    Else
        MsgBox "Error: Por favor elija el 'Tipo de Registro' ...."
        Exit Sub
    End If
    If TxtRes <> "" Then
        varRes = TxtRes
    Else
        MsgBox "Error: Por favor registre el 'Nro. de Resolución' ...."
        Exit Sub
    End If
    If dtcFteA <> "" Then
        varFte = dtcFteA
    Else
        MsgBox "Error: Por favor elija la 'Fuente de Financiamiento' ...."
        Exit Sub
    End If
    If DtcOrgA <> "" Then
        varorg = DtcOrgA
    Else
        MsgBox "Error: Por favor elija el 'Organismo Financiador' ...."
        Exit Sub
    End If
    If dtcPryA <> "" Or dtcProA <> "" Or dtcActA <> "" Then
        varpro = dtcProA
        varpry = dtcPryA
        varAct = dtcActA
    Else
        MsgBox "Error: Por favor elija el 'Proyecto o Actividad' ...."
        Exit Sub
    End If
    If dtcParA <> "" Then
        varpar = dtcParA
    Else
        MsgBox "Error: Por favor elija la 'Partida del Gasto' ...."
        Exit Sub
    End If
    If txtmontoOrigen <> "" Then
        'Or Val(txtmontoOrigen) >= 0
        varmontoO = Val(txtmontoOrigen) + Val(txt_monto_new)
    Else
        MsgBox "Error: Por favor registre el correctamente el 'Monto de Transacción Bs' ...."
        Exit Sub
    End If
    If swgraba = "A" Then
        varNro = adoAdicion.Recordset.RecordCount
    Else
        varNro = Text9.Text
    End If
    If swgraba = "A" Then
        parametro = "fv_formulacion_trn.tipo_transaccion" + " = " + "'" + varTipo + "'" + " and " + "fv_formulacion_trn.fte_codigo" + " = " + "'" + varFte + "'" + " and " + "fv_formulacion_trn.org_codigo" + " = " + "'" + varorg + "'" + " and " + "fv_formulacion_trn.pro_proyecto" + " = " + "'" + varpry + "'" + " and " + "fv_formulacion_trn.par_codigo" + " = " + "'" + varpar + "'"
    Else
        parametro = "fv_formulacion_trn.tipo_transaccion" + " = " + "'" + varTipo + "'" + " and " + "fv_formulacion_trn.nro_transaccion" + " = " + "'" + varNro + "'"
    End If
    Call abrir_adicion                   'Abrir Adicion o Reducion
    If rsAdicion.RecordCount > 0 Then
       ' COMENTA POR AHORA  ***************************
       If swgraba = "A" Then
            MsgBox "La estructura presupuestaria ya fue registrada como Adición ..."
            adoAdicion.Recordset.CancelUpdate
       Else
          'Modifica una Adicion
          parametro = "fv_formulacion_gasto.fte_codigo" + " = " + "'" + varFte + "'" + " and " + "fv_formulacion_gasto.org_codigo" + " = " + "'" + varorg + "'" + " and " + "fv_formulacion_gasto.pro_proyecto" + " = " + "'" + varpry + "'" + " and " + "fv_formulacion_gasto.par_codigo" + " = " + "'" + varpar + "'"
          Call abrir_formulacion
          If rsformulacion.RecordCount > 0 Then
            Call graba_origen
            adoAdicion.Recordset("tipo_transaccion") = varTipo
            adoAdicion.Recordset("uni_codigo") = "01"
            adoAdicion.Recordset("pro_programa") = varpro
            adoAdicion.Recordset("pro_proyecto") = varpry
            adoAdicion.Recordset("pro_actividad") = varAct
            adoAdicion.Recordset("fte_codigo") = varFte
            adoAdicion.Recordset("org_codigo") = varorg
            adoAdicion.Recordset("par_codigo") = varpar
            adoAdicion.Recordset("ent_codigo") = "0000"
            adoAdicion.Recordset("trn_monto_origen") = varmontoO
            adoAdicion.Recordset("resolucion") = varRes
            adoAdicion.Recordset("fecha_transaccion") = Date
            adoAdicion.Recordset("usr_usuario") = GlUsuario
            adoAdicion.Recordset("fecha_registro") = Date
            adoAdicion.Recordset("hora_registro") = Format(Time, "hh:mm:ss")
            adoAdicion.Recordset.Update
          End If
       ' COMENTA POR AHORA  ***************************
       End If
    Else
      parametro = "fv_formulacion_gasto.fte_codigo" + " = " + "'" + varFte + "'" + " and " + "fv_formulacion_gasto.org_codigo" + " = " + "'" + varorg + "'" + " and " + "fv_formulacion_gasto.pro_proyecto" + " = " + "'" + varpry + "'" + " and " + "fv_formulacion_gasto.par_codigo" + " = " + "'" + varpar + "'"
      Call abrir_formulacion
       
      If rsformulacion.RecordCount > 0 Then
        If swgraba = "A" Then
            solo_form = "S"
            MsgBox "Atención: Se adicionará o reducirá el monto de una estructura presupuestaria ya Formulada ..."
            Call graba_origen
            adoAdicion.Recordset.AddNew
            adoAdicion.Recordset("nro_transaccion") = varNro
            adoAdicion.Recordset("tipo_transaccion") = varTipo
            adoAdicion.Recordset("uni_codigo") = "01"
            adoAdicion.Recordset("pro_programa") = varpro
            adoAdicion.Recordset("pro_proyecto") = varpry
            adoAdicion.Recordset("pro_actividad") = varAct
            adoAdicion.Recordset("fte_codigo") = varFte
            adoAdicion.Recordset("org_codigo") = varorg
            adoAdicion.Recordset("par_codigo") = varpar
            adoAdicion.Recordset("ent_codigo") = "000"
            If varTipo = "R" Then
                adoAdicion.Recordset("trn_monto_origen") = varmontoO * (-1)
            Else
                adoAdicion.Recordset("trn_monto_origen") = varmontoO
            End If
            adoAdicion.Recordset("resolucion") = varRes
            adoAdicion.Recordset("fecha_transaccion") = Date
            adoAdicion.Recordset("usr_usuario") = GlUsuario
            adoAdicion.Recordset("fecha_registro") = Date
            adoAdicion.Recordset("hora_registro") = Format(Time, "hh:mm:ss")
            adoAdicion.Recordset.Update
        Else
            'Modifica una Adicion
            Call graba_origen
            adoAdicion.Recordset("tipo_transaccion") = varTipo
            adoAdicion.Recordset("uni_codigo") = "01"
            adoAdicion.Recordset("pro_programa") = varpro
            adoAdicion.Recordset("pro_proyecto") = varpry
            adoAdicion.Recordset("pro_actividad") = varAct
            adoAdicion.Recordset("fte_codigo") = varFte
            adoAdicion.Recordset("org_codigo") = varorg
            adoAdicion.Recordset("par_codigo") = varpar
            adoAdicion.Recordset("ent_codigo") = "0000"
            adoAdicion.Recordset("trn_monto_origen") = varmontoO
            adoAdicion.Recordset("resolucion") = varRes
            adoAdicion.Recordset("fecha_transaccion") = Date
            adoAdicion.Recordset("usr_usuario") = GlUsuario
            adoAdicion.Recordset("fecha_registro") = Date
            adoAdicion.Recordset("hora_registro") = Format(Time, "hh:mm:ss")
            adoAdicion.Recordset.Update
        End If
      Else
        ' Registro nuevo Adición y Formulado
        Call graba_origen
        adoAdicion.Recordset.AddNew
        adoAdicion.Recordset("nro_transaccion") = varNro
        adoAdicion.Recordset("tipo_transaccion") = varTipo
        adoAdicion.Recordset("uni_codigo") = "01"
        adoAdicion.Recordset("pro_programa") = varpro
        adoAdicion.Recordset("pro_proyecto") = varpry
        adoAdicion.Recordset("pro_actividad") = varAct
        adoAdicion.Recordset("fte_codigo") = varFte
        adoAdicion.Recordset("org_codigo") = varorg
        adoAdicion.Recordset("par_codigo") = varpar
        adoAdicion.Recordset("ent_codigo") = "000"
        adoAdicion.Recordset("trn_monto_origen") = varmontoO
        adoAdicion.Recordset("resolucion") = varRes
        adoAdicion.Recordset("fecha_transaccion") = Date
        adoAdicion.Recordset("usr_usuario") = GlUsuario
        adoAdicion.Recordset("fecha_registro") = Date
        adoAdicion.Recordset("hora_registro") = Format(Time, "hh:mm:ss")
        adoAdicion.Recordset.Update
      End If
    End If
    parametro = "fv_formulacion_trn.tipo_transaccion" + " = " + "'A'" + " or " + "fv_formulacion_trn.tipo_transaccion" + " = " + "'R'"
    Call abrir_adicion                   'Abrir Adicion o Reducion
    
    fraprincipalAd.Visible = True
    fragrabarAd.Visible = False
    Frame1.Enabled = False
    solo_form = "N"
    Call Objetos_Ad
End Sub

Private Sub BtnGrabarT_Click()
  Text6.Visible = True
  If Text5.Text = 0 Then
        MsgBox "Ingrese monto para realizar el traspaso ..."
        Text5.SetFocus
  Else
  If dtcTipoT <> "" And TxtResT <> "" Then
    varTipo = dtcTipoT
    varRes = TxtResT
  Else
    MsgBox "Ingrese correctamente Tipo de Registro y/o Resolución ..."
    Exit Sub
  End If
  If dtcFteT <> "" And DtcOrgT <> "" And dtcProT <> "" And dtcPryT <> "" And dtcActT <> "" And dtcParT <> "" Then
    varFte = dtcFteT
    varorg = DtcOrgT
    varpro = dtcProT
    varpry = dtcPryT
    varAct = dtcActT
    varpar = dtcParT
    varmontoO = txtmontoOrigenT
  Else
    MsgBox "Ingrese correctamente los datos del Origen ..."
    Exit Sub
  End If
  If dtcFteT_des <> "" And DtcOrgT_des <> "" And dtcProT_des <> "" And dtcPryT_des <> "" And dtcActT_des <> "" And dtcParT_des <> "" Then
    varFteD = dtcFteT_des
    varorgD = DtcOrgT_des
    varproD = dtcProT_des
    varpryD = dtcPryT_des
    varActD = dtcActT_des
    varparD = dtcParT_des
    varmontoD = txtmontoDestino
  Else
    MsgBox "Ingrese correctamente los datos del Destino ..."
    Exit Sub
  End If
  If dtcFteT = dtcFteT_des And DtcOrgT = DtcOrgT_des And dtcPryT = dtcPryT_des And dtcParT = dtcParT_des Then
    MsgBox "Error, NO se puede realizar un Traspaso a si mismo, vuelva a intentar ..."
    Exit Sub
  End If
  
    If swgraba = "A" Then             'ADICION REGISTROS
        varNro = Adotraspaso.Recordset.RecordCount
        'Verificar el restricciones para sacar y poner
        parametro = "po_formulacion_trn.tipo_transaccion" + " = " + "'T'" + " and " + "po_formulacion_trn.org_codigo" + " = " + "'" + varorgD + "'" + " and " + "po_formulacion_trn.pro_proyecto" + " = " + "'" + varpryD + "'" + " and " + "po_formulacion_trn.par_codigo" + " = " + "'" + varparD + "'"
        Call abrir_traspaso                   'Abrir Traspaso
        If rsTraspaso.RecordCount > 0 Then
           MsgBox "No se puede sacar el presupuesto (origen), a una estructura que ya se depositó como destino ..."
           Adotraspaso.Recordset.CancelUpdate
           Exit Sub
        Else
           parametro = "fv_formulacion_gasto.org_codigo" + " = " + "'" + varorg + "'" + " and " + "fv_formulacion_gasto.pro_proyecto" + " = " + "'" + varpry + "'" + " and " + "fv_formulacion_gasto.par_codigo" + " = " + "'" + varpar + "'"
           Call abrir_formulacion
           
          If rsformulacion.RecordCount < 1 Then
            If swgraba = "A" Then
                MsgBox "No se puede Trasnferir desde una estructura presupuestaria origen inexistente, VUELVA A INTENTAR ..."
                Exit Sub
            End If
          Else
            ' Registro Transferencia
            parametro = "fv_formulacion_gasto.org_codigo" + " = " + "'" + varorg + "'" + " and " + "fv_formulacion_gasto.pro_proyecto" + " = " + "'" + varpry + "'" + " and " + "fv_formulacion_gasto.par_codigo" + " = " + "'" + varpar + "'"
            Call abrir_formulacion
            If (adoformulacion.Recordset("fgs_vigente") + varmontoO) >= 0 Then
                Call graba_origen_T
                parametro = "fv_formulacion_gasto.org_codigo" + " = " + "'" + varorgD + "'" + " and " + "fv_formulacion_gasto.pro_proyecto" + " = " + "'" + varpryD + "'" + " and " + "fv_formulacion_gasto.par_codigo" + " = " + "'" + varparD + "'"
                Call abrir_formulacion
                Call graba_destino_T
                Adotraspaso.Recordset.AddNew
                Adotraspaso.Recordset("nro_transaccion") = varNro
                Adotraspaso.Recordset("tipo_transaccion") = varTipo
                Adotraspaso.Recordset("uni_codigo") = "01"
                
                Adotraspaso.Recordset("pro_programa") = varpro
                Adotraspaso.Recordset("pro_proyecto") = varpry
                Adotraspaso.Recordset("pro_actividad") = varAct
                Adotraspaso.Recordset("fte_codigo") = varFte
                Adotraspaso.Recordset("org_codigo") = varorg
                Adotraspaso.Recordset("par_codigo") = varpar
                Adotraspaso.Recordset("ent_codigo") = "000"
                Adotraspaso.Recordset("trn_monto_origen") = varmontoO
                
                Adotraspaso.Recordset("pro_programa_des") = varproD
                Adotraspaso.Recordset("pro_proyecto_des") = varpryD
                Adotraspaso.Recordset("pro_actividad_des") = varActD
                Adotraspaso.Recordset("fte_codigo_des") = varFteD
                Adotraspaso.Recordset("org_codigo_des") = varorgD
                Adotraspaso.Recordset("par_codigo_des") = varparD
                Adotraspaso.Recordset("ent_codigo_des") = "000"
                Adotraspaso.Recordset("trn_monto_destino") = varmontoD
                
                Adotraspaso.Recordset("resolucion") = varRes
                Adotraspaso.Recordset("fecha_transaccion") = Date
                Adotraspaso.Recordset("usr_usuario") = GlUsuario
                Adotraspaso.Recordset("fecha_registro") = Date
                Adotraspaso.Recordset("hora_registro") = Format(Time, "hh:mm:ss")
    
                Adotraspaso.Recordset.Update
            Else
                MsgBox "ERROR. El monto a transferir sobrepasa el Saldo Vigente, el proceso será cancelado ... "
                Exit Sub
            End If
          End If
        End If
    End If
    
    If swgraba = "M" Then             'MODIFICACION REGISTROS
       varNro = Text6.Text
       'Verificar el restricciones para sacar y poner
       parametro = "fv_formulacion_gasto.org_codigo" + " = " + "'" + varorg + "'" + " and " + "fv_formulacion_gasto.pro_proyecto" + " = " + "'" + varpry + "'" + " and " + "fv_formulacion_gasto.par_codigo" + " = " + "'" + varpar + "'"
       Call abrir_formulacion
        
       If rsformulacion.RecordCount < 1 Then
             MsgBox "Error: Estructura presupuestaria origen inexistente ..."
       Else
         ' Registro Transferencia
         parametro = "fv_formulacion_gasto.org_codigo" + " = " + "'" + varorg + "'" + " and " + "fv_formulacion_gasto.pro_proyecto" + " = " + "'" + varpry + "'" + " and " + "fv_formulacion_gasto.par_codigo" + " = " + "'" + varpar + "'"
         Call abrir_formulacion
         If (adoformulacion.Recordset("fgs_vigente") + varmontoO) >= 0 Then
            Call graba_origen_T
            parametro = "fv_formulacion_gasto.org_codigo" + " = " + "'" + varorgD + "'" + " and " + "fv_formulacion_gasto.pro_proyecto" + " = " + "'" + varpryD + "'" + " and " + "fv_formulacion_gasto.par_codigo" + " = " + "'" + varparD + "'"
            Call abrir_formulacion
            Call graba_destino_T
            
            Adotraspaso.Recordset("nro_transaccion") = varNro
            Adotraspaso.Recordset("tipo_transaccion") = varTipo
            Adotraspaso.Recordset("uni_codigo") = "01"
            
            Adotraspaso.Recordset("pro_programa") = varpro
            Adotraspaso.Recordset("pro_proyecto") = varpry
            Adotraspaso.Recordset("pro_actividad") = varAct
            Adotraspaso.Recordset("fte_codigo") = varFte
            Adotraspaso.Recordset("org_codigo") = varorg
            Adotraspaso.Recordset("par_codigo") = varpar
            Adotraspaso.Recordset("ent_codigo") = "000"
            Adotraspaso.Recordset("trn_monto_origen") = varmontoO '* (-1)
            
            Adotraspaso.Recordset("pro_programa_des") = varproD
            Adotraspaso.Recordset("pro_proyecto_des") = varpryD
            Adotraspaso.Recordset("pro_actividad_des") = varActD
            Adotraspaso.Recordset("fte_codigo_des") = varFteD
            Adotraspaso.Recordset("org_codigo_des") = varorgD
            Adotraspaso.Recordset("par_codigo_des") = varparD
            Adotraspaso.Recordset("ent_codigo_des") = "000"
            Adotraspaso.Recordset("trn_monto_destino") = varmontoD
            Adotraspaso.Recordset("resolucion") = varRes
            Adotraspaso.Recordset("fecha_transaccion") = Date
            Adotraspaso.Recordset.Update
         Else
                MsgBox "ERROR. El monto a transferir sobrepasa el Saldo Vigente, el proceso será cancelado ... "
                Exit Sub
         End If

       End If
    End If
    parametro = "po_formulacion_trn.tipo_transaccion" + " = " + "'T'" + " or " + "po_formulacion_trn.tipo_transaccion" + " = " + "'F'"
    Call abrir_traspaso                   'Abrir Traspaso
    fraprincipalTr.Visible = True
    fragrabarTr.Visible = False
    Frame2.Enabled = False
     
    Text5.Visible = False
    Label16.Visible = False
  End If

End Sub

Private Sub BtnImprimir_Click()
'   'Dim e As Long
''    'e = Shell(App.Path & "\saf2003\Reportes\Presupuesto\ProyRepPresupuesto.exe", 1)
''
'  glRepPresup = "REP002"
'  frmRepPresupuesto.Show
End Sub

Private Sub BtnImprimirB_Click()
Dim iResult As Variant
    
    crTraspaso.ReportFileName = App.Path & "\Reportes\ADICION PRESUPUESTARIA.rpt"
    iResult = crTraspaso.PrintReport
    If iResult <> 0 Then
     MsgBox crTraspaso.LastErrorNumber & "   " & crTraspaso.LastErrorString
    End If

End Sub

Private Sub BtnModificar_Click()
    MsgBox "No se puede Modificar el formulado, cuando ya existe una Adición o Transferencia ..."
End Sub

Private Sub BtnModificarA_Click()
    swgraba = "M"
    fraprincipalAd.Visible = False
    fragrabarAd.Visible = True
    Frame1.Enabled = True
    'Desactiva Objetos
    Text9.Enabled = False
    dtcTipoA.Enabled = False
    dtcTipoDesA.Enabled = False
    dtcFteA.Enabled = False
    DtcFteDesA.Enabled = False
    DtcOrgA.Enabled = False
    DtcOrgDesA.Enabled = False
    dtcPryA.Enabled = False
    DtcPryDes.Enabled = False
    dtcParA.Enabled = False
    DtcPasDesA.Enabled = False
    
    txtmontoOrigen.Enabled = False
    txt_monto_new.Enabled = True
    txt_monto_total.Enabled = False
End Sub

Private Sub BtnModificarT_Click()
    swgraba = "M"
    fraprincipalTr.Visible = False
    fragrabarTr.Visible = True
    Frame2.Enabled = True
    
    Text5.Visible = True
    Label16.Visible = True
    Frame2.Enabled = False
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub BtnSalirA_Click()
    Unload Me
End Sub

Private Sub BtnSalirT_Click()
    Unload Me
End Sub

Private Sub dtcFteA_Click(Area As Integer)
   DtcFteDesA.BoundText = dtcFteA.BoundText
   Call pOrganismo(DtcFteDesA.BoundText)
End Sub

Private Sub DtcFteDesA_Click(Area As Integer)
    dtcFteA.BoundText = DtcFteDesA.BoundText
    Call pOrganismo(dtcFteA.BoundText)
End Sub

Private Sub DtcOrgA_Click(Area As Integer)
'    DtcOrg.BoundText = DtcDesOrg.BoundText
'    Call pConv(DtcOrg.BoundText)
End Sub

Private Sub DtcOrgDesA_Click(Area As Integer)
'DtcOrg.BoundText = DtcDesOrg.BoundText
'    Call pConv(DtcOrg.BoundText)
End Sub

Private Sub Form_Load()
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    parametro = "fv_formulacion_gasto.ges_gestion" + " <> " + "'2008'"
    Call abrir_tablas
    Call abrir_formulacion
    'varnroF = fv_formulacion_gasto.Recordset.RecordCount
    Call FrmFormulacion.Totales
    FrmFormulacion.lblFormulado = Format(montoTotal, "###,###,##0")
    FrmFormulacion.lblAdiciones = Format(montoTotalA, "###,###,##0")
    FrmFormulacion.lblModificaciones = Format(montoTotalM, "###,###,##0")
    FrmFormulacion.lblVigente = Format((montoTotal + montoTotalA + montoTotalM), "###,###,##0")
   
	Call SeguridadSet(Me)
End Sub

Public Sub abrir_formulacion()
  Set rsformulacion = New ADODB.Recordset       'Abrir fv_formulacion_gasto
    If rsformulacion.State = 1 Then rsformulacion.Close
    rsformulacion.Open "select * from fv_formulacion_gasto where " & parametro & " order by org_codigo, pro_proyecto, par_codigo ", db, adOpenDynamic, adLockOptimistic
    If rsformulacion.RecordCount > 0 Then
        Set adoformulacion.Recordset = rsformulacion
        Set Dtgformulacion.DataSource = adoformulacion.Recordset
    Else
        Set rsNada = New ADODB.Recordset
        Set adoformulacion.Recordset = rsformulacion
        Set Dtgformulacion.DataSource = rsNada
    End If
End Sub

Public Sub abrir_adicion()
    Set rsAdicion = New ADODB.Recordset           'Abrir fo_formulacion_trn
    If rsAdicion.State = 1 Then rsAdicion.Close
    rsAdicion.Open "select * from fv_formulacion_trn where " & parametro & " order by nro_transaccion ", db, adOpenDynamic, adLockOptimistic
    If rsAdicion.RecordCount > 0 Then
            Set adoAdicion.Recordset = rsAdicion
            Set dtgAdicion.DataSource = adoAdicion.Recordset
    Else
        Set rsNada = New ADODB.Recordset
        Set adoAdicion.Recordset = rsAdicion
        Set dtgAdicion.DataSource = rsNada
    End If
End Sub

Public Sub abrir_traspaso()
    Set rsTraspaso = New ADODB.Recordset           'Abrir fo_formulacion_trn
    If rsTraspaso.State = 1 Then rsTraspaso.Close
    'rsTraspaso.Open "select * from fo_formulacion_trn where " & parametro & " order by nro_transaccion ", db, adOpenDynamic, adLockOptimistic
    rsTraspaso.Open "select * from po_formulacion_trn where " & parametro & " order by nro_transaccion ", db, adOpenDynamic, adLockOptimistic
    If rsTraspaso.RecordCount > 0 Then
            Set Adotraspaso.Recordset = rsTraspaso
            Set dtgTraspaso.DataSource = Adotraspaso.Recordset
    Else
        Set rsNada = New ADODB.Recordset
        Set Adotraspaso.Recordset = rsTraspaso
        Set dtgTraspaso.DataSource = rsNada
    End If
End Sub

Private Sub abrir_tablas()
    Set rsfuente = New ADODB.Recordset       ' Fuente de Financiamiento
    If rsfuente.State = 1 Then rsfuente.Close
    rsfuente.Open "select * from fc_fuente_financiamiento  ", db, adOpenDynamic, adLockOptimistic
    If rsfuente.RecordCount > 0 Then
        Set Adofuente.Recordset = rsfuente
    End If
    
    Set rsOrganismo = New ADODB.Recordset       ' Organismo de Financiamiento
    If rsOrganismo.State = 1 Then rsOrganismo.Close
    rsOrganismo.Open "select * from fc_organismo_financiamiento  ", db, adOpenDynamic, adLockOptimistic
    If rsOrganismo.RecordCount > 0 Then
        Set adoorganismo.Recordset = rsOrganismo
    End If
    
    Set rsproyecto = New ADODB.Recordset       ' Categoría Programática
    If rsproyecto.State = 1 Then rsproyecto.Close
    rsproyecto.Open "select * from fc_estructura_programatica  ", db, adOpenDynamic, adLockOptimistic
    If rsproyecto.RecordCount > 0 Then
        Set adoproyecto.Recordset = rsproyecto
    End If
    
    Set rspartida = New ADODB.Recordset       ' Organismo de Financiamiento
    If rspartida.State = 1 Then rspartida.Close
    rspartida.Open "select * from fc_partida_gasto  ", db, adOpenDynamic, adLockOptimistic
    If rspartida.RecordCount > 0 Then
        Set Adopartida.Recordset = rspartida
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub sstab1_Click(PreviousTab As Integer)
  If FrmFormulacion.sstab1.Tab = 0 Then        ' Formulacion
    parametro = "fv_formulacion_gasto.ges_gestion" + " = " + "'2004'"
    Call abrir_formulacion
    
    Call FrmFormulacion.Totales
    FrmFormulacion.lblFormulado = Format(montoTotal, "###,###,##0")
    FrmFormulacion.lblAdiciones = Format(montoTotalA, "###,###,##0")
    FrmFormulacion.lblModificaciones = Format(montoTotalM, "###,###,##0")
    FrmFormulacion.lblVigente = Format((montoTotal + montoTotalA + montoTotalM), "###,###,##0")
    
  End If
  
  If FrmFormulacion.sstab1.Tab = 1 Then        ' Adiciones o Reducciones
    parametro = "fv_formulacion_trn.tipo_transaccion" + " = " + "'A'" + " or " + "fv_formulacion_trn.tipo_transaccion" + " = " + "'R'"
    Call abrir_adicion
    Frame1.Enabled = False
    fraprincipalAd.Visible = True
    fragrabarAd.Visible = False
    
    Call totalesA
    FrmFormulacion.lblAdiciones2 = Format(montoTotalA2, "###,###,##0")
    
    tipoT = "fc_tipo_transaccion.estado_transaccion" + " = " + "'A'"
    Call abrir_tipo
  End If

End Sub

Private Sub meses(mes)
    Select Case mes
    Case 1
        mes = "enero"
    Case 2
        mes = "febrero"
    Case 3
        mes = "marzo"
    Case 4
        mes = "abril"
    Case 5
        mes = "mayo"
    Case 6
        mes = "junio"
    Case 7
        mes = "julio"
    Case 8
        mes = "agosto"
    Case 9
        mes = "septiembre"
    Case 10
        mes = "octubre"
    Case 11
        mes = "noviembre"
    Case 12
        mes = "diciembre"
    Case Else
         MsgBox "seleccione otro color"
      
  End Select

End Sub

Public Sub Totales()
'      Dim RsDevenga As ADODB.Recordset
'      Dim RsCompro As ADODB.Recordset
      Dim GlSqlAux As String
'      Set RsDevenga = New ADODB.Recordset
      Set RsCompro = New ADODB.Recordset
      
'      ' Para ACCESS
'    GlSQLAux = "SELECT IIF(ISNULL(SUM(fgs_formulado)), 0, SUM(fgs_formulado)) AS TotalFormulado, " & _
'                "IIF(ISNULL(SUM(fgs_adiciones)), 0, SUM(fgs_adiciones)) AS TotalAdiciones, " & _
'                "IIF(ISNULL(SUM(fgs_modificaciones)), 0, SUM(fgs_modificaciones)) AS TotalModificaciones " & _
'                 "FROM fv_formulacion_gasto " & _
'                 "WHERE " & parametro & " "
                 
        ' Para SQL
    GlSqlAux = "SELECT ISNULL(SUM(fgs_formulado), 0) AS TotalFormulado, " & _
                "ISNULL(SUM(fgs_adiciones), 0) AS TotalAdiciones, " & _
                "ISNULL(SUM(fgs_modificaciones), 0) AS TotalModificaciones " & _
                 "FROM fv_formulacion_gasto " & _
                 "WHERE " & parametro & " "
                     
                 '"IIF(ISNULL(SUM(fgs_vigente)), 0, SUM(fgs_vigente)) AS TotalVigente " & _

'      ' No sirve
'      GlSQLAux = "SELECT ISNULL(SUM(monto_Total), 0) AS TotalDevengado " & _
'                 "FROM pagos, pago_Detalle " & _
'                 "WHERE (pagos.codigo_pago = pago_detalle.codigo_pago) AND (pagos.Tipo_formulario = 'DEV') AND (pagos.estado_devengado = 'S') AND (pagos.Nro_Comprobante_Anterior = '" & AdoRegularizacion.Recordset!Nro_Comprobante_Anterior & "')"
'      RsDevenga.Open GlSQLAux, db, adOpenStatic
'      GlSQLAux = "SELECT Sum(Monto_Total) AS MontoTotal FROM pago_detalle " & _
'                 "WHERE pago_detalle.Codigo_Pago = " & AdoRegularizacion.Recordset!Nro_Comprobante_Anterior & " "

        
      RsCompro.Open GlSqlAux, db, adOpenStatic
      montoTotal = RsCompro!TotalFormulado
      montoTotalA = RsCompro!TotalAdiciones
      montoTotalM = RsCompro!TotalModificaciones
      'montoTotalV = RsCompro!TotalVigente
      
'      If (RsCompro!MontoTotal < RsDevenga!TotalDevengado + rsDet("monto_total")) Then
'        MsgBox "La Suma de lo DEVENGADO excede el Monto del Compromiso del Comprobante '" & AdoRegularizacion.Recordset!Nro_Comprobante_Anterior & "'.", vbExclamation + vbOKOnly, "ERROR" '"La estructura presupuestaria NO es válida o NO EXISTE PRESUPUESTO "
'        Exit Sub
'      Else
'        rsPpto("fgs_devengado") = rsPpto("fgs_devengado") + rsDet("monto_total")
'        rsPpto.Update
'      End If

End Sub

Private Sub graba_origen()
    If swgraba = "A" Then
        If solo_form <> "S" Then
            adoformulacion.Recordset.AddNew
        End If
    End If
    If solo_form <> "S" Then
        adoformulacion.Recordset("ges_gestion") = Year(Date)
        adoformulacion.Recordset("uni_codigo") = "01"
        adoformulacion.Recordset("pro_programa") = varpro
        adoformulacion.Recordset("pro_proyecto") = varpry
        adoformulacion.Recordset("pro_actividad") = varAct
        adoformulacion.Recordset("fte_codigo") = varFte
        adoformulacion.Recordset("org_codigo") = varorg
        adoformulacion.Recordset("par_codigo") = varpar
        adoformulacion.Recordset("ent_codigo") = "000"
        adoformulacion.Recordset("fgs_formulado") = IIf(IsNull(adoformulacion.Recordset("fgs_formulado")), 0, adoformulacion.Recordset("fgs_formulado"))
    End If
    If varTipo = "A" Then
'        adoformulacion.Recordset("fgs_adiciones") = Val(varmontoO + IIf(IsNull(adoformulacion.Recordset("fgs_adiciones")), 0, adoformulacion.Recordset("fgs_adiciones")))
'        adoformulacion.Recordset("fgs_adicion") = Val(varmontoO + IIf(IsNull(adoformulacion.Recordset("fgs_adicion")), 0, adoformulacion.Recordset("fgs_adicion")))
        adoformulacion.Recordset("fgs_adiciones") = Val(varmontoO)
        adoformulacion.Recordset("fgs_adicion") = Val(varmontoO)
        adoformulacion.Recordset("estado_adicion") = "S"
    End If
    If varTipo = "R" Then
        adoformulacion.Recordset("fgs_adiciones") = Val(varmontoO * (-1) + IIf(IsNull(adoformulacion.Recordset("fgs_adiciones")), 0, adoformulacion.Recordset("fgs_adiciones")))
        adoformulacion.Recordset("fgs_adicion") = Val(varmontoO * (-1) + IIf(IsNull(adoformulacion.Recordset("fgs_adiciones")), 0, adoformulacion.Recordset("fgs_adiciones")))
        'adoformulacion.Recordset("fgs_adiciones") = varmontoO * (-1) + adoformulacion.Recordset("fgs_adiciones")
        'adoformulacion.Recordset("fgs_adicion") = varmontoO * (-1) + adoformulacion.Recordset("fgs_adicion")
        adoformulacion.Recordset("estado_adicion") = "S"
    End If
    If solo_form <> "S" Then
        adoformulacion.Recordset("fgs_modificaciones") = IIf(IsNull(adoformulacion.Recordset("fgs_modificaciones")), 0, adoformulacion.Recordset("fgs_modificaciones"))
    End If
    adoformulacion.Recordset("fgs_vigente") = adoformulacion.Recordset("fgs_formulado") + adoformulacion.Recordset("fgs_adiciones") + adoformulacion.Recordset("fgs_modificaciones")
    
    adoformulacion.Recordset("nro_transaccion") = varNro
    If varTipo = "A" Then
        adoformulacion.Recordset("fgs_adicion_techo") = varNro
    End If
    adoformulacion.Recordset("tipo_transaccion") = varTipo
    adoformulacion.Recordset("fecha_formulacion") = Date
    adoformulacion.Recordset("usr_usuario") = GlUsuario
    adoformulacion.Recordset("fecha_registro") = Date
    adoformulacion.Recordset("hora_registro") = Format(Time, "hh:mm:ss")
    adoformulacion.Recordset.Update
    
End Sub

Private Sub graba_origen_T()
  If swgraba = "A" Then
    If varTipo = "T" Or varTipo = "F" Then
        If adoformulacion.Recordset("fgs_modificaciones") <> 0 Then
            adoformulacion.Recordset("fgs_modificaciones") = adoformulacion.Recordset("fgs_modificaciones") + varmontoO
        Else
            adoformulacion.Recordset("fgs_modificaciones") = varmontoO
        End If
        adoformulacion.Recordset("estado_origen") = "S"
    End If
  Else
    If varTipo = "T" Or varTipo = "F" Then
        adoformulacion.Recordset("fgs_modificaciones") = varmontoO
        adoformulacion.Recordset("estado_origen") = "S"
    End If
  End If
    'adoformulacion.Recordset("fgs_vigente") = adoformulacion.Recordset("fgs_formulado") + adoformulacion.Recordset("fgs_adiciones") + adoformulacion.Recordset("fgs_modificaciones")
    adoformulacion.Recordset("fgs_vigente") = adoformulacion.Recordset("fgs_formulado") + adoformulacion.Recordset("fgs_modificaciones")
    adoformulacion.Recordset("nro_transaccion") = varNro
    adoformulacion.Recordset("tipo_transaccion") = varTipo
    adoformulacion.Recordset("fecha_formulacion") = Date
    adoformulacion.Recordset.Update

End Sub

Private Sub graba_destino_T()
  If swgraba = "A" Then
    If varTipo = "T" Or varTipo = "F" Then
        If adoformulacion.Recordset("fgs_modificaciones") <> 0 Then
            adoformulacion.Recordset("fgs_modificaciones") = adoformulacion.Recordset("fgs_modificaciones") + varmontoD
        Else
            adoformulacion.Recordset("fgs_modificaciones") = varmontoD
        End If
        adoformulacion.Recordset("estado_destino") = "S"
    End If
  Else
    If varTipo = "T" Or varTipo = "F" Then
        adoformulacion.Recordset("fgs_modificaciones") = varmontoD
        adoformulacion.Recordset("estado_destino") = "S"
    End If
  End If
    'adoformulacion.Recordset("fgs_vigente") = adoformulacion.Recordset("fgs_formulado") + adoformulacion.Recordset("fgs_adiciones") + adoformulacion.Recordset("fgs_modificaciones")
    adoformulacion.Recordset("fgs_vigente") = adoformulacion.Recordset("fgs_formulado") + adoformulacion.Recordset("fgs_modificaciones")
    adoformulacion.Recordset("nro_transaccion") = varNro
    adoformulacion.Recordset("tipo_transaccion") = varTipo
    adoformulacion.Recordset("fecha_formulacion") = Date
    adoformulacion.Recordset.Update

End Sub

Private Sub SSTab3_Click(PreviousTab As Integer)
  If FrmFormulacion.SSTab3.Tab = 0 Then        ' Tipo - A
    parametro = "fv_formulacion_trn.tipo_transaccion" + " = " + "'A'" + " or " + "fv_formulacion_trn.tipo_transaccion" + " = " + "'R'"
    Call abrir_adicion
    tipoT = "fc_tipo_transaccion.estado_transaccion" + " = " + "'A'"
    Call abrir_tipo
  End If
  
  If FrmFormulacion.SSTab3.Tab = 1 Then        ' Tipo - T
    parametro = "po_formulacion_trn.tipo_transaccion" + " = " + "'T'" + " or " + "po_formulacion_trn.tipo_transaccion" + " = " + "'F'"
    Call abrir_traspaso
    fraprincipalTr.Visible = True
    fragrabarTr.Visible = False
    Frame2.Enabled = False
    tipoT = "fc_tipo_transaccion.estado_transaccion" + " = " + "'T'"
    Call abrir_tipo
    
    Text5.Visible = False
    Label16.Visible = False
  End If
End Sub

Public Sub abrir_tipo()
    Set rsTipo = New ADODB.Recordset           'Abrir fc_tipo_transaccion
    If rsTipo.State = 1 Then rsTipo.Close
    rsTipo.Open "select * from fc_tipo_transaccion where " & tipoT & " order by tipo_transaccion ", db, adOpenDynamic, adLockOptimistic
    If rsTipo.RecordCount > 0 Then
        Set AdoTipo.Recordset = rsTipo
    End If

End Sub

Private Sub Text5_LostFocus()
    If Text5.Text = 0 Then
        MsgBox "Ingrese monto para realizar el traspaso ..."
        Text5.SetFocus
    Else
        Frame2.Enabled = True
        txtmontoOrigenT.Enabled = True
        txtmontoDestino.Enabled = True
        txtmontoOrigenT = CDbl(Text5.Text) * (-1)
        txtmontoDestino = CDbl(Text5.Text)
        txtmontoOrigenT.Enabled = False
        txtmontoDestino.Enabled = False
        Frame2.Enabled = False
    End If
End Sub

Public Sub totalesA()
      Dim GlSqlAux As String
      Set RsCompro = New ADODB.Recordset
      'Access
'      GlSQLAux = "SELECT IIF(ISNULL(SUM(trn_monto_origen)), 0, SUM(trn_monto_origen)) AS TotalAdiciones2 " & _
'                 "FROM fv_formulacion_trn " & _
'                 "WHERE " & parametro & " "
      'SQL
      GlSqlAux = "SELECT ISNULL(SUM(trn_monto_origen), 0) AS TotalAdiciones2 " & _
                 "FROM fv_formulacion_trn " & _
                 "WHERE " & parametro & " "
                 
      RsCompro.Open GlSqlAux, db, adOpenStatic
      montoTotalA2 = RsCompro!TotalAdiciones2
      
End Sub

Private Sub pOrganismo(CodFuente As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from fc_organismo_financiamiento where fte_codigo='" & CodFuente & "'"
   
   Set DtcOrgA.RowSource = Nothing
   Set DtcOrgA.RowSource = db.Execute(strConsultaF, , adCmdText)
   DtcOrgA.ReFill
   DtcOrgA.BoundText = Empty
   
   Set DtcOrgDesA.RowSource = Nothing
   Set DtcOrgDesA.RowSource = db.Execute(strConsultaF, , adCmdText)
   DtcOrgDesA.ReFill
   DtcOrgDesA.BoundText = Empty

End Sub

Private Sub Objetos_Ad()
'Desactiva Objetos
    Text9.Enabled = True
    dtcTipoA.Enabled = True
    dtcTipoDesA.Enabled = True
    dtcFteA.Enabled = True
    DtcFteDesA.Enabled = True
    DtcOrgA.Enabled = True
    DtcOrgDesA.Enabled = True
    dtcPryA.Enabled = True
    DtcPryDes.Enabled = True
    dtcParA.Enabled = True
    DtcPasDesA.Enabled = True
End Sub

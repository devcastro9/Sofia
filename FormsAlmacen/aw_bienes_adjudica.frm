VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form aw_bienes_adjudica 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COMEX - Compra de Servicios - Adjudicación"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H80000006&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   9315
      TabIndex        =   35
      Top             =   0
      Width           =   9375
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "aw_bienes_adjudica.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   37
         Top             =   120
         Width           =   1335
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1680
         Picture         =   "aw_bienes_adjudica.frx":07D6
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   36
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADJUDICACION POR PROVEEDOR"
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
         Left            =   3945
         TabIndex        =   38
         Top             =   240
         Width           =   4065
      End
   End
   Begin MSAdodcLib.Adodc Ado_clasif1 
      Height          =   330
      Left            =   360
      Top             =   7200
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
      Caption         =   "Ado_clasif1"
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
   Begin MSAdodcLib.Adodc Ado_clasif2 
      Height          =   330
      Left            =   2520
      Top             =   7200
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
      Caption         =   "Ado_clasif2"
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
   Begin MSAdodcLib.Adodc Ado_clasif3 
      Height          =   330
      Left            =   4680
      Top             =   7200
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
      Caption         =   "Ado_clasif3"
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
   Begin MSAdodcLib.Adodc Ado_clasif4 
      Height          =   330
      Left            =   360
      Top             =   6840
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
      Caption         =   "Ado_clasif4"
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
   Begin MSAdodcLib.Adodc Ado_clasif5 
      Height          =   330
      Left            =   2520
      Top             =   6840
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
      Caption         =   "Ado_clasif5"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   9135
      Begin VB.ComboBox cmb_tipomoneda 
         DataField       =   "tipo_moneda"
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
         Height          =   315
         ItemData        =   "aw_bienes_adjudica.frx":10C2
         Left            =   5280
         List            =   "aw_bienes_adjudica.frx":10CC
         TabIndex        =   47
         Top             =   3360
         Width           =   1620
      End
      Begin VB.TextBox txt_tipocambio 
         BackColor       =   &H80000011&
         DataField       =   "tipo_cambio"
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
         Height          =   285
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   44
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txt_adjudica 
         DataField       =   "adjudica_cantidad_total"
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   42
         Top             =   3360
         Width           =   1095
      End
      Begin VB.ComboBox cmd_unimed2 
         DataField       =   "unimed_codigo_pag"
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
         Height          =   315
         ItemData        =   "aw_bienes_adjudica.frx":10DB
         Left            =   6480
         List            =   "aw_bienes_adjudica.frx":10F1
         TabIndex        =   30
         Text            =   "ANUAL"
         Top             =   5300
         Width           =   1875
      End
      Begin VB.TextBox txtCantCuota 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "cantidad_cuotas_pag"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   1
         EndProperty
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3480
         TabIndex        =   29
         Text            =   "1"
         Top             =   5300
         Width           =   1785
      End
      Begin VB.ComboBox cmb_mes_ini 
         DataField       =   "mes_inicio_crono"
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
         Height          =   315
         ItemData        =   "aw_bienes_adjudica.frx":1119
         Left            =   480
         List            =   "aw_bienes_adjudica.frx":1141
         TabIndex        =   28
         Text            =   "SEPTIEMBRE"
         Top             =   5300
         Width           =   1620
      End
      Begin VB.TextBox txt_pais 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5880
         MaxLength       =   80
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt_Nota 
         DataField       =   "nro_nota_remision"
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   21
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txt_total_dol 
         DataField       =   "adjudica_monto_dol"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
         Height          =   285
         Left            =   7200
         MaxLength       =   20
         TabIndex        =   19
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox txt_total_bs 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "adjudica_monto_bs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7080
         MaxLength       =   20
         TabIndex        =   18
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "PROVEEDOR"
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   8685
         Begin VB.TextBox Text5 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   8055
            TabIndex        =   26
            Top             =   1690
            Width           =   375
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   6465
            TabIndex        =   25
            Top             =   1090
            Width           =   260
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   2150
            TabIndex        =   24
            Top             =   1090
            Width           =   260
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "aw_bienes_adjudica.frx":11AA
            DataField       =   "beneficiario_codigo"
            DataSource      =   "aw_compra_bienes.Ado_detalle4"
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "aw_bienes_adjudica.frx":11C4
            DataField       =   "beneficiario_codigo"
            DataSource      =   "aw_compra_bienes.Ado_detalle4"
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   741
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_aux4 
            Bindings        =   "aw_bienes_adjudica.frx":11DE
            DataField       =   "beneficiario_codigo"
            DataSource      =   "aw_compra_bienes.Ado_detalle4"
            Height          =   315
            Left            =   3240
            TabIndex        =   9
            Top             =   1080
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   741
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_telefono_Cel"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_aux5 
            Bindings        =   "aw_bienes_adjudica.frx":11F8
            DataField       =   "beneficiario_codigo"
            DataSource      =   "aw_compra_bienes.Ado_detalle4"
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   1680
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   741
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_domicilio_legal"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblbien 
            BackColor       =   &H00000000&
            Caption         =   "Pais"
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
            Height          =   195
            Index           =   1
            Left            =   7080
            TabIndex        =   34
            Top             =   840
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label lblprov 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "NIT/CI Proveedor"
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
            Height          =   195
            Left            =   165
            TabIndex        =   23
            Top             =   840
            Width           =   1530
         End
         Begin VB.Label lblbien 
            BackColor       =   &H00000000&
            Caption         =   "Teléfonos"
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
            Height          =   195
            Index           =   11
            Left            =   3240
            TabIndex        =   22
            Top             =   840
            Width           =   1050
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Dirección"
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
            Height          =   195
            Index           =   6
            Left            =   165
            TabIndex        =   11
            Top             =   1440
            Width           =   810
         End
      End
      Begin VB.TextBox txt_campo1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         DataField       =   "unidad_codigo"
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3240
         MaxLength       =   80
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtSW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5160
         MaxLength       =   80
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComCtl2.DTPicker txtFecha 
         DataField       =   "fecha_inicio_contrato"
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   4200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   98893825
         CurrentDate     =   42248
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txtFecha2 
         DataField       =   "fecha_fin_contrato"
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
         Height          =   315
         Left            =   2520
         TabIndex        =   14
         Top             =   4200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   98893825
         CurrentDate     =   42248
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txtFecha3 
         DataField       =   "fecha_recibe_almacen"
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
         Height          =   315
         Left            =   4755
         TabIndex        =   15
         Top             =   4200
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   98893825
         CurrentDate     =   42248
         MinDate         =   32874
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tipo Moneda"
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
         Height          =   195
         Index           =   5
         Left            =   5280
         TabIndex        =   46
         Top             =   3120
         Width           =   1125
      End
      Begin VB.Label lbladjudica 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tipo Cambio"
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
         Height          =   195
         Index           =   0
         Left            =   3360
         TabIndex        =   45
         Top             =   3120
         Width           =   1065
      End
      Begin VB.Label lbladjudica 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cantidad Total"
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
         Height          =   195
         Index           =   5
         Left            =   1800
         TabIndex        =   43
         Top             =   3075
         Width           =   1260
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha.Salida.de.Fabrica"
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
         Height          =   195
         Index           =   4
         Left            =   4680
         TabIndex        =   41
         Top             =   3960
         Width           =   2085
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha.Fin.Fabricacion"
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
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   40
         Top             =   3960
         Width           =   1905
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha.Inicio.Fabricacion"
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
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   39
         Top             =   3960
         Width           =   2115
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   0
         X2              =   9110
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Periodicidad.de.Pago"
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
         Height          =   195
         Left            =   6450
         TabIndex        =   33
         Top             =   5040
         Width           =   1845
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Cuotas"
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
         Height          =   195
         Left            =   3480
         TabIndex        =   32
         Top             =   5025
         Width           =   960
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes.Inicio.Pago"
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
         Height          =   195
         Left            =   480
         TabIndex        =   31
         Top             =   5025
         Width           =   1380
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro.Proforma"
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
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   3075
         Width           =   1125
      End
      Begin VB.Label lbl_adjudica 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "adjudica_codigo"
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
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
         Left            =   7920
         TabIndex        =   17
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Gestion"
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
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Monto USD"
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
         Height          =   195
         Left            =   7080
         TabIndex        =   13
         Top             =   3915
         Width           =   990
      End
      Begin VB.Label lbl_campo2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Monto  Bs"
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
         Height          =   195
         Left            =   7080
         TabIndex        =   12
         Top             =   3075
         Width           =   870
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro  Compra"
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
         Height          =   195
         Index           =   1
         Left            =   6795
         TabIndex        =   4
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label txtCodigo1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "ges_gestion"
         DataSource      =   "aw_compra_bienes.Ado_detalle4"
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
         Left            =   1200
         TabIndex        =   3
         Top             =   255
         Width           =   1095
      End
   End
End
Attribute VB_Name = "aw_bienes_adjudica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Para_Aceptado As String
Dim rs_clasif1 As New ADODB.Recordset
Dim rs_clasif2 As New ADODB.Recordset
Dim rs_clasif3 As New ADODB.Recordset
Dim rs_clasif4 As New ADODB.Recordset
Dim rs_clasif5 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim VAR_ADJUDICA As String

Dim VAR_OCUP, VAR_MED2, MControl As String

Dim VAR_COMPRA, CONT_MED, corrprog As Integer
Dim VAR_MES2, CONT3, CONT4, VAR_COBR2   As Integer
      
Dim FControl, FInicio As Date

Private Sub BtnCancelar_Click()
'cancela la edicion de datos
    Para_Aceptado = "N"
'    txtSW = "0"
    Unload Me
End Sub

Private Sub BtnGrabar_Click()
'acepta las modificaciones realizadas
If Valida Then
    Dim SQLS As String
    SQLS = ""
   'If txtSW = "ADD" Then
   '             '    fecha_recibe_almacen, almacen_codigo, poa_codigo, usr_codigo_aprueba, fecha_aprueba

   If swnuevo = 1 Then
      'DB.Execute "Insert INTO ro_Beneficiario_Dependiente (beneficiario_codigo, cod_dependiente, Cod_asegurado, Fecha_asegurado, fecha_nacimiento, primer_apellido, segundo_apellido, nombres, cod_pariente, nomb_pariente, estado_codigo, beneficiario_denominacion, ocupacion_pariente) Values ('" & txtBenef.Text & "', '" & txtCI.Text & "', '" & TxtItem.Text & "', '" & DTPFec_Seguro.Value & "', '" & txtNac.Value & "', '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', " & dtc_codigo1.Text & ", '" & dtc_desc1.Text & "', '" & txtEstado.Text & "', '" & nomb2 & "', '" & TxtOcupacion & "')"
      ''" & txtBenef.Caption & "',
       'DB.Execute "Insert INTO ao_solicitud_persona (ges_gestion, unidad_codigo, solicitud_codigo, benef_primer_apellido, benef_segundo_apellido, benef_nombres, benef_direccion_domicilio, benef_telefonos_ref, benef_codigo, puesto_codigo, ocup_codigo, munic_codigo, nivel_educ_codigo, observaciones, benef_fecha, estado_codigo, fecha_registro, usr_codigo) Values ('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
       '('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
      aw_compra_bienes.Ado_detalle4.Recordset("ges_gestion") = glGestion
      
  '    aw_bienes_adjudica.Ado_detalle4.Recordset("compra_codigo").Value = fw_compras_gral.Ado_datos.Recordset!compra_codigo
'      aw_compra_bienes.Ado_detalle4.Recordset("adjudica_codigo") = aw_compra_bienes.Ado_detalle4.Recordset.RecordCount
      'VAR_COMPRA = fw_compras_gral.Ado_datos13.Recordset!compra_codigo
   Else
      'DB.Execute "update ro_Beneficiario_Dependiente set beneficiario_codigo='" & txtBenef.Text & "', cod_dependiente='" & txtCI.Text & "', Cod_asegurado='" & TxtItem.Text & "', primer_apellido='" & txtPat.Text & "', segundo_apellido='" & txtMat.Text & "', nombres='" & txtNom.Text & "', cod_pariente=" & dtc_codigo1.Text & ", nomb_pariente='" & dtc_desc1.Text & "', estado_codigo='" & txtEstado.Text & "', beneficiario_denominacion='" & nomb2 & "'  "
      ' fecha_registro  hora_registro usr_usuario
      VAR_ADJUDICA = aw_compra_bienes.Ado_detalle4.Recordset("adjudica_codigo")
   End If
'    fw_compras_gral.Ado_detalle2.Recordset("adjudica_fecha").Value = Format(Date, "dd/mm/yyyy")
'    fw_compras_gral.Ado_detalle2.Recordset("proceso_codigo") = "CMX"
'    fw_compras_gral.Ado_detalle2.Recordset("subproceso_codigo") = "CMX-01"
'    fw_compras_gral.Ado_detalle2.Recordset("etapa_codigo").Value = "CMX-01-01"

'    aw_compra_bienes.Ado_detalle4.Recordset("clasif_codigo").Value = "CMX"
'    fw_compras_gral.Ado_detalle2.Recordset("doc_codigo").Value = "RE-402"
'    fw_compras_gral.Ado_detalle2.Recordset("doc_numero").Value = 0
    aw_compra_bienes.Ado_detalle4.Recordset("beneficiario_codigo").Value = dtc_codigo5.Text
    VAR_BENEF = aw_compra_bienes.Ado_detalle4.Recordset!beneficiario_codigo

'   aw_bienes_adjudica.Ado_detalle4.Recordset("adjudica_descripcion").Value = dtc_desc5.Text         'fw_compras_gral.Ado_datos.Recordset!compra_descripcion
'    fw_compras_gral.Ado_detalle2.Recordset("adjudica_cantidad_total").Value = fw_compras_gral.Ado_datos.Recordset!compra_cantidad_total

   
    aw_compra_bienes.Ado_detalle4.Recordset("apertura_fecha") = Date
    aw_compra_bienes.Ado_detalle4.Recordset("apertura_sobre_a") = "N"
    aw_compra_bienes.Ado_detalle4.Recordset("apertura_sobre_b") = "N"
    aw_compra_bienes.Ado_detalle4.Recordset("apertura_sobre_c") = "N"
    aw_compra_bienes.Ado_detalle4.Recordset("apertura_sobre_todos") = "N"
    aw_compra_bienes.Ado_detalle4.Recordset("proceso_codigo") = "TEC"
    aw_compra_bienes.Ado_detalle4.Recordset("subproceso_codigo") = "TEC-06"
    aw_compra_bienes.Ado_detalle4.Recordset("etapa_codigo") = "TEC-06-02"
    aw_compra_bienes.Ado_detalle4.Recordset("clasif_codigo") = "ADM"
    aw_compra_bienes.Ado_detalle4.Recordset("doc_codigo") = "R-114"
    aw_compra_bienes.Ado_detalle4.Recordset("doc_numero") = 0
    aw_compra_bienes.Ado_detalle4.Recordset("apertura_califica_sobre_a") = 0
    aw_compra_bienes.Ado_detalle4.Recordset("apertura_califica_sobre_b") = 0
    aw_compra_bienes.Ado_detalle4.Recordset("apertura_califica_sobre_c") = 0
    aw_compra_bienes.Ado_detalle4.Recordset("apertura_califica_total") = 0
    aw_compra_bienes.Ado_detalle4.Recordset("nro_nota_remision") = txt_Nota.Text
    aw_compra_bienes.Ado_detalle4.Recordset("apertura_fecha_estimada_adjudica") = Date
    aw_compra_bienes.Ado_detalle4.Recordset("apertura_tipo_recomendacion") = 0
    aw_compra_bienes.Ado_detalle4.Recordset("poa_codigo") = "3.2.8"
    aw_compra_bienes.Ado_detalle4.Recordset("usr_codigo") = glusuario
    aw_compra_bienes.Ado_detalle4.Recordset("fecha_registro") = Date
    aw_compra_bienes.Ado_detalle4.Recordset("hora_registro") = Format(Time, "HH:mm:ss")
    aw_compra_bienes.Ado_detalle4.Recordset("adjudica_cantidad_total") = txt_adjudica.Text
    aw_compra_bienes.Ado_detalle4.Recordset("adjudica_monto_bs").Value = txt_total_bs.Text
    aw_compra_bienes.Ado_detalle4.Recordset("adjudica_monto_dol").Value = txt_total_dol.Text
   
    aw_compra_bienes.Ado_detalle4.Recordset("fecha_inicio_contrato").Value = txtFecha.Value
    aw_compra_bienes.Ado_detalle4.Recordset("fecha_fin_contrato").Value = txtFecha2.Value
    aw_compra_bienes.Ado_detalle4.Recordset("fecha_recibe_almacen") = txtFecha3.Value
    aw_compra_bienes.Ado_detalle4.Recordset("concepto_proveedor") = dtc_desc5.Text
    aw_compra_bienes.Ado_detalle4.Recordset("mes_inicio_crono") = cmb_mes_ini
    aw_compra_bienes.Ado_detalle4.Recordset("cantidad_cuotas_pag") = txtCantCuota
    aw_compra_bienes.Ado_detalle4.Recordset("unimed_codigo_pag") = cmd_unimed2
    
 
    
'    sino = MsgBox("Desea APROBAR el Registro ? (Ya no podrá modificarlo)", vbYesNo + vbQuestion, "Atención")
'    If sino = vbYes Then
'        fw_compras_gral.Ado_detalle2.Recordset("estado_codigo") = "APR"
'        fw_compras_gral.Ado_detalle2.Recordset("usr_codigo_aprueba") = glusuario
'        fw_compras_gral.Ado_detalle2.Recordset("fecha_aprueba") = Date
'        db.Execute "update ao_compra_cabecera set estado_codigo_eqp = 'APR' WHERE compra_codigo = " & fw_compras_gral.Ado_detalle2.Recordset!compra_codigo & " "
'    Else
'        fw_compras_gral.Ado_detalle2.Recordset("estado_codigo") = "REG"
'    End If

    aw_compra_bienes.Ado_detalle4.Recordset.Update
'     db.Execute "UPDATE ao_compra_apertura_sobres SET adjudica_cantidad_total =  '" & Ado_datos13.Recordset!adjudica_cantidad_total & "' "
'   db.Execute "UPDATE ao_compra_apertura_sobres SET ao_compra_apertura_sobres.adjudica_cantidad_total = (SELECT SUM(adjudica_cantidad)AS TOTAL FROM ao_compra_adjudica_bienes  WHERE adjudica_codigo = " & lbl_adjudica.Caption & " ) WHERE ao_compra_apertura_sobres.adjudica_codigo = " & lbl_adjudica.Caption & " "
   Para_Aceptado = "S"
  Call CRONO_PAGO
'   frm_ao_solicitud_rrhh.Ado_detalle2.Refresh '.Recordset.Requery
   txtSW = "0"
   Unload Me
End If
End Sub

Private Sub CRONO_PAGO()
    Set rs_aux5 = New ADODB.Recordset
    If rs_aux5.State = 1 Then rs_aux5.Close
    rs_aux5.Open "select * from ao_compra_apertura_sobres where adjudica_codigo= " & VAR_ADJUDICA & " ", db, adOpenKeyset, adLockBatchOptimistic
    'Set AdoAux.Recordset = rsAuxDetalle
    If rs_aux5.RecordCount > 0 Then
      CONT2 = 1
      FInicio = rs_aux5!fecha_inicio_contrato
      VAR_MED2 = aw_compra_bienes.Ado_detalle4.Recordset!unimed_codigo_pag
      VAR_COBR2 = aw_compra_bienes.Ado_detalle4.Recordset!cantidad_cuotas_pag
      MControl = aw_compra_bienes.Ado_detalle4.Recordset!mes_inicio_crono
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
      FControl = FInicio
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
      
      While (CONT2 <= VAR_COBR2)
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from ao_compra_planilla_pagos where adjudica_codigo = " & VAR_ADJUDICA & "  ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 And corrprog >= VAR_COBR2 Then
            MsgBox "El Cronograma ya fue generado... ", , "Atención"
            CONT2 = CONT2 + 1
        Else
           'wwwwwwwwwwwwwwwwwwwwww
'          Set rs_aux1 = New ADODB.Recordset
'          If rs_aux1.State = 1 Then rs_aux1.Close
'          rs_aux1.Open "select * from ao_ventas_cabecera where ges_gestion='" & Ado_datos.Recordset!ges_gestion & "' and venta_codigo=" & Ado_datos.Recordset!venta_codigo & "  ", db, adOpenKeyset, adLockOptimistic
'          If rs_aux1.RecordCount > 0 Then
', , , , , , , , pendientes!!!!!
'                     pago_nro_cmpbte_factura, pago_nro_autorizacion,
'

            corrprog = rs_aux5!correl_pagos_prog + 1
            rs_aux5!correl_pagos_prog = rs_aux5!correl_pagos_prog + 1
            rs_aux5.Update
            
            rs_aux2.AddNew
            rs_aux2!ges_gestion = glGestion
'            rs_aux2!compra_codigo = VAR_COMPRA 'Ado_datos.Recordset("venta_codigo")
            rs_aux2!adjudica_codigo = aw_compra_bienes.Ado_detalle4.Recordset!adjudica_codigo
            rs_aux2!pago_codigo = corrprog
            rs_aux2!beneficiario_codigo = IIf(VAR_BENEF = "", dtc_codigo5.Text, VAR_BENEF)                 'Codigo Beneficiario/Cliente

            'OJO MODIFICAR COBRADOR - JQA 03-ENE-2015
'            rs_aux2!beneficiario_codigo_resp = "4333735"  'dtc_codigo4A.Text                                                     'Codigo Cobrador
            'rs_aux2!nombre_cobrador = dtc_desc4A.Text   '+ " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
'            Set rs_aux6 = New ADODB.Recordset
'            If rs_aux6.State = 1 Then rs_aux6.Close
'            rs_aux6.Open "select sum(venta_precio_unitario_bs) as acumBs from ao_ventas_detalle where venta_codigo = '" & var_cod5 & "' AND (par_codigo = '99990' or par_codigo = '43340') ", db, adOpenKeyset, adLockReadOnly
'            If rs_aux6.RecordCount > 0 Then
'                rs_aux2!cobranza_programada_bs = rs_aux6!acumBs                     'Monto Programado Bs
'                'db.Execute "INSERT INTO ao_almacen_detalle (almacen_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, stock_ingreso) VALUES (" & rs_aux6!almacen_codigo & ", '" & rs_aux6!bien_codigo & "', '" & rs_aux6!grupo_codigo & "', '" & rs_aux6!subgrupo_codigo & "', '" & rs_aux6!par_codigo & "' , " & rs_aux6!venta_det_cantidad & ")"
'                'acumDl = IIf(IsNull(rsrecalc!acumDl), 0, rsrecalc!acumDl)
'            Else
'                rs_aux2!cobranza_programada_bs = 0
'            End If
            rs_aux2!pago_monto_dol = Round(rs_aux5!adjudica_monto_dol / VAR_COBR2, 2) 'Monto Programado en Dolares
            rs_aux2!pago_monto_bs = Round(rs_aux5!adjudica_monto_bs / VAR_COBR2, 2)                                       'Monto Bs
            rs_aux2!pago_total_dol = Round(rs_aux5!adjudica_monto_dol / VAR_COBR2, 2)  'Monto Programado en Dolares
            rs_aux2!pago_total_bs = Round(rs_aux5!adjudica_monto_bs / VAR_COBR2, 2)                                       'Monto Bs
            'aquiiiiiiiiiiiiiiiiiiwwwwwwwwwwwwww
            rs_aux2!pago_descuento_bs = 0
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
                        If fanio = "2016" Or fanio = "2012" Or fanio = "2020" Or fanio = "2024" Then
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
                    rs_aux2!pago_fecha_prog = FControl
                    'rs_aux2!cobranza_fecha_conformidad = FControl + 10
                    rs_aux2!pago_fecha_efectiva = FControl + 20
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
            If rs_aux2!pago_fecha_prog = Null Then
                rs_aux2!pago_fecha_prog = Date
            End If
            'VAR_FEC2 = MonthName(Month(IIf(IsNull(rs_aux2!cobranza_fecha_prog), Date, rs_aux2!cobranza_fecha_prog)))
            
            VAR_FEC2 = MonthName(Month(IIf(IsNull(rs_aux2!pago_fecha_prog), FControl, rs_aux2!pago_fecha_prog)))
            'rs_aux2!cobranza_fecha_cobro = FControl + 10 ' rs_aux2!cobranza_fecha_prog + 10
            'If VAR_MED2 = "MES" Then
            '    FControl = FControl + Dias_Mes
            'End If
            'rs_aux2!cobranza_observaciones = "CUOTA Nro. " + Str(corrprog) + " - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(Date)) + " - " + lbl_titulo
            'rs_aux2!cobranza_observaciones = "CUOTA Nro. " + Str(corrprog) + " - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + lbl_titulo
            Select Case parametro
              Case "COMEX"
                  'rs_aux2!cobranza_observaciones = lbl_titulo + " - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + "Trámite: " + VAR_CITE + "-C-" + Str(corrprog)
                  rs_aux2!pago_descripcion = "PAGO " + lbl_titulo + " - " + Trim(dtc_desc5) + " - CUOTA Nro." + Str(corrprog)
              Case "DVTA"
                  'rs_aux2!cobranza_observaciones = "REPARACION DE EQUIPOS Y/O PROVISION E INSTALACION DE REPUESTOS - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + "Trámite: " + VAR_CITE + "-C-" + Str(corrprog)
                  rs_aux2!pago_descripcion = "PAGO " + lbl_titulo + " - " + Trim(dtc_desc5) + " - CUOTA Nro." + Str(corrprog)
              'Case "DVTA"
                  'rs_aux2!cobranza_observaciones = "PROVISION E INSTALACION DE EQUIPOS - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + "Trámite: " + VAR_CITE + "-C-" + Str(corrprog)
              Case Else
                  'rs_aux2!cobranza_observaciones = lbl_titulo + " - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + "Trámite: " + VAR_CITE + "-C-" + Str(corrprog)
                  rs_aux2!pago_descripcion = "PAGO " + lbl_titulo + " - " + Trim(dtc_desc5) + " - CUOTA Nro." + Str(corrprog)
            End Select
            CONT2 = CONT2 + 1
            rs_aux2!pago_emite_factura = "S"
            
            If rs_aux2!pago_total_dol <> 0 Then
                rs_aux2!Literal = Literal(CStr(rs_aux2!pago_total_dol)) + " DOLARES AMERICANOS"
            End If
            'rs_aux2!proceso_codigo = "CMX"
            'rs_aux2!subproceso_codigo = "CMX-01"
            'rs_aux2!etapa_codigo = "CMX-01-02"
            'rs_aux2!clasif_codigo = "TEC"
            'rs_aux2!doc_codigo = "R-105"    ' R-307 Certificado de Mantenimiento ' Colocar en la conformidad
            'rs_aux2!doc_numero = "0"        'var_cod5
            rs_aux2!poa_codigo = "4.1.1"
            ', poa_codigo,
            rs_aux2!estado_codigo = "REG"
            rs_aux2!usr_codigo = glusuario
            rs_aux2!Fecha_Registro = Format(Date, "dd/mm/yyyy")
            rs_aux2!hora_registro = Format(Time, "hh:mm:ss")
            rs_aux2.Update
            
        End If
           'wwwwwwwwwwwwwwwwwwwwww
            'GRABA ao_ventas_cobranza_prog
'           Set rs_aux4 = New ADODB.Recordset
'           If rs_aux4.State = 1 Then rs_aux4.Close
'           rs_aux4.Open "Select * from av_acumula_compras_detalle where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = '" & Ado_datos.Recordset!solicitud_codigo & "'   ", db, adOpenKeyset, adLockOptimistic
'           'rs_aux4.Open "Select * from ao_almacen_detalle where almacen_codigo = 0 and bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "'   ", db, adOpenKeyset, adLockOptimistic
'           If rs_aux4.RecordCount > 0 Then
'               db.Execute "INSERT INTO ao_almacen_detalle (almacen_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, stock_ingreso) SELECT " & rs_aux4!almacen_codigo & ", '" & rs_aux4!bien_codigo & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "' , '" & rs_aux4!bien_cantidad_adjudica & "' FROM av_acumula_compras_detalle WHERE almacen_codigo = '" & rs_almacen2!almacen_codigo & "'   And bien_codigo = '" & rs_almacen2!bien_codigo & "'    "
'           Else
'               If Ado_datos.Recordset!venta_tipo = "V" Then
'                   'db.Execute "INSERT INTO ao_almacen_detalle (almacen_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, stock_ingreso) VALUES (" & rsAuxDetalle!almacen_codigo & ", '" & rsAuxDetalle!bien_codigo & "', '" & rsAuxDetalle!grupo_codigo & "', '" & rsAuxDetalle!subgrupo_codigo & "', '" & rsAuxDetalle!par_codigo & "' , " & rsAuxDetalle!venta_det_cantidad & ")"
'                   db.Execute "INSERT INTO ao_almacen_detalle (almacen_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, stock_ingreso) VALUES (" & rsAuxDetalle!almacen_codigo & ", '" & rsAuxDetalle!bien_codigo & "', '" & rsAuxDetalle!grupo_codigo & "', '" & rsAuxDetalle!subgrupo_codigo & "', '" & rsAuxDetalle!par_codigo & "' , " & rsAuxDetalle!venta_det_cantidad & ")"
'               Else
'                   MsgBox "Error Verifique la Adjudicación de Bienes (Equipos, Repuestos u otros) ..."
'               End If
'           End If
'        End If
        'rsAuxDetalle.MoveNext
      Wend
      MsgBox "El Cronograma fue Generado Exitosamente... ", , "Atención"
      If corrprog > 0 Then
        db.Execute "update ao_compra_adjudica set correl_pagos_prog = '" & corrprog & "' "
        'db.Execute "update ao_compra_adjudica set venta_plazo_dias_calendario = " & CONT4 & " "
      End If
'      db.Execute "update ao_almacen_detalle set stock_actual = stock_ingreso - stock_salida"
    Else
       MsgBox "Error Verifique la Venta de Productos..."
    End If
End Sub

Private Sub GRABA_FICHA()
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "SELECT * FROM ro_rrhh_apertura_sobres where rrhh_codigo = " & aw_compra_bienes.Ado_datos13.Recordset!rrhh_codigo & "  ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        VAR_OCUP = rs_aux3!ocup_codigo
    Else
        VAR_OCUP = "0"
    End If
    
''    db.Execute "Insert INTO ro_personal_contratado_new (rrhh_codigo, beneficiario_codigo, puesto_codigo, estado_codigo, usr_codigo, fecha_registro) Values ('" & fw_compras_gral.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "',  'REG', '" & glusuario & "',  '" & Date & "')"
''    db.Execute "Insert INTO ro_personal_contratado (rrhh_codigo, beneficiario_codigo, puesto_codigo, estado_codigo, usr_codigo, fecha_registro) Values ('" & fw_compras_gral.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "',  'REG', '" & glusuario & "',  '" & Date & "')"
'
'    Set rs_aux2 = New ADODB.Recordset
'    If rs_aux2.State = 1 Then rs_aux2.Close
'    'rs_clasif1.Open "SELECT * FROM rc_puestos where puesto_vacante = 'SI' ORDER BY puesto_descripcion  ", DB, adOpenStatic
'    rs_aux2.Open "SELECT * FROM rc_puestos where puesto_codigo = '" & GlPuesto & "'  ", db, adOpenStatic
'    If rs_aux2.RecordCount > 0 Then
'        db.Execute "Insert INTO ro_personal_contratado (rrhh_codigo, beneficiario_codigo, puesto_codigo, unidad_codigo, cargo_codigo, fecha_ingreso, fecha_expiracion, ocup_codigo, beneficiario_haber_mensual, estado_codigo, usr_codigo, fecha_registro) Values ('" & fw_compras_gral.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "', '" & rs_aux2!unidad_codigo & "',  '" & rs_aux2!cargo_codigo & "',  '" & fw_compras_gral.Ado_detalle2.Recordset!beneficiario_fecha_inicio & "', '" & fw_compras_gral.Ado_detalle2.Recordset!beneficiario_fecha_fin & "', '" & VAR_OCUP & "', " & fw_compras_gral.Ado_detalle2.Recordset!beneficiario_monto_adjudica_dol & ", 'REG', '" & glusuario & "',  '" & Date & "')"
'        'db.Execute "Insert INTO ro_personal_contratado_NEW (rrhh_codigo, beneficiario_codigo, puesto_codigo, unidad_codigo, cargo_codigo, estado_codigo, usr_codigo, fecha_registro) Values ('" & fw_compras_gral.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "', '" & rs_aux2!unidad_codigo & "',  '" & rs_aux2!cargo_codigo & "',  'REG', '" & glusuario & "',  '" & Date & "')"
'    Else
'        db.Execute "Insert INTO ro_personal_contratado (rrhh_codigo, beneficiario_codigo, puesto_codigo, estado_codigo, usr_codigo, fecha_registro) Values ('" & fw_compras_gral.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "',  'REG', '" & glusuario & "',  '" & Date & "')"
'    End If
'    'Set Ado_clasif1.Recordset = rs_aux2

End Sub

Function Valida()
'valida que el monto asignado al beneficiario no sobrepase el monto pendiente de asignacion
    Valida = True
'  If (dtc_codigo1.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'  End If
  If (dtc_codigo5.Text = "") Then
    MsgBox "Debe registrar ... " + lblprov.Caption, vbCritical + vbExclamation, "Validación de datos"
    Valida = False
  End If
'  If (dtc_codigo3.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'  End If
'  If (dtc_codigo4.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'  End If
'  If txtPat = "" Then
'        Valida = False
'    End If
'    If txtNom = "" Then
'        Valida = False
'    End If
End Function



Private Sub cmb_tipomoneda_Click()
If cmb_tipomoneda = "USD" Then
txt_total_dol.Enabled = True
txt_total_bs.Enabled = False
Else
txt_total_dol.Enabled = False
txt_total_bs.Enabled = True
End If
End Sub

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux4.BoundText
    dtc_desc5.BoundText = dtc_aux4.BoundText
    dtc_aux5.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_aux5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux5.BoundText
    dtc_desc5.BoundText = dtc_aux5.BoundText
    dtc_aux4.BoundText = dtc_aux5.BoundText
End Sub
Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux4.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    dtc_aux4.BoundText = dtc_desc5.BoundText
    dtc_aux5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub Form_Activate()
    Set rs_clasif5 = New ADODB.Recordset
    If rs_clasif5.State = 1 Then rs_clasif5.Close
    Select Case Glaux
        Case "PROVI"    'PROVISION DE EQUIPOS
            rs_clasif5.Open "SELECT * FROM gc_beneficiario where pais_codigo= '" & txt_pais.Text & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
        Case "TRANS"    'TRANSPORTE
            rs_clasif5.Open "SELECT * FROM gc_beneficiario where tipoben_codigo = '3' or tipoben_codigo = '22' ORDER BY beneficiario_denominacion ", db, adOpenStatic
        Case "ADUAN"    'DESADUANIZACION
            rs_clasif5.Open "SELECT * FROM gc_beneficiario where tipoben_codigo = '3' or tipoben_codigo = '22' ORDER BY beneficiario_denominacion ", db, adOpenStatic
        Case "DESCA"    'DESCARGUIO Y OTROS
            rs_clasif5.Open "SELECT * FROM gc_beneficiario where tipoben_codigo = '3' or tipoben_codigo = '22' ORDER BY beneficiario_denominacion ", db, adOpenStatic
            Case "UALMI"    'PROVISION DE EQUIPOS
            rs_clasif5.Open "SELECT * FROM gc_beneficiario where tipoben_codigo <> '1'  ORDER BY beneficiario_denominacion ", db, adOpenStatic
    End Select
    Set Ado_clasif5.Recordset = rs_clasif5

End Sub

Private Sub Form_Load()
    'txtSW = "0"
'    Set rs_clasif1 = New ADODB.Recordset
'    If rs_clasif1.State = 1 Then rs_clasif1.Close
'    rs_clasif1.Open "SELECT * FROM ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & txt_codigo.Caption & " ORDER BY puesto_descripcion ", db, adOpenStatic
'    Set Ado_clasif1.Recordset = rs_clasif1
'
'    Set rs_clasif2 = New ADODB.Recordset
'    If rs_clasif2.State = 1 Then rs_clasif2.Close
'    rs_clasif2.Open "SELECT * FROM gc_ocupacion_profesion ORDER BY ocup_descripcion ", db, adOpenStatic
'    Set Ado_clasif2.Recordset = rs_clasif2
'
'    Set rs_clasif3 = New ADODB.Recordset
'    If rs_clasif3.State = 1 Then rs_clasif3.Close
'    rs_clasif3.Open "SELECT * FROM rc_nivel_educacional ORDER BY nivel_educ_descripcion ", db, adOpenStatic
'    Set Ado_clasif3.Recordset = rs_clasif3
'
'    Set rs_clasif4 = New ADODB.Recordset
'    If rs_clasif4.State = 1 Then rs_clasif4.Close
'    rs_clasif4.Open "SELECT * FROM gc_municipio where region_codigo = 'SI' ORDER BY munic_descripcion ", db, adOpenStatic
'    Set Ado_clasif4.Recordset = rs_clasif4

End Sub





Private Sub txt_total_bs_LostFocus()
If cmb_tipomoneda = "BOB" Then
    txt_total_dol.Text = Round(CDbl(IIf(txt_total_bs.Text = "", "0", txt_total_bs.Text)) / CDbl(txt_tipocambio), 2)
Else
    txt_total_bs.Text = Round(CDbl(IIf(txt_total_dol.Text = "", "0", txt_total_dol.Text)) / CDbl(txt_tipocambio), 2)
End If
End Sub

'Private Sub txt_total_dol_LostFocus()
'    If txt_total_dol.Text = "" Then
'        txt_total_dol.Text = "0"
'    End If
'    txt_total_bs.Text = CDbl(txt_total_dol) * GlTipoCambioOficial
'End Sub



Private Sub txt_total_dol_LostFocus()
If cmb_tipomoneda = "USD" Then
txt_total_bs.Text = Round(CDbl(IIf(txt_total_dol.Text = "", "0", txt_total_dol.Text)) / CDbl(txt_tipocambio), 2)
Else
txt_total_dol.Text = Round(CDbl(IIf(txt_total_bs.Text = "", "0", txt_total_bs.Text)) / CDbl(txt_tipocambio), 2)
End If
End Sub

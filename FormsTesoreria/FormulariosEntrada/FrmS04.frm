VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmS04 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesos Administrativos - Compras - Solicitudes de Compra"
   ClientHeight    =   9480
   ClientLeft      =   30
   ClientTop       =   2070
   ClientWidth     =   14835
   Icon            =   "FrmS04.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   37278.26
   ScaleMode       =   0  'User
   ScaleWidth      =   50790.41
   WindowState     =   2  'Maximized
   Begin VB.Frame Frmnavega 
      BackColor       =   &H00FFFFC0&
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
      Height          =   5130
      Left            =   0
      TabIndex        =   18
      Top             =   720
      Width           =   3300
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sin Enviar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   180
         Value           =   -1  'True
         Width           =   1215
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1980
         TabIndex        =   16
         Top             =   120
         Width           =   915
      End
      Begin MSAdodcLib.Adodc adosolicitud 
         Height          =   330
         Left            =   60
         Top             =   4770
         Width           =   3195
         _ExtentX        =   5636
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmS04.frx":0A02
         Height          =   4275
         Left            =   60
         TabIndex        =   110
         Top             =   480
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   7541
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
         Enabled         =   -1  'True
         ForeColor       =   0
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
         Caption         =   "CABECERA - PEDIDOS"
         ColumnCount     =   41
         BeginProperty Column00 
            DataField       =   "codigo_unidad"
            Caption         =   "UNIDAD"
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
            DataField       =   "codigo_solicitud"
            Caption         =   "Nro.Sol."
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
            DataField       =   "Ges_Gestion"
            Caption         =   "Ges_Gestion"
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
            DataField       =   "estado_aprobado"
            Caption         =   "APR"
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
            DataField       =   "estado_enviado"
            Caption         =   "ENV"
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
            DataField       =   "tipo_formulario"
            Caption         =   "TIPO"
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
            DataField       =   "justificacion_solicitud"
            Caption         =   "justificacion_solicitud"
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
            DataField       =   "CI"
            Caption         =   "CI"
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
            DataField       =   "Codigo_puesto"
            Caption         =   "Codigo_puesto"
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
            DataField       =   "CI_aprueba"
            Caption         =   "CI_aprueba"
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
            DataField       =   "Fecha_recepción"
            Caption         =   "Fecha_recepción"
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
            DataField       =   "fecha_solicitud"
            Caption         =   "fecha_solicitud"
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
            DataField       =   "codigo_poa"
            Caption         =   "codigo_poa"
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
            DataField       =   "tipo_moneda"
            Caption         =   "tipo_moneda"
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
            DataField       =   "monto_bolivianos"
            Caption         =   "monto_bolivianos"
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
            DataField       =   "monto_dolares"
            Caption         =   "monto_dolares"
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
            DataField       =   "Tipo_cambio"
            Caption         =   "Tipo_cambio"
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
         BeginProperty Column17 
            DataField       =   "monto_bolivianos_contra"
            Caption         =   "monto_bolivianos_contra"
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
         BeginProperty Column18 
            DataField       =   "monto_dolares_contra"
            Caption         =   "monto_dolares_contra"
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
         BeginProperty Column19 
            DataField       =   "org_codigo_contra"
            Caption         =   "org_codigo_contra"
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
         BeginProperty Column20 
            DataField       =   "Uni_codigo"
            Caption         =   "Uni_codigo"
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
         BeginProperty Column21 
            DataField       =   "consultor_empresa"
            Caption         =   "consultor_empresa"
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
         BeginProperty Column22 
            DataField       =   "nacional_extranjero"
            Caption         =   "nacional_extranjero"
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
         BeginProperty Column23 
            DataField       =   "funcion_actividad"
            Caption         =   "funcion_actividad"
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
         BeginProperty Column24 
            DataField       =   "duracion_estimada_numero"
            Caption         =   "duracion_estimada_numero"
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
         BeginProperty Column25 
            DataField       =   "duracion_estimada_tiempo"
            Caption         =   "duracion_estimada_tiempo"
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
         BeginProperty Column26 
            DataField       =   "impuestos"
            Caption         =   "impuestos"
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
         BeginProperty Column27 
            DataField       =   "por_tiempo"
            Caption         =   "por_tiempo"
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
         BeginProperty Column28 
            DataField       =   "fecha_estimada_inicio"
            Caption         =   "fecha_estimada_inicio"
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
         BeginProperty Column29 
            DataField       =   "tr_adjuntos"
            Caption         =   "tr_adjuntos"
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
         BeginProperty Column30 
            DataField       =   "observaciones"
            Caption         =   "observaciones"
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
         BeginProperty Column31 
            DataField       =   "codigo_bien"
            Caption         =   "codigo_bien"
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
         BeginProperty Column32 
            DataField       =   "caracteristicas"
            Caption         =   "caracteristicas"
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
         BeginProperty Column33 
            DataField       =   "usr_usuario"
            Caption         =   "usr_usuario"
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
         BeginProperty Column34 
            DataField       =   "pas_viat"
            Caption         =   "pas_viat"
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
         BeginProperty Column35 
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
         BeginProperty Column36 
            DataField       =   "hora_registro"
            Caption         =   "hora_registro"
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
         BeginProperty Column37 
            DataField       =   "usuario_aprueba"
            Caption         =   "usuario_aprueba"
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
         BeginProperty Column38 
            DataField       =   "fecha_aprueba"
            Caption         =   "fecha_aprueba"
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
         BeginProperty Column39 
            DataField       =   "hora_aprueba"
            Caption         =   "hora_aprueba"
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
         BeginProperty Column40 
            DataField       =   "Lista_adjunta"
            Caption         =   "Lista_adjunta"
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
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   434.835
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column21 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column22 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column23 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column24 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column25 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column26 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column27 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column28 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column29 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column30 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column31 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column32 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column33 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column34 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column35 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column36 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column37 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column38 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column39 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column40 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrmABMDet 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   2655
      Left            =   120
      TabIndex        =   106
      Top             =   6041
      Width           =   1335
      Begin VB.CommandButton BtnAnlDetalle 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Elimina Detalle"
         Height          =   705
         Left            =   60
         Picture         =   "FrmS04.frx":0A1D
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Anula Producto Elegido"
         Top             =   1800
         Width           =   1200
      End
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Modif. Detalle"
         Height          =   705
         Left            =   60
         Picture         =   "FrmS04.frx":0E5F
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Modifica Producto Elegido"
         Top             =   960
         Width           =   1200
      End
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Nuevo Detalle"
         Height          =   705
         Left            =   60
         Picture         =   "FrmS04.frx":12A1
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Adiciona Producto"
         Top             =   120
         Width           =   1200
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   3294
      TabIndex        =   26
      Top             =   720
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483632
      ForeColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CABECERA - PEDIDOS (Solicitud de Cotización)"
      TabPicture(0)   =   "FrmS04.frx":16E3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmgrabcabeza"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmabm"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "DETALLE (Productos del Pedido)"
      TabPicture(1)   =   "FrmS04.frx":16FF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmEditaDet"
      Tab(1).Control(1)=   "FrmGrabaDet"
      Tab(1).ControlCount=   2
      Begin VB.Frame frmabm 
         BackColor       =   &H00C0FFC0&
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
         Height          =   930
         Left            =   120
         TabIndex        =   114
         Top             =   360
         Width           =   11295
         Begin VB.CommandButton BtnAprobar 
            BackColor       =   &H8000000D&
            Caption         =   "Aprobar"
            Height          =   720
            Left            =   3840
            Picture         =   "FrmS04.frx":171B
            Style           =   1  'Graphical
            TabIndex        =   116
            ToolTipText     =   "Aprueba Registro"
            Top             =   120
            Width           =   770
         End
         Begin VB.CommandButton BtnAñadir 
            BackColor       =   &H8000000A&
            Caption         =   "Nuevo"
            Height          =   720
            Left            =   480
            Picture         =   "FrmS04.frx":1925
            Style           =   1  'Graphical
            TabIndex        =   124
            ToolTipText     =   "Nuevo Registro"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnModificar 
            BackColor       =   &H8000000A&
            Caption         =   "Modificar"
            Height          =   720
            Left            =   1320
            Picture         =   "FrmS04.frx":1F49
            Style           =   1  'Graphical
            TabIndex        =   123
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnEliminar 
            BackColor       =   &H8000000A&
            Caption         =   "Anular"
            Height          =   720
            Left            =   2160
            Picture         =   "FrmS04.frx":2529
            Style           =   1  'Graphical
            TabIndex        =   122
            ToolTipText     =   "Anula Registro Activo"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnSalir 
            BackColor       =   &H8000000A&
            Caption         =   "Cerrar"
            Height          =   720
            Left            =   9960
            Picture         =   "FrmS04.frx":31F3
            Style           =   1  'Graphical
            TabIndex        =   121
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnImprimir 
            BackColor       =   &H8000000A&
            Caption         =   "C/Precio"
            Height          =   720
            Left            =   7200
            Picture         =   "FrmS04.frx":33FD
            Style           =   1  'Graphical
            TabIndex        =   120
            ToolTipText     =   "Imprime Pedido con Precios"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnBuscar 
            BackColor       =   &H8000000A&
            Caption         =   "Buscar"
            Height          =   720
            Left            =   6360
            Picture         =   "FrmS04.frx":39BA
            Style           =   1  'Graphical
            TabIndex        =   119
            ToolTipText     =   "Busca un Registro"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnEnviar 
            BackColor       =   &H8000000D&
            Caption         =   "Enviar"
            Height          =   720
            Left            =   4680
            Picture         =   "FrmS04.frx":3F72
            Style           =   1  'Graphical
            TabIndex        =   118
            ToolTipText     =   "Envia a Pagos"
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton BtnDesAprobar 
            BackColor       =   &H8000000D&
            Caption         =   "Desapro."
            Height          =   720
            Left            =   3840
            Picture         =   "FrmS04.frx":4C3C
            Style           =   1  'Graphical
            TabIndex        =   117
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton BtnImprimirA 
            BackColor       =   &H8000000A&
            Caption         =   "S/Precio"
            Height          =   720
            Left            =   8040
            Picture         =   "FrmS04.frx":4E46
            Style           =   1  'Graphical
            TabIndex        =   115
            ToolTipText     =   "Imprime Pedido sin Precios"
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.Frame frmgrabcabeza 
         BackColor       =   &H00C0FFC0&
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
         Height          =   810
         Left            =   120
         TabIndex        =   111
         Top             =   360
         Visible         =   0   'False
         Width           =   11295
         Begin VB.CommandButton BtnCancelar 
            BackColor       =   &H8000000A&
            Caption         =   "Cancelar"
            Height          =   675
            Left            =   6360
            Picture         =   "FrmS04.frx":5403
            Style           =   1  'Graphical
            TabIndex        =   113
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnGrabar 
            BackColor       =   &H8000000A&
            Caption         =   "Grabar"
            Height          =   675
            Left            =   4320
            Picture         =   "FrmS04.frx":560D
            Style           =   1  'Graphical
            TabIndex        =   112
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.Frame FrmGrabaDet 
         BackColor       =   &H00FFFF80&
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
         Height          =   915
         Left            =   -74880
         TabIndex        =   94
         Top             =   360
         Visible         =   0   'False
         Width           =   11340
         Begin VB.CommandButton BtnGrabarDet 
            BackColor       =   &H80000000&
            Caption         =   "Grabar"
            Height          =   735
            Left            =   4080
            Picture         =   "FrmS04.frx":5817
            Style           =   1  'Graphical
            TabIndex        =   97
            ToolTipText     =   "Graba Datos del Producto"
            Top             =   120
            Width           =   825
         End
         Begin VB.CommandButton BtnCancelarDet 
            BackColor       =   &H80000000&
            Caption         =   "Cancelar"
            Height          =   735
            Left            =   6480
            Picture         =   "FrmS04.frx":5B21
            Style           =   1  'Graphical
            TabIndex        =   96
            ToolTipText     =   "Cancela Grabación"
            Top             =   120
            Width           =   825
         End
         Begin VB.CommandButton cmdElige 
            BackColor       =   &H80000000&
            Caption         =   "New Prod"
            Height          =   720
            Left            =   5280
            MaskColor       =   &H80000004&
            Picture         =   "FrmS04.frx":5E2B
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Registro Nuevo Producto"
            Top             =   120
            Visible         =   0   'False
            Width           =   825
         End
      End
      Begin VB.Frame FrmEditaDet 
         BackColor       =   &H80000010&
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
         Height          =   3765
         Left            =   -74900
         TabIndex        =   56
         Top             =   1320
         Width           =   11310
         Begin VB.Frame Frame5 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   10550
            TabIndex        =   91
            Top             =   2160
            Width           =   255
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   7350
            TabIndex        =   90
            Top             =   2160
            Width           =   255
         End
         Begin MSDataListLib.DataCombo Dtc_UniMed 
            Bindings        =   "FrmS04.frx":626D
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   9480
            TabIndex        =   75
            Top             =   3360
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "Unidad"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo DtcdesAnt 
            Bindings        =   "FrmS04.frx":6286
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   2100
            TabIndex        =   74
            Top             =   3360
            Width           =   7665
            _ExtentX        =   13520
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   0
            ListField       =   "Nombre_Anterior"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcPrecioUV 
            Bindings        =   "FrmS04.frx":629F
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   6120
            TabIndex        =   76
            Top             =   3000
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "Precio_estimado"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
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
         Begin VB.Frame Frame1 
            BackColor       =   &H8000000A&
            Caption         =   "REGISTRE LOS DATOS DEL PRODUCTO ..."
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
            Height          =   2025
            Left            =   20
            TabIndex        =   58
            Top             =   45
            Width           =   11265
            Begin MSDataListLib.DataCombo Dtcdesbien 
               Bindings        =   "FrmS04.frx":62B8
               DataField       =   "CodDetalle"
               DataSource      =   "adoao_solicitud_lista"
               Height          =   315
               Left            =   4080
               TabIndex        =   11
               Top             =   360
               Width           =   6975
               _ExtentX        =   12303
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "DescDetalle"
               BoundColumn     =   "CodDetalle"
               Text            =   ""
            End
            Begin VB.TextBox Txtrazon_s 
               CausesValidation=   0   'False
               DataField       =   "DescDetalle"
               DataSource      =   "adoao_solicitud_lista"
               Height          =   525
               Left            =   2040
               MaxLength       =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
               Top             =   1275
               Width           =   9015
            End
            Begin VB.TextBox TxtCantidad 
               Alignment       =   2  'Center
               DataField       =   "cantidad"
               DataSource      =   "adoao_solicitud_lista"
               Height          =   285
               Left            =   2040
               TabIndex        =   12
               Text            =   "1"
               Top             =   840
               Width           =   855
            End
            Begin VB.TextBox TxtPrecioU 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               DataField       =   "precio_compra"
               DataSource      =   "adoao_solicitud_lista"
               Height          =   285
               Left            =   5475
               TabIndex        =   59
               Text            =   "0"
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox TxtPrecioC 
               Alignment       =   2  'Center
               BackColor       =   &H8000000A&
               DataField       =   "precio_venta"
               DataSource      =   "adoao_solicitud_lista"
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
               Height          =   360
               Left            =   9840
               Locked          =   -1  'True
               TabIndex        =   13
               Text            =   "0"
               Top             =   800
               Width           =   1215
            End
            Begin MSDataListLib.DataCombo Dtccodbien 
               Bindings        =   "FrmS04.frx":62D1
               DataField       =   "CodDetalle"
               DataSource      =   "adoao_solicitud_lista"
               Height          =   315
               Left            =   2040
               TabIndex        =   10
               Top             =   360
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483638
               ForeColor       =   0
               ListField       =   "CodDetalle"
               BoundColumn     =   "CodDetalle"
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
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               Caption         =   "Elija el Producto:"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   240
               TabIndex        =   64
               Top             =   375
               Width           =   1185
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               Caption         =   "Características Comple- mentarias del Producto:"
               ForeColor       =   &H00C00000&
               Height          =   390
               Left            =   240
               TabIndex        =   63
               Top             =   1380
               Width           =   1860
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               Caption         =   "Cantidad a Solicitar:"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   240
               TabIndex        =   62
               Top             =   840
               Width           =   1410
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               Caption         =   "Precio Referencial Actual:"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   3555
               TabIndex        =   61
               Top             =   840
               Width           =   1845
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               Caption         =   "Precio Actual con % de Descuento:"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   7200
               TabIndex        =   60
               Top             =   840
               Width           =   2520
            End
         End
         Begin MSDataListLib.DataCombo DtcCodUniv 
            Bindings        =   "FrmS04.frx":62EA
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   4800
            TabIndex        =   57
            Top             =   3000
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   12632256
            ListField       =   "cod_univ"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo DtcSubgrupoDes 
            Bindings        =   "FrmS04.frx":6303
            DataField       =   "COD_MONTADOR"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   2520
            TabIndex        =   65
            Top             =   2565
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "descripcion"
            BoundColumn     =   "COD_MONTADOR"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcGrupo 
            Bindings        =   "FrmS04.frx":631D
            DataField       =   "CodGrupo"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   2160
            TabIndex        =   66
            Top             =   2160
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "DescGrupo"
            BoundColumn     =   "CodGrupo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcSubgrupo 
            Bindings        =   "FrmS04.frx":6334
            DataField       =   "COD_MONTADOR"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   1320
            TabIndex        =   67
            Top             =   2565
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "COD_MONTADOR"
            BoundColumn     =   "COD_MONTADOR"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcCodGrupo 
            Bindings        =   "FrmS04.frx":634E
            DataField       =   "CodGrupo"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   1320
            TabIndex        =   68
            Top             =   2160
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "CodGrupo"
            BoundColumn     =   "CodGrupo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcCodGrupoP 
            Bindings        =   "FrmS04.frx":6365
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   120
            TabIndex        =   69
            Top             =   2160
            Visible         =   0   'False
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483633
            ListField       =   "CodGrupo"
            BoundColumn     =   "CodDetalle"
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
         Begin MSDataListLib.DataCombo DtcSubgrupoP 
            Bindings        =   "FrmS04.frx":637E
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   120
            TabIndex        =   70
            Top             =   2520
            Visible         =   0   'False
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483633
            ListField       =   "COD_MONTADOR"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcPrecioC 
            Bindings        =   "FrmS04.frx":6397
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   9480
            TabIndex        =   71
            Top             =   2565
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "Precio_Compra"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo DtcPrecioU 
            Bindings        =   "FrmS04.frx":63B0
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   9480
            TabIndex        =   72
            Top             =   2160
            Visible         =   0   'False
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "Precio_Salon"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo DtcCodAnt 
            Bindings        =   "FrmS04.frx":63C9
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   480
            TabIndex        =   73
            Top             =   3360
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "Cod_Ant"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Unidad Medida"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   9465
            TabIndex        =   78
            Top             =   3120
            Width           =   1200
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "SubGrupo:"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   480
            TabIndex        =   83
            Top             =   2595
            Width           =   765
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Grupo :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   480
            TabIndex        =   82
            Top             =   2205
            Width           =   525
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Características del  Producto:"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   2280
            TabIndex        =   81
            Top             =   3120
            Width           =   2100
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Precio Referencial"
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   8040
            TabIndex        =   80
            Top             =   2205
            Visible         =   0   'False
            Width           =   1350
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Precio de Compra Base"
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   7680
            TabIndex        =   79
            Top             =   2595
            Width           =   1755
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Código Anterior"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   480
            TabIndex        =   77
            Top             =   3120
            Width           =   1200
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000010&
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
         Height          =   3870
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   11295
         Begin VB.Frame Frame2 
            BackColor       =   &H80000010&
            Caption         =   "--------------------------------------------- Datos Complementarios"
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
            Height          =   1995
            Left            =   180
            TabIndex        =   28
            Top             =   1800
            Width           =   10875
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "0.000%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   5280
               TabIndex        =   29
               Text            =   "0"
               Top             =   1530
               Visible         =   0   'False
               Width           =   420
            End
            Begin VB.CheckBox ChkTdr 
               BackColor       =   &H80000010&
               Caption         =   "Se Adjuntan Especificaciones, Folletos o Documentos ?"
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   240
               TabIndex        =   8
               Top             =   1560
               Value           =   1  'Checked
               Width           =   4335
            End
            Begin VB.TextBox txtterref 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4560
               MaxLength       =   1
               TabIndex        =   31
               Top             =   1500
               Visible         =   0   'False
               Width           =   420
            End
            Begin VB.TextBox Txt_porcentaje 
               Alignment       =   2  'Center
               DataField       =   "por_tiempo"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adosolicitud"
               Height          =   285
               Left            =   9600
               TabIndex        =   9
               Top             =   1560
               Width           =   780
            End
            Begin VB.TextBox txtjustifica 
               DataField       =   "justificacion_solicitud"
               DataSource      =   "adosolicitud"
               Height          =   285
               Left            =   2160
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Top             =   1180
               Width           =   8535
            End
            Begin VB.TextBox Txtcaracteristicas 
               DataField       =   "caracteristicas"
               DataSource      =   "adosolicitud"
               Height          =   500
               Left            =   2160
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Top             =   600
               Width           =   8535
            End
            Begin VB.TextBox Txtobservaciones 
               DataField       =   "observaciones"
               DataSource      =   "adosolicitud"
               Height          =   285
               Left            =   2160
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   30
               Top             =   1200
               Width           =   8535
            End
            Begin MSDataListLib.DataCombo DtcPOADes 
               Bindings        =   "FrmS04.frx":63E2
               DataField       =   "codigo_poa"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   3600
               TabIndex        =   5
               Top             =   240
               Width           =   7080
               _ExtentX        =   12488
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "descripcion_poa"
               BoundColumn     =   "codigo_poa"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtcdesbien1 
               Bindings        =   "FrmS04.frx":63F7
               DataField       =   "codGrupo"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   3600
               TabIndex        =   32
               Top             =   600
               Visible         =   0   'False
               Width           =   5880
               _ExtentX        =   10372
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "DescGrupo"
               BoundColumn     =   "CodGrupo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtccodbien1 
               Bindings        =   "FrmS04.frx":6412
               DataField       =   "codGrupo"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   2160
               TabIndex        =   33
               Top             =   600
               Visible         =   0   'False
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               BackColor       =   14737632
               ListField       =   "CodGrupo"
               BoundColumn     =   "codGrupo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DtcPOA 
               Bindings        =   "FrmS04.frx":642D
               DataField       =   "codigo_poa"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   2160
               TabIndex        =   34
               Top             =   240
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   -2147483632
               ForeColor       =   -2147483624
               ListField       =   "codigo_poa"
               BoundColumn     =   "codigo_poa"
               Text            =   ""
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
            Begin MSDataListLib.DataCombo DtcMarca 
               Bindings        =   "FrmS04.frx":6442
               DataField       =   "codigo_poa"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   9240
               TabIndex        =   84
               Top             =   120
               Visible         =   0   'False
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Appearance      =   0
               BackColor       =   12632256
               ForeColor       =   12648447
               ListField       =   "codigo_unidad"
               BoundColumn     =   "codigo_poa"
               Text            =   ""
            End
            Begin VB.Label Label2 
               BackColor       =   &H80000010&
               Caption         =   "Actividad del POA :"
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   240
               TabIndex        =   40
               Top             =   300
               Width           =   1875
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H80000010&
               Caption         =   "S=Si  /  N=No"
               ForeColor       =   &H00800000&
               Height          =   165
               Left            =   4080
               TabIndex        =   39
               Top             =   1590
               Visible         =   0   'False
               Width           =   1305
            End
            Begin VB.Label Label5 
               BackColor       =   &H80000010&
               Caption         =   "Porcentaje de Descuento asignado por el Proveedor:"
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   5775
               TabIndex        =   38
               Top             =   1575
               Width           =   3900
            End
            Begin VB.Label Label16 
               BackColor       =   &H80000010&
               Caption         =   "Observaciones :"
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   240
               TabIndex        =   37
               Top             =   1180
               Width           =   1275
            End
            Begin VB.Label Label26 
               BackColor       =   &H80000010&
               Caption         =   "Características Generales:"
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   240
               TabIndex        =   36
               Top             =   720
               Width           =   1980
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackColor       =   &H80000010&
               Caption         =   " % "
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
               Height          =   165
               Left            =   10440
               TabIndex        =   35
               Top             =   1560
               Width           =   225
            End
         End
         Begin VB.Frame Frame2A 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000040&
            Height          =   480
            Left            =   -60
            TabIndex        =   41
            Top             =   1320
            Width           =   11115
            Begin VB.OptionButton Option2 
               BackColor       =   &H80000010&
               Caption         =   "Persona"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   210
               Left            =   1440
               TabIndex        =   93
               Top             =   220
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H80000010&
               Caption         =   "Empresa"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   210
               Left            =   240
               TabIndex        =   92
               Top             =   220
               Visible         =   0   'False
               Width           =   1095
            End
            Begin MSDataListLib.DataCombo Dtcpaternobe 
               Bindings        =   "FrmS04.frx":6457
               DataField       =   "ci_aprueba"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   4200
               TabIndex        =   4
               Top             =   55
               Width           =   6915
               _ExtentX        =   12197
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "denominacion_beneficiario"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtccibe 
               Bindings        =   "FrmS04.frx":6471
               DataField       =   "ci_aprueba"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   2625
               TabIndex        =   42
               Top             =   60
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   -2147483632
               ListField       =   "codigo_beneficiario"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00000040&
               X1              =   240
               X2              =   2400
               Y1              =   300
               Y2              =   300
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackColor       =   &H80000010&
               Caption         =   "Proveedor"
               ForeColor       =   &H00C00000&
               Height          =   165
               Left            =   240
               TabIndex        =   87
               Top             =   0
               Width           =   2175
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Frasolic 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000040&
            Height          =   465
            Left            =   -60
            TabIndex        =   49
            Top             =   780
            Width           =   11115
            Begin MSDataListLib.DataCombo Dtcpaternosol 
               Bindings        =   "FrmS04.frx":648B
               DataField       =   "ci"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   4200
               TabIndex        =   3
               Top             =   130
               Width           =   6915
               _ExtentX        =   12197
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "denominacion_beneficiario"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtccisol 
               Bindings        =   "FrmS04.frx":64A6
               DataField       =   "ci"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   2625
               TabIndex        =   50
               Top             =   135
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   -2147483632
               ListField       =   "codigo_beneficiario"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H80000010&
               Caption         =   "Responsable de la Solicitud :"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   240
               TabIndex        =   86
               Top             =   180
               Width           =   2415
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame FrmApertura 
            Caption         =   "APERTURA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   705
            Left            =   360
            TabIndex        =   43
            Top             =   2400
            Visible         =   0   'False
            Width           =   8055
            Begin VB.ComboBox cmbSubCta2 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "FrmS04.frx":64C1
               Left            =   6240
               List            =   "FrmS04.frx":64CB
               TabIndex        =   44
               Top             =   255
               Visible         =   0   'False
               Width           =   1695
            End
            Begin MSDataListLib.DataCombo DtCvalor1 
               Bindings        =   "FrmS04.frx":64E1
               Height          =   315
               Left            =   1680
               TabIndex        =   45
               Top             =   255
               Visible         =   0   'False
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "valor1"
               BoundColumn     =   "TIPO"
               Text            =   ""
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Trámite:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   48
               Top             =   285
               Width           =   1395
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Cargo de Cuenta:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   4200
               TabIndex        =   47
               Top             =   285
               Width           =   2055
            End
            Begin VB.Label Lbltipo_bien_Cta_doc1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   225
               Left            =   6105
               TabIndex        =   46
               Top             =   345
               Width           =   1875
            End
         End
         Begin MSComCtl2.DTPicker DTPfechasol 
            DataField       =   "fecha_solicitud"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   9360
            TabIndex        =   2
            Top             =   480
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            Format          =   21233665
            CurrentDate     =   40909
            MaxDate         =   73415
            MinDate         =   36526
         End
         Begin MSDataListLib.DataCombo DtcUnidad 
            Bindings        =   "FrmS04.frx":6500
            DataField       =   "codigo_unidad"
            DataSource      =   "adosolicitud"
            Height          =   315
            Left            =   1800
            TabIndex        =   0
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "codigo_unidad"
            BoundColumn     =   "codigo_unidad"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo DtcUnidadDes 
            Bindings        =   "FrmS04.frx":6518
            DataField       =   "codigo_unidad"
            DataSource      =   "adosolicitud"
            Height          =   315
            Left            =   4440
            TabIndex        =   51
            Top             =   480
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Uni_descripcion_larga"
            BoundColumn     =   "codigo_unidad"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcTipoS 
            Bindings        =   "FrmS04.frx":6530
            DataField       =   "TipoF1"
            DataSource      =   "adosolicitud"
            Height          =   315
            Left            =   5040
            TabIndex        =   1
            Top             =   480
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "descripcion_t_solicitud"
            BoundColumn     =   "Codigo_t_solicitud"
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
         Begin MSDataListLib.DataCombo DtcTipo 
            Bindings        =   "FrmS04.frx":6549
            DataField       =   "TipoF1"
            DataSource      =   "adosolicitud"
            Height          =   315
            Left            =   7440
            TabIndex        =   89
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Codigo_t_solicitud"
            BoundColumn     =   "Codigo_t_solicitud"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "FrmS04.frx":6562
            DataField       =   "codigo_unidad"
            DataSource      =   "adosolicitud"
            Height          =   315
            Left            =   3360
            TabIndex        =   105
            Top             =   120
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "codigo_poa"
            BoundColumn     =   "codigo_unidad"
            Text            =   ""
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Tipo de Solicitud:"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   5160
            TabIndex        =   88
            Top             =   240
            Width           =   1230
         End
         Begin VB.Label txtnrosol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            Caption         =   "Label29"
            DataField       =   "codigo_solicitud"
            DataSource      =   "adosolicitud"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Lbltipo_bien_Cta_doc 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   2880
            TabIndex        =   55
            Top             =   1065
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Unidad Productiva:"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1920
            TabIndex        =   54
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Fecha de Solicitud:"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   9480
            TabIndex        =   53
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000010&
            Caption         =   "No.Solicitud:"
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
            Height          =   285
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   1125
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   14835
      TabIndex        =   19
      Top             =   0
      Width           =   14895
      Begin VB.Label LblUni_descripcion_larga 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   3480
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   5160
      End
      Begin VB.Label label7 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1200
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label lblUni_codigo 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1200
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOLICITUDES DE COMPRA (PEDIDOS)"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   405
         Left            =   8565
         TabIndex        =   20
         Top             =   120
         Width           =   5910
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmS04.frx":657A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15360
      End
   End
   Begin MSAdodcLib.Adodc Adocc_parametros 
      Height          =   330
      Left            =   0
      Top             =   9120
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
      Caption         =   "Adocc_parametros"
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
   Begin MSAdodcLib.Adodc adopuestosol 
      Height          =   330
      Left            =   2160
      Top             =   9120
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
      Caption         =   "adopuestosol"
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
   Begin MSAdodcLib.Adodc adopuestobe 
      Height          =   330
      Left            =   2160
      Top             =   8760
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
      Caption         =   "adopuestobe"
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
   Begin MSAdodcLib.Adodc AdoUnidad 
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
      Caption         =   "AdoUnidad"
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
   Begin MSAdodcLib.Adodc AdoPOA 
      Height          =   330
      Left            =   6360
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
      Caption         =   "AdoPOA"
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
   Begin MSAdodcLib.Adodc adoao_solicitud_detalle 
      Height          =   330
      Left            =   8400
      Top             =   8760
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "Navegar"
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
   Begin MSAdodcLib.Adodc ado_bienes 
      Height          =   330
      Left            =   0
      Top             =   8760
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
      Caption         =   "ado_bienes"
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
   Begin VB.Frame Frame10 
      BackColor       =   &H00400000&
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
      ForeColor       =   &H00808000&
      Height          =   2880
      Left            =   0
      TabIndex        =   24
      Top             =   5880
      Width           =   14844
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "FrmS04.frx":4C1BC
         Height          =   2295
         Left            =   1560
         TabIndex        =   17
         Top             =   180
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777088
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
         Caption         =   "DETALLE (Productos del Pedido)"
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "CodGrupo"
            Caption         =   "Grupo"
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
            DataField       =   "cod_montador"
            Caption         =   "Sub-Grupo"
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
            DataField       =   "CodDetalle"
            Caption         =   "Codigo Producto"
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
            DataField       =   "DescDetalle"
            Caption         =   "Denominación del Producto"
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
            DataField       =   "cantidad"
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
         BeginProperty Column05 
            DataField       =   "precio_compra"
            Caption         =   "Precio.Actual"
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
            DataField       =   "Total_compra"
            Caption         =   "Total Actual"
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
            DataField       =   "precio_venta"
            Caption         =   "Precio.c/Dscto."
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
            DataField       =   "Total_venta"
            Caption         =   "Total c/Dscto."
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
            DataField       =   "profesion"
            Caption         =   "Caracteristicas del Bien"
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
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   4619.906
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DtGao_solicitud_detalle 
         Bindings        =   "FrmS04.frx":4C1E0
         Height          =   1275
         Left            =   1800
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2249
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648447
         Enabled         =   -1  'True
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "codigo_solicitud"
            Caption         =   "Solicitud"
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
            DataField       =   "codigo_detalle"
            Caption         =   "Nro.Det."
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
            DataField       =   "codigo_unidad"
            Caption         =   "Unidad"
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
            DataField       =   "codigo_poa"
            Caption         =   "Frente Servic."
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
            DataField       =   "monto_bolivianos"
            Caption         =   "Monto_Bs. (B)"
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
            DataField       =   "monto_dolares"
            Caption         =   "Monto_$US (B)"
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
            DataField       =   "Tipo_cambio"
            Caption         =   "TDC"
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
            DataField       =   "monto_bolivianos_contra"
            Caption         =   "Monto_Bs. (I)"
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
            DataField       =   "monto_dolares_contra"
            Caption         =   "Monto_$US (I)"
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
            DataField       =   "tipo_moneda"
            Caption         =   "Moneda"
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
            DataField       =   "org_codigo_ext"
            Caption         =   "Fin_principal"
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
            DataField       =   "org_codigo_contra"
            Caption         =   "Fin_Impuesto"
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
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoao_solicitud_lista 
         Height          =   330
         Left            =   1560
         Top             =   2520
         Visible         =   0   'False
         Width           =   6435
         _ExtentX        =   11351
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
         BackColor       =   16777088
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
         Caption         =   " <-- Inicio                      Detalle de Productos                              Fin -->"
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
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTALES -->  "
         Height          =   330
         Left            =   8040
         TabIndex        =   104
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label LblTotCant 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   330
         Left            =   9360
         TabIndex        =   103
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label LblTotPrecAc 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   330
         Left            =   10080
         TabIndex        =   102
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label LblTotAct 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   330
         Left            =   11160
         TabIndex        =   101
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label LblTotPrecDsto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   330
         Left            =   12360
         TabIndex        =   100
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label LblTotDscto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   330
         Left            =   13440
         TabIndex        =   99
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lbltipoVenta 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "DETALLE DE DATOS DE LOS PRODUCTOS POR CADA PEDIDO ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1155
         Left            =   120
         TabIndex        =   98
         Top             =   720
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   5
         X1              =   0
         X2              =   14760
         Y1              =   60
         Y2              =   60
      End
   End
   Begin MSAdodcLib.Adodc AdoGrupo 
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
      Caption         =   "AdoGrupo"
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
   Begin MSAdodcLib.Adodc AdoMontador 
      Height          =   330
      Left            =   6360
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
      Caption         =   "AdoMontador"
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
   Begin MSAdodcLib.Adodc AdoTipoSol 
      Height          =   330
      Left            =   8400
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
      Caption         =   "AdoTipoSol"
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
   Begin Crystal.CrystalReport CryF01 
      Left            =   0
      Top             =   9480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "FrmS04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rscc_parametros As New ADODB.Recordset
Dim rstdestino As New ADODB.Recordset
Dim conv1, conv2, conv_nal, CONVE As String
Dim cta1, txtGes_gestion As String
Dim cat_nal, CATEG, SOLISTA As String
Dim parametro, parametro2 As String
Dim GCODIGO_PAGO As String
'
Dim rstdetsalalm As New ADODB.Recordset
Dim rstAo_solicitud As New ADODB.Recordset
Dim rstao_solicitud_detalle As New ADODB.Recordset
Dim rstpoa As New ADODB.Recordset
Dim rstpoaAux As New ADODB.Recordset
Dim rstrc_personalSoli As New ADODB.Recordset
Dim rstrc_personalCargo As New ADODB.Recordset
Dim rstfc_partida_gasto As New ADODB.Recordset
Dim rstFc_unidad_ejecutora As New ADODB.Recordset
Dim rstac_bienes As New ADODB.Recordset
Dim rstfc_relacionador_poa_ppto As New ADODB.Recordset
Dim rstOrganismo_finanExt As New ADODB.Recordset
Dim rstao_solicitud_lista As New ADODB.Recordset
Dim rs_ResponsableAaux As New ADODB.Recordset
Dim rs_soldetaux As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim rs_Bienes As New ADODB.Recordset
Dim rs_montador As New ADODB.Recordset
Dim rsgrupo, rsgrupo2 As New ADODB.Recordset
Dim rstcodigo_detalle As New ADODB.Recordset
Dim rsdetalle As New ADODB.Recordset
Dim rs_TipoSol As New ADODB.Recordset
Dim rs_proveedor As New ADODB.Recordset

Dim swgrabar, valida As Integer
Dim correlsolic As Integer
Dim correldetalle As Integer
Dim swunidad, tot_form, prev_dev As Integer
Dim marca1 As BookmarkEnum
Dim ext1, tgn1 As Double
Dim precuni, precTot, precSln, precTotV As Double
Dim cantTot As Integer
'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir As String
Dim queryinicial As String
Dim queryinicial2 As String
Dim sino As Integer
'MODI A
Dim V_accion As String
'Para PCE, Pagos_espera y Pagos
  'Dim rstdestino As New ADODB.Recordset
  Dim rstorigen As New ADODB.Recordset
  Dim rstpagos As New ADODB.Recordset
  Dim rstpago_detalle As New ADODB.Recordset
  Dim rscorrelativo As New ADODB.Recordset
  
  Dim Proyecto1 As String
  Dim Par_Codigo1 As String
  Dim Organismo1 As String
  Dim fte_codigo1 As String
  Dim Org_Codigo1 As String
  Dim pro_Programa1 As String
'  Dim Pro_SubPrograma1 As String
  Dim Pro_Proyecto1 As String
  Dim Pro_Actividad1 As String
  Dim gestion1 As String
  Dim uni_codigo1 As String
  Dim COD_SOL As Integer
  Dim codigo_categoria1 As String
  Dim codigo_convenio1 As String
  Dim Fte_contraparte1 As String
  Dim Org_Contraparte1 As String
  
  Dim por_fte_ext1 As Double
  Dim por_fte_nal1 As Double
  Dim codigo_pago1 As Double
  Dim ges_gestion1 As String
  Dim ConvExt, ConvNAl As String
  Dim CatExt, CatNal As String
  
  Dim swpresup As Integer
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim v_por_fte(3, 3)
  Dim tot_reg As Integer
  Dim rectot As Integer
  Dim CODPAG As Integer

  Dim rssolista As New ADODB.Recordset
  Dim rstao_solicitud_recibido As New ADODB.Recordset
  Dim swSubir As String
  Dim swnuevo As Integer

Private Sub adosalalm_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If pRecordset.EOF Or pRecordset.BOF Then Exit Sub
        Select Case pRecordset.EditMode
        Case adEditNone
            If rstdetsalalm.State = 1 Then rstdetsalalm.Close
            rstdetsalalm.Open "Select * from ao_detallesalidaalmacen where correlativo_salida = '" & pRecordset("correlativo_salida") & "'", db, adOpenDynamic, adLockOptimistic
'            Set DataGrid2.DataSource = Nothing
'            Set DataGrid2.DataSource = rstdetsalalm
'            DataGrid2.ReBind
        End Select
End Sub


Private Sub adoao_solicitud_lista_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    'Set adoao_solicitud_lista.Recordset = rstao_solicitud_lista
    'Set DtGLista.DataSource = rstao_solicitud_lista
    Set DtGLista.DataSource = AdoAo_solicitud_Lista.Recordset
End Sub

Private Sub adosolicitud_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'JQA JUN/2012
  If swgrabar = 0 Then
   If (Not adosolicitud.Recordset.BOF) And (Not adosolicitud.Recordset.EOF) Then
      If Not IsNull(adosolicitud.Recordset("codigo_solicitud")) And (adosolicitud.Recordset("ci") <> " ") Then
         If adosolicitud.Recordset("tr_adjuntos") = "S" Then ChkTdr.Value = 1
         If adosolicitud.Recordset("tr_adjuntos") = "N" Then ChkTdr.Value = 0
         If adosolicitud.Recordset("tr_adjuntos") = "E" Then ChkTdr.Value = 2
         DTPfechasol.Value = adosolicitud.Recordset("fecha_solicitud")
         'If Not (IsNull(adosolicitud.Recordset("ci"))) Then
            lblUni_codigo = IIf(IsNull(adosolicitud.Recordset("codigo_unidad")) = True, "ADMIN", adosolicitud.Recordset("codigo_unidad"))
            Set rstao_solicitud_detalle = New ADODB.Recordset
            If rstao_solicitud_detalle.State = 1 Then rstao_solicitud_detalle.Close
            queryinicial2 = "select * from ao_solicitud_detalle where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud
            rstao_solicitud_detalle.Open queryinicial2, db, adOpenKeyset, adLockReadOnly
            If rstao_solicitud_detalle.RecordCount > 0 Then
                SOLISTA = "A"
                Frame10.Visible = True
                DtGLista.Visible = True
                AdoAo_solicitud_Lista.Visible = True
                parametro = " ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
                Call ABRE_SOL_LISTA
            Else
                SOLISTA = "B"
                Frame10.Visible = False
                DtGLista.Visible = False
                AdoAo_solicitud_Lista.Visible = False
            End If
            Set adoao_solicitud_detalle.Recordset = rstao_solicitud_detalle
            adoao_solicitud_detalle.Refresh
         'End If
         'jqa DIC-2012
         'DtGLista.Caption = "DETALLE PRODUCTOS - SOLICITUD NRO. " + Str((adosolicitud.Recordset("CODIGO_SOLICITUD")))
            'jqa
         If adosolicitud.Recordset!estado_enviado = "S" And adosolicitud.Recordset!ESTADO_APROBADO = "S" Then
             BtnEnviar.Visible = False
             BtnEliminar.Visible = False
             BtnAprobar.Visible = False
             BtnDesAprobar.Visible = False
             BtnModificar.Visible = False
             FrmABMDet.Visible = False
         End If
         If adosolicitud.Recordset!estado_enviado = "E" And adosolicitud.Recordset!ESTADO_APROBADO = "E" Then
             BtnEnviar.Visible = False
             BtnEliminar.Visible = False
             BtnAprobar.Visible = False
             BtnDesAprobar.Visible = False
             BtnModificar.Visible = False
             FrmABMDet.Visible = False
         End If
         If adosolicitud.Recordset!ESTADO_APROBADO = "E" And adosolicitud.Recordset!estado_enviado = "N" Then
             BtnEnviar.Visible = False
             BtnEliminar.Visible = False
             BtnAprobar.Visible = False
             BtnDesAprobar.Visible = False
             BtnModificar.Visible = False
             FrmABMDet.Visible = False
         End If
         If adosolicitud.Recordset!ESTADO_APROBADO = "S" And adosolicitud.Recordset!estado_enviado = "N" Then
             BtnEnviar.Visible = True
             BtnEliminar.Visible = True
             BtnAprobar.Visible = False
             BtnDesAprobar.Visible = True
             BtnModificar.Visible = False
             FrmABMDet.Visible = False
         End If
         If adosolicitud.Recordset!ESTADO_APROBADO = "N" And adosolicitud.Recordset!estado_enviado = "N" Then
             BtnEnviar.Visible = False
             BtnEliminar.Visible = False
             BtnAprobar.Visible = True
             BtnDesAprobar.Visible = False
             BtnModificar.Visible = True
             FrmABMDet.Visible = True
         End If
         
      Else
            ' por si es nuevo
      End If
      If IsNull(adosolicitud.Recordset!por_tiempo) Then
        Text1.Text = "100"
      Else
        Text1.Text = CDbl(adosolicitud.Recordset!por_tiempo) * 100
      End If
'      Text1.Text = IIf(IsNull(adosolicitud.Recordset!por_tiempo), 100, CDbl(adosolicitud.Recordset!por_tiempo) * 100)
   Else
        BtnModificar.Visible = False
        BtnEliminar.Visible = False
        BtnDesAprobar.Visible = False
        BtnAprobar.Visible = False
        BtnEnviar.Visible = False
        FrmABMDet.Visible = False
   End If
  End If
'JQA JUN/2012
End Sub

Private Sub BtnCancelarDet_Click()
  Call cerea
  swnuevo = 0
  Frmnavega.Enabled = True
  Frmnavega.Visible = True
  Frame10.Enabled = True
  FrmEditaDet.Enabled = False
  Call OptFilGral1_Click
  rstao_solicitud_lista.Requery
  AdoAo_solicitud_Lista.Refresh
  sstab1.Tab = 0
  sstab1.TabEnabled(1) = True
  sstab1.TabEnabled(0) = True
  SOLISTA = "B"
  FrmGrabaDet.Visible = False
  FrmABMDet.Visible = True
End Sub


Private Sub cmdElige_Click()
    AlFrmCreaMaterial.Show
End Sub

Private Sub BtnImprimir_Click()
    'JQA JUN/2008
If adosolicitud.Recordset!Lista_adjunta = "S" Then
    Dim co As New ADODB.Command
    CryF01.ReportFileName = App.Path & "\Reportes\Solicitudes\C01_F04_SP.rpt"
    CryF01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
      CryF01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
      CryF01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
    'Call CREAVISTAF11          'JQA JUN-2008
    CryF01.StoredProcParam(0) = Me.adosolicitud.Recordset!ges_gestion
    CryF01.StoredProcParam(1) = Me.adosolicitud.Recordset!codigo_unidad
    CryF01.StoredProcParam(2) = Me.adosolicitud.Recordset!codigo_solicitud
    iResult = CryF01.PrintReport
    If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar el detalle del registro ...", , "Atención"
End If

End Sub

Private Sub BtnEliminarDet_Click()
  If adosolicitud.Recordset("estado_aprobado") = "N" And adosolicitud.Recordset("estado_enviado") = "N" Then
    sino = MsgBox("Está seguro de eliminar este registro", vbYesNo + vbQuestion, "Atención ...")
    If sino = vbYes Then
      AdoAo_solicitud_Lista.Recordset.Delete
      AdoAo_solicitud_Lista.Recordset.Update
      Call ABRE_SOL_LISTA
      'rstao_solicitud_lista.Requery
      'adoao_solicitud_lista.Refresh
    End If
  Else
    MsgBox "No se puede ANULAR un registro Aprobado ó Enviado !! ", vbExclamation
  End If
End Sub


Private Sub DtcTipo_Click(Area As Integer)
    DtcTipoS.BoundText = DtcTipo.BoundText
End Sub

Private Sub DtcTipoS_Click(Area As Integer)
    DtcTipo.BoundText = DtcTipoS.BoundText
End Sub

Private Sub Option1_Click()
    Set rs_proveedor = New ADODB.Recordset
    If rs_proveedor.State = 1 Then rs_proveedor.Close
    rs_proveedor.Open "select * from fc_beneficiario where tipo_beneficiario=22 AND procedencia='" & DtCUnidad & "'ORDER BY denominacion_beneficiario", db, adOpenKeyset, adLockReadOnly
    If rs_proveedor.RecordCount = 0 Then
        If rs_proveedor.State = 1 Then rs_proveedor.Close
        rs_proveedor.Open "select * from fc_beneficiario where tipo_beneficiario=22 ORDER BY denominacion_beneficiario", db, adOpenKeyset, adLockReadOnly
    End If
    Set adopuestobe.Recordset = rs_proveedor
    adopuestobe.Refresh
    Dtcpaternobe.Enabled = True
End Sub

Private Sub Option2_Click()
    Set rs_proveedor = New ADODB.Recordset
    If rs_proveedor.State = 1 Then rs_proveedor.Close
    rs_proveedor.Open "select * from fc_beneficiario where tipo_beneficiario=2 ORDER BY denominacion_beneficiario", db, adOpenKeyset, adLockReadOnly
    Set adopuestobe.Recordset = rs_proveedor
    adopuestobe.Refresh
    Dtcpaternobe.Enabled = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If sstab1.Tab = 0 Then
        'SSTab1.TabEnabled(0) = True
        'SSTab1.TabEnabled(1) = False
    Else
      'If adoao_solicitud_lista.Recordset!codigo_solicitud = 0 Then
      If swnuevo = 1 Then
        'MsgBox "ERR"
        FrmEditaDet.Visible = True
        'DtGLista.Visible = True
        Frame10.Enabled = True
        AdoAo_solicitud_Lista.Visible = True
      Else
        If adosolicitud.Recordset!Lista_adjunta = "S" Then
        'If SOLISTA = "A" Then
           FrmEditaDet.Visible = True
           'DtGLista.Visible = True
           Frame10.Enabled = True
           AdoAo_solicitud_Lista.Visible = True
         Else
           FrmEditaDet.Visible = False
           'DtGLista.Visible = False
           Frame10.Enabled = False
           AdoAo_solicitud_Lista.Visible = False
         End If
      End If
    End If
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub valida2()
   Set rssolista = New ADODB.Recordset
   If rssolista.State = 1 Then rssolista.Close
   rssolista.Open "select * from ao_solicitud_lista where codigo_solicitud = " & COD_SOL & " and codigo_unidad = '" & uni_codigo1 & "' and CodDetalle = '" & Dtccodbien & "' ", db, adOpenKeyset, adLockOptimistic
'   Set AdoUnidad.Recordset = rssolista
'   AdoUnidad.Refresh
    If rssolista.RecordCount > 0 Then
        valida = 2
        MsgBox "El producto ya fue registrado. Intente nuevamente !! ", vbExclamation + vbOKOnly, "Validación de Datos"
        Exit Sub
    End If
End Sub

Private Sub BtnGrabarDet_Click()
 valida = 1
 If swnuevo = 1 Then
    Call valida2
 End If
 If valida = 1 Then
  If Not IsNumeric(TxtCantidad.Text) Then
     MsgBox "El dato registrado en <Cantidad a Solicitar:> debe ser un Valor Numérico Válido.", vbExclamation + vbOKOnly, "Validación de Datos"
     Exit Sub
  End If
  If Not IsNumeric(TxtPrecioU.Text) Then
     MsgBox "El dato registrado en <Precio Referncial Actual:> debe ser un Valor Numérico Válido.", vbExclamation + vbOKOnly, "Validación de Datos"
     Exit Sub
  End If
  If Dtccodbien <> "" And Val(TxtCantidad.Text) >= 0 And CDbl(TxtPrecioU.Text) >= 0 Then
    db.BeginTrans
    If swnuevo = 1 Then
      AdoAo_solicitud_Lista.Recordset!ges_gestion = gestion1    'adosolicitud.Recordset("ges_gestion")     'Year(Date)
      AdoAo_solicitud_Lista.Recordset!codigo_unidad = uni_codigo1   'adosolicitud.Recordset("codigo_UNIDAD")    'Trim(DtcUnidad.Text)
      AdoAo_solicitud_Lista.Recordset!codigo_solicitud = COD_SOL    'adosolicitud.Recordset("codigo_solicitud") 'Trim(txtnrosol.Text)
'      adoao_solicitud_lista.Recordset!id_beneficiario = id_beneficiario1
    End If
'    If swnuevo = 2 Then
'      rstdestino.Open "select * from ao_solicitud_lista where codigo_unidad = '" & lblcodigo_unidad & "' and codigo_solicitud = " & lblcodigo_solicitud & " and id_beneficiario = " & adoao_solicitud_lista.Recordset!id_beneficiario, db, adOpenKeyset, adLockOptimistic
'    End If
      AdoAo_solicitud_Lista.Recordset!CodGrupo = DtcCodGrupo.Text                       'Grupo Bien
      AdoAo_solicitud_Lista.Recordset!cod_MONTADOR = DtcSubgrupo.Text                   'Sub-Grupo Bien
      AdoAo_solicitud_Lista.Recordset!codDetalle = Trim(Dtccodbien.Text)                'Codigo Bien
      AdoAo_solicitud_Lista.Recordset!doc_identidad = Trim(Dtccodbien.Text)             'Codigo de Bien Copia
      AdoAo_solicitud_Lista.Recordset!profesion = Trim(Dtcdesbien.Text)                 'Descripcion del Bien
      AdoAo_solicitud_Lista.Recordset!descdetalle = Trim(Txtrazon_s.Text)               'Descripcion del Bien + Caracteristicas
      AdoAo_solicitud_Lista.Recordset!razon_s = Trim(Txtrazon_s)                        'Descripcion del Bien + Caracteristicas (Copia)
      AdoAo_solicitud_Lista.Recordset!grado_instruccion = DtcdesAnt                     'Nombre Antiguo del Producto (Copia)
      AdoAo_solicitud_Lista.Recordset!Nombre_Anterior = DtcdesAnt                       'Nombre Antiguo del Producto
      AdoAo_solicitud_Lista.Recordset!aplanilla = IIf(TxtCantidad = "", 1, TxtCantidad) 'Cantidad Solicitada (Copia)
      AdoAo_solicitud_Lista.Recordset!cantidad = IIf(TxtCantidad = "", 1, TxtCantidad)  'Cantidad Solicitada del Bien
      AdoAo_solicitud_Lista.Recordset!Nro_pagos = IIf(TxtPrecioU = "", 0, CDbl(TxtPrecioU))            'Precio Unitario Actual (Copia)
      AdoAo_solicitud_Lista.Recordset!Precio_Compra = IIf(TxtPrecioU = "", 0, CDbl(TxtPrecioU))        'Precio Unitario Actual para la solicitud
      AdoAo_solicitud_Lista.Recordset!Total_compra = AdoAo_solicitud_Lista.Recordset!cantidad * AdoAo_solicitud_Lista.Recordset!Precio_Compra   'Precio Total Actual para la solicitud
      AdoAo_solicitud_Lista.Recordset!Precio_venta = ((AdoAo_solicitud_Lista.Recordset!Precio_Compra) * (1 - Val(Txt_porcentaje) / 100))        'Precio Unitario con el % de Descuento
      AdoAo_solicitud_Lista.Recordset!total_venta = AdoAo_solicitud_Lista.Recordset!cantidad * AdoAo_solicitud_Lista.Recordset!Precio_venta     'Precio Total con el % de Descuento
      'adoao_solicitud_lista.Recordset!Monto_solicitud_dl = adoao_solicitud_lista.Recordset!cantidad * adoao_solicitud_lista.Recordset!Precio_venta      'Precio Total (Copia)
      AdoAo_solicitud_Lista.Recordset!Monto_solicitud_dl = AdoAo_solicitud_Lista.Recordset!Precio_venta / GlTipoCambioOficial                   'Precio Total Dolares con el % de Descuento
      AdoAo_solicitud_Lista.Recordset!Unidad = Dtc_UniMed.Text                          'Unidad de Medidad del Bien
      AdoAo_solicitud_Lista.Recordset!aunidad = Dtc_UniMed.Text                         'Unidad de Medidad del Bien (Copia)
      AdoAo_solicitud_Lista.Recordset!Precio_salon = IIf(DtcPrecioU.Text = "", CDbl(TxtPrecioU), DtcPrecioU.Text)                    'Precio Venta Base Intermediario
      AdoAo_solicitud_Lista.Recordset!Precio_estimado = IIf(DtcPrecioUV.Text = "", CDbl(TxtPrecioU), DtcPrecioUV.Text)                'Precio Venta Cliente Final
      
      AdoAo_solicitud_Lista.Recordset!tipo_beneficiario = "F"  'Trim(lbltipo_beneficiario)
      AdoAo_solicitud_Lista.Recordset!usr_usuario = GlUsuario
      AdoAo_solicitud_Lista.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
      AdoAo_solicitud_Lista.Recordset!hora_registro = Format(Time, "hh:mm:ss")
      AdoAo_solicitud_Lista.Recordset.Update
      'JQA 04/2008
'     adoao_solicitud_lista.Recordset = rstdestino
      db.CommitTrans
    db.Execute " UPDATE AO_SOLICITUD SET Lista_adjunta='S' " & _
            "WHERE (ao_Solicitud.Ges_Gestion) = '" & adosolicitud.Recordset!ges_gestion & "' and " & _
            "(ao_Solicitud.codigo_unidad) = '" & adosolicitud.Recordset!codigo_unidad & "' and " & _
            "(ao_Solicitud.codigo_solicitud) =  " & adosolicitud.Recordset!codigo_solicitud & ""
    'adosolicitud.Recordset("Lista_adjunta") = "S"
    If swnuevo = 1 Then
'     Call abre_solicitud_lista
      'parametro = "ges_gestion > '2010'"
      parametro = "ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
      Call ABRE_SOL_LISTA
      AdoAo_solicitud_Lista.Recordset.MoveLast
    End If
    If swnuevo = 2 Then
      marca1 = AdoAo_solicitud_Lista.Recordset.Bookmark
'     Call abre_solicitud_lista
'     rstao_solicitud_lista.Update
'     rstao_solicitud_lista.Requery
'     Set adoao_solicitud_lista.Recordset = rstao_solicitud_lista
      If rstao_solicitud_lista.RecordCount > 0 Then
        parametro = "ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
        Call ABRE_SOL_LISTA
        AdoAo_solicitud_Lista.Recordset.Move marca1 - 1
      End If
    End If
    Frmnavega.Enabled = True
    Frmnavega.Visible = True
    Frame10.Enabled = True
    FrmEditaDet.Enabled = False
    'BtnAñadirDet.Enabled = True
    'BtnModificarDet.Enabled = True
    swnuevo = 0
    'rstAo_solicitud!Lista_adjunta = "S"
    'Call GRABADET
    sstab1.Tab = 0
    sstab1.TabEnabled(1) = False
    sstab1.TabEnabled(0) = True
    SOLISTA = "A"
    TxtPrecioU.Enabled = True
    TxtPrecioC.Enabled = False
    FrmGrabaDet.Visible = False
    FrmABMDet.Visible = True
  Else
    MsgBox "Error en registro de datos, vuelva a intentar.!", vbCritical, ""
  End If
 End If
End Sub

Private Sub BtnModificarDet_Click()
  'MsgBox "Cod: " + adoao_solicitud_lista.Recordset!codDetalle
  If adosolicitud.Recordset("estado_enviado") = "N" And adosolicitud.Recordset!Lista_adjunta = "S" Then
    'marca1 = adosolicitud.Recordset.BookMark
    'marca1 = adoao_solicitud_lista.Recordset.BookMark
    Frmnavega.Enabled = False
    Frmnavega.Visible = False
    Frame10.Enabled = False
    swnuevo = 2
    'adoao_solicitud_lista.Recordset.Move marca1 - 1
    sstab1.Tab = 1
    sstab1.TabEnabled(1) = True
    sstab1.TabEnabled(0) = False
    FrmEditaDet.Visible = True
    FrmEditaDet.Enabled = True
    FrmGrabaDet.Visible = True
    FrmABMDet.Visible = False
    'If GlSistema = "C" Or GlSistema = "Z" Then
        TxtPrecioU.Enabled = True
        TxtPrecioC.Enabled = False
    'End If
  Else
    MsgBox "No se puede Modificar un registro Aprobado, Enviado o Inexistente!! ", vbExclamation
  End If
End Sub

Private Sub ChkTdr_Click()
    If ChkTdr.Value = 0 Then txtterref.Text = "N"
    If ChkTdr.Value = 1 Then txtterref.Text = "S"
    If ChkTdr.Value = 2 Then txtterref.Text = "E"
End Sub

Private Sub BtnAprobar_Click()
If adosolicitud.Recordset!Lista_adjunta = "S" Then
    sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
    If sino = vbYes Then
        Dim rstdestino As New ADODB.Recordset
        Set rstdestino = New ADODB.Recordset
        If rstdestino.State = 1 Then rstdestino.Close
        rstdestino.Open "select * from Ao_solicitud where ges_gestion = '" & adosolicitud.Recordset("ges_gestion") & "' and formulario = '" & adosolicitud.Recordset("formulario") & "' and codigo_solicitud = " & adosolicitud.Recordset("codigo_solicitud") & " and codigo_unidad = '" & adosolicitud.Recordset("codigo_unidad") & "'", db, adOpenDynamic, adLockOptimistic
        If Not rstdestino.BOF Then rstdestino.MoveFirst
        If Not rstdestino.BOF And Not rstdestino.EOF Then
            rstdestino("estado_aprobado") = "S"
            rstdestino.Update
        End If
        If rstdestino.State = 1 Then rstdestino.Close
        marca1 = adosolicitud.Recordset.Bookmark
        Call OptFilGral1_Click
        'adosolicitud.Recordset.Requery
        'adosolicitud.Refresh
        adosolicitud.Recordset.Move marca1 - 1
    End If
Else
    MsgBox "No se puede APROBAR. Debe registrar el detalle del registro ...", , "Atención"
End If
End Sub

Private Sub BtnBuscar_Click()
'JQA
'  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
'  Dim ClBuscaSec As ClBuscaSecuencialEnRS
  PosibleApliqueFiltro = False
  Dim rsNada As ADODB.Recordset
  Dim GrSqlAux As String
  Set ClBuscaGrid = New ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.EsTdbGrid = False
  Set ClBuscaGrid.GridTrabajo = DataGrid1
  ClBuscaGrid.QueryUtilizado = queryinicial
  Set ClBuscaGrid.RecordsetTrabajo = adosolicitud.Recordset
  ClBuscaGrid.CamposVisibles = "110"
  ClBuscaGrid.Ejecutar
  PosibleApliqueFiltro = True
End Sub

Private Sub BtnDesAprobar_Click()
  If adosolicitud.Recordset!estado_enviado = "S" Then
    MsgBox "No se puede DESAPROBAR si el registro está ENVIADO ...", vbCritical, "Advertencia !"
  Else
    If adosolicitud.Recordset!ESTADO_APROBADO = "S" Then
       sino = MsgBox("Esta seguro de DESAPROBAR el registro ?", vbYesNo, "Confirmando")
       If sino = vbYes Then
          Dim rstdestino As New ADODB.Recordset
          Set rstdestino = New ADODB.Recordset
          If rstdestino.State = 1 Then rstdestino.Close
          rstdestino.Open "select * from Ao_solicitud where ges_gestion = '" & adosolicitud.Recordset("ges_gestion") & "' and formulario = '" & adosolicitud.Recordset("formulario") & "' and codigo_solicitud = " & adosolicitud.Recordset("codigo_solicitud") & " and codigo_unidad = '" & adosolicitud.Recordset("codigo_unidad") & "'", db, adOpenDynamic, adLockOptimistic
          If Not rstdestino.BOF Then rstdestino.MoveFirst
          If Not rstdestino.BOF And Not rstdestino.EOF Then
            rstdestino("estado_aprobado") = "N"
            rstdestino.Update
          End If
          If rstdestino.State = 1 Then rstdestino.Close
          marca1 = adosolicitud.Recordset.Bookmark
          'adosolicitud.Recordset.AddNew
          adosolicitud.Recordset.Cancel
          adosolicitud.Refresh
          adosolicitud.Recordset.Move marca1 - 1
       End If
    Else
        MsgBox "No se puede DESAPROBAR si el registro NO está APROBADO ...", vbCritical, "Advertencia !"
    End If
  End If
End Sub

'Private Sub CmdDetallePoa_Click()
'  If adosolicitud.Recordset.RecordCount > 0 Then
'  marca1 = adosolicitud.Recordset.BookMark
'   ''''' ALB
'  FrmPoasCapturaALB.Lblformulario = "F01"
'  FrmPoasCapturaALB.lblges_gestion = adosolicitud.Recordset!ges_gestion
'  FrmPoasCapturaALB.lblcodigo_unidad = adosolicitud.Recordset!codigo_unidad
'  FrmPoasCapturaALB.lblcodigo_solicitud = adosolicitud.Recordset!codigo_solicitud
'  FrmPoasCapturaALB.lbltipo_beneficiario = "N" 'adosolicitud.Recordset!tipo_beneficiario
'  FrmPoasCapturaALB.tXTaprobado = adosolicitud.Recordset!aprobado
'  FrmPoasCapturaALB.Lbltipo_bien_Cta_doc = adosolicitud.Recordset!tipo_bien_Cta_doc
'  FrmPoasCapturaALB.Lblcategoria_Cta_doc = adosolicitud.Recordset!subcta2
'  FrmPoasCapturaALB.Show vbModal
'  adosolicitud.Refresh
'  Else
'    MsgBox "No Existen Registros ", vbInformation, "Formulario 1"
'  End If
'  If adosolicitud.Recordset.RecordCount > 0 Then
'    adosolicitud.Recordset.Move marca1 - 1
'  End If
'End Sub

Private Sub BtnEnviar_Click()
    If adosolicitud.Recordset!ESTADO_APROBADO = "S" Then
      swunidad = 0
      sino = MsgBox("Esta seguro de ENVIAR el registro Aprobado ? (Nota: Ya no se podrá Desaprobar!)...", vbYesNo, "Confirmando ...")
      If sino = vbYes Then
        'JQA 04/2008
        Call GRABADET
        CODPAG = 0
        marca1 = adosolicitud.Recordset.Bookmark
        Call val_presupF04(adosolicitud.Recordset, GlNombFor)
        Set rs_soldetaux = New ADODB.Recordset
        If rs_soldetaux.State = 1 Then rs_soldetaux.Close
        rs_soldetaux.Open "select ges_gestion, codigo_unidad, codigo_solicitud, org_codigo_contra, sum(monto_bolivianos) as monto_bolivianos, sum(monto_dolares) as monto_dolares, sum(monto_bolivianos_contra) as monto_bolivianos_contra, sum(monto_dolares_contra) as monto_dolares_contra, tipo_moneda, codigo_convenio, codigo_poa from ao_solicitud_detalle where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & "  GROUP BY ges_gestion, codigo_unidad, codigo_solicitud, org_codigo_contra, tipo_moneda, codigo_convenio, codigo_poa ", db, adOpenKeyset, adLockOptimistic
        If rs_soldetaux.RecordCount > 0 Then
            'deCD.dbo_ap_Graba_No_Objecion_D2 adosolicitud.Recordset!ges_gestion, adosolicitud.Recordset!codigo_solicitud, adosolicitud.Recordset!TipoF1, conv_nal, "10", "00", "00", "00", adosolicitud.Recordset!caracteristicas, adosolicitud.Recordset!observaciones, GlUsuario, Format(Date, "dd/mm/yyyy"), Format(Time, "hh:mm:ss"), rs_soldetaux!org_codigo_contra, 0, rstAo_solicitud!formulario, 0, "D", "10", rstAo_solicitud!codigo_unidad, rstAo_solicitud!codigo_unidad, rs_soldetaux!monto_dolares, rs_soldetaux!MONTO_DOLARES_CONTRA, rs_soldetaux!monto_bolivianos, rs_soldetaux!monto_bolivianos_contra, rs_soldetaux!tipo_moneda, "0"
            ' CORREGIR 06/03/2012 ADALID
            deCD.dbo_ap_Graba_No_Objecion_D2 adosolicitud.Recordset!ges_gestion, adosolicitud.Recordset!codigo_solicitud, adosolicitud.Recordset!tipoF1, rs_soldetaux!codigo_convenio, "10", "00", "00", "00", adosolicitud.Recordset!caracteristicas, adosolicitud.Recordset!observaciones, GlUsuario, Format(Date, "dd/mm/yyyy"), Format(Time, "hh:mm:ss"), rs_soldetaux!org_codigo_contra, 0, rstAo_solicitud!formulario, 0, "D", "10", rstAo_solicitud!codigo_unidad, adosolicitud.Recordset!codigo_unidad, rs_soldetaux!monto_dolares, rs_soldetaux!monto_dolares_contra, rs_soldetaux!monto_bolivianos, rs_soldetaux!monto_bolivianos_contra, rs_soldetaux!tipo_moneda, rs_soldetaux!codigo_poa
            db.Execute " UPDATE AO_SOLICITUD SET estado_enviado='S' " & _
            "WHERE (ao_Solicitud.Ges_Gestion) = '" & Me.adosolicitud.Recordset!ges_gestion & "' and " & _
            "(ao_Solicitud.codigo_unidad) = '" & Me.adosolicitud.Recordset!codigo_unidad & "' and " & _
            "(ao_Solicitud.codigo_solicitud) =  " & Me.adosolicitud.Recordset!codigo_solicitud & ""
            'db.Execute "update FC_BENEFICIARIO set EMITIDO= 'NO' Where codigo_beneficiario = '" & adosolicitud.Recordset!codigo_beneficiario & "' And (EMITIDO <> 'NO' And EMITIDO <> 'SI'"
            db.Execute "update FC_BENEFICIARIO set EMITIDO= 'NO' Where codigo_beneficiario = '" & adosolicitud.Recordset!CI_aprueba & "' And (EMITIDO <> 'NO' And EMITIDO <> 'SI')"
            adosolicitud.Refresh
            If marca1 > 1 Then
                adosolicitud.Recordset.Move marca1 - 1
            End If
'            db.Execute "update AlCldetalle set AlCldetalle.stockingreso= av_acumula_compra.cantidad_cotizada from AlCldetalle, av_acumula_compra Where AlCldetalle.CodGrupo = av_acumula_compra.CodGrupo And AlCldetalle.cod_MONTADOR = av_acumula_compra.cod_MONTADOR And AlCldetalle.codDetalle = av_acumula_compra.codDetalle"
'            db.Execute "update AlCldetalle set StockActual= Stockinicial + stockingreso - StockSalida"
            
        Else
            MsgBox "NO se registro el detalle del registro, intente nuevamente...", vbExclamation, "-"
        End If
        adosolicitud.Refresh
        'JQA 04/2008
      End If
    Else
        MsgBox "No se puede ENVIAR. Debe Aprobar previamente el registro ...", , "Atención"
    End If
End Sub

Private Sub CmdImporta_Click()
'    FrmImporta.Show
End Sub

Private Sub BtnAñadir_Click()
    Frame3.Enabled = True
    Frame10.Enabled = False
    Frame10.Visible = False
'    Frame1.Visible = False
    frmabm.Visible = False
    Frmnavega.Enabled = False
    frmgrabcabeza.Visible = True
    Frasolic.Enabled = True
    swgrabar = 1
    Call cerea
    adosolicitud.Refresh
    adosolicitud.Recordset.AddNew
    'FrmApertura.Visible = True
    DTPfechasol.CheckBox = True
    DTPfechasol.Value = Format(Date, "dd/mm/yyyy")
    DTPfechasol.CheckBox = False
    Txt_porcentaje.Text = 0
    Txtcaracteristicas.Text = "SOLICITUD DE COMPRA REGULAR"
    txtjustifica.Text = "NINGUNA"
    Option1.Visible = True
    Option2.Visible = True
    DtCUnidad.Enabled = True
    Dtcpaternobe.Enabled = False
'    If GlSistema = "A" Or GlSistema = "Z" Then
'        DtCUnidad.Text = "LINEAZ"
'        Dtccodbien.Text = "27"
'        Dtcdesbien.Text = "DIVISION PROFESIONAL - LINEA Z"
'        DtcPOA.Text = "2.3.1.2.1"
'        DtcPOADes.Text = "Productos de Belleza LINEA Z"
'        Dtccibe.Text = "9000000001"
'        Dtcpaternobe.Text = "EMPRESA S.A."
'    End If
    DataGrid1.Visible = False
    Lbltipo_bien_Cta_doc.Visible = False
    Lbltipo_bien_Cta_doc.Caption = "APERTURA"
  
    sstab1.Tab = 0
    sstab1.TabEnabled(0) = True
    sstab1.TabEnabled(1) = False
    
End Sub

Private Sub CmdGraBoleta_Click()
'  FraModBoleta.Visible = False
End Sub

Private Sub BtnImprimirA_Click()
'JQA JUN/2008
If adosolicitud.Recordset!Lista_adjunta = "S" Then
    Dim co As New ADODB.Command
'    Dim rs As New ADODB.Recordset
'    rs.Open "select * from ao_solicitud_detalle where ges_gestion='" & Me.adosolicitud.Recordset!ges_gestion & "' and " & _
'            "codigo_unidad='" & Me.adosolicitud.Recordset!codigo_unidad & "' and " & _
'            "codigo_solicitud=" & Me.adosolicitud.Recordset!codigo_solicitud, db, adOpenStatic, adLockReadOnly
    CryF01.ReportFileName = App.Path & "\Reportes\Solicitudes\C01_F11.rpt"
    CryF01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
      CryF01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
      CryF01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
    'Call CREAVISTAF11          'JQA JUN-2008
    CryF01.StoredProcParam(0) = Me.adosolicitud.Recordset!ges_gestion
    CryF01.StoredProcParam(1) = Me.adosolicitud.Recordset!codigo_unidad
    CryF01.StoredProcParam(2) = Me.adosolicitud.Recordset!codigo_solicitud
    iResult = CryF01.PrintReport
    If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar el detalle del registro ...", , "Atención"
End If
' ... (Jorge)
'  Dim V_cmbSubCta2 As String
'  Dim IResult As Variant
'  Dim PaternoS, MaternoS, NombreS, PaternoB, MaternoB, NombreB, UnidadT As String
'  Dim rsunidad As New ADODB.Recordset
'  Set rsunidad = New ADODB.Recordset
'
'  adoao_solicitud_detalle.Refresh
'  '---- ini version actual
''  db.Execute "drop view av_F01"
''  db.Execute "create view av_F01 as SELECT ao_Solicitud.ges_gestion, ao_Solicitud.codigo_unidad, ao_Solicitud.codigo_solicitud, ao_Solicitud.formulario, ao_Solicitud.justificacion_solicitud, ao_Solicitud.CI, ao_Solicitud.fecha_solicitud, ao_Solicitud_detalle.tipo_moneda, ao_Solicitud_detalle.monto_bolivianos, ao_Solicitud_detalle.monto_dolares, ao_Solicitud.fecha_registro, ao_Solicitud.CI_aprueba, ao_Solicitud_detalle.monto_bolivianos_contra, ao_Solicitud_detalle.monto_dolares_contra, ao_Solicitud_detalle.Tipo_cambio " & _
''            " FROM ao_Solicitud INNER JOIN ao_Solicitud_detalle ON (ao_Solicitud.codigo_solicitud = ao_Solicitud_detalle.codigo_solicitud) AND (ao_Solicitud.codigo_unidad = ao_Solicitud_detalle.codigo_unidad) AND (ao_Solicitud.Ges_Gestion = ao_Solicitud_detalle.Ges_Gestion) " & _
''            " WHERE (((ao_Solicitud.codigo_solicitud)= " & txtnrosol & ") AND ((ao_Solicitud.codigo_unidad)='" & Trim(lblUni_codigo) & "')) "
'  '---- fin version actual
'
'  db.Execute "drop view av_F01"
'  db.Execute "create view av_F01 as SELECT ao_Solicitud.ges_gestion, ao_Solicitud.codigo_unidad, ao_Solicitud.codigo_solicitud, ao_Solicitud.formulario, ao_Solicitud.justificacion_solicitud, ao_Solicitud.CI, ao_Solicitud.fecha_solicitud, ao_Solicitud_detalle.tipo_moneda, ao_Solicitud_detalle.monto_bolivianos, ao_Solicitud_detalle.monto_dolares, ao_Solicitud.fecha_registro, ao_Solicitud.CI_aprueba, ao_Solicitud.estatus, ao_Solicitud.estado_aprobacion, ao_Solicitud_detalle.monto_bolivianos_contra, ao_Solicitud_detalle.monto_dolares_contra, ao_Solicitud_detalle.Tipo_cambio , ao_Solicitud_detalle.codigo_convenio, por_fte_ext, por_fte_nal  " & _
'            " FROM ao_Solicitud INNER JOIN ao_Solicitud_detalle ON (ao_Solicitud.codigo_solicitud = ao_Solicitud_detalle.codigo_solicitud) AND (ao_Solicitud.codigo_unidad = ao_Solicitud_detalle.codigo_unidad) AND (ao_Solicitud.Ges_Gestion = ao_Solicitud_detalle.Ges_Gestion) " & _
'            " WHERE (((ao_Solicitud.codigo_solicitud)= " & txtnrosol & ") AND ((ao_Solicitud.codigo_unidad)='" & Trim(lblUni_codigo) & "')) "
'
'    'and codigo_unidad='" & adosolicitud.Recordset!codigo_unidad & "'
'
'  rsunidad.Open "select * from fc_unidad_ejecutora where codigo_unidad='" & adosolicitud.Recordset("codigo_unidad") & "' ", db, adOpenKeyset, adLockReadOnly
'  CryF01.WindowShowRefreshBtn = True
'  If rsunidad.RecordCount > 0 Then
'     CryF01.Formulas(0) = "UnidadT = '" & rsunidad("uni_descripcion_larga") & "' "
'  Else
'     CryF01.Formulas(0) = "UnidadT = '" & "-" & "' "
'  End If
'  CryF01.Formulas(1) = "PaternoS = '" & Dtcpaternosol.Text & "' "
'  CryF01.Formulas(2) = "MaternoS = '" & dtcmaternosol.Text & "' "
'  CryF01.Formulas(3) = "NombreS = '" & dtcnombresol.Text & "' "
'  CryF01.Formulas(4) = "PaternoB = '" & Dtcpaternobe.Text & "' "
'  CryF01.Formulas(5) = "MaternoB = '" & Dtcmaternobe.Text & "' "
'  CryF01.Formulas(6) = "NombreB = '" & Dtcnombrebe.Text & "' "
'  CryF01.Formulas(7) = "Tipo = '" & Lbltipo_bien_Cta_doc.Caption & "' "
'  V_cmbSubCta2 = " : " & cmbSubCta2.Text
'  CryF01.Formulas(14) = "tipof1 = '" & DtCvalor1.Text & " " & IIf(cmbSubCta2.Visible = False, "", V_cmbSubCta2) & "' "
'  If cmbSubCta2.Visible = True And cmbSubCta2.Text = "PASE" Then
'    CryF01.Formulas(15) = "titmunicipio = 'UNIDAD EDUCATIVA:' "
'    CryF01.Formulas(16) = "codmuni = '" & adoao_solicitud_detalle.Recordset!aux3 & "' "
'    CryF01.Formulas(17) = "desmuni = '" & fbusmuni(adoao_solicitud_detalle.Recordset!aux3) & "' "
'  Else
'    CryF01.Formulas(15) = "titmunicipio = ''"
'    CryF01.Formulas(16) = "codmuni = '' "
'    CryF01.Formulas(17) = "desmuni = '' "
'  End If
'      If Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "C" Or Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "CC" Then
''        FraBoleta.Visible = True
'      Else
''        FraBoleta.Visible = False
'      End If
'
'  If Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "C" Or Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "CC" Then
'    CryF01.Formulas(8) = "boletaTit = 'BOLETA BANCARIA :'"
'
'
'    ' aqui 30/05/2001
''    CryF01.Formulas(9) = "boletanumero = '" & Txtnro_boleta.Text & "' " 'TxtPlanilla_depto
'    ' aqui 30/05/2001
'
'
''    CryF01.Formulas(10) = "boletaCtaTit = 'Cuenta :'"
''    CryF01.Formulas(11) = "boletaCta = '" & DtCcta_codigo.Text & "' "  'DtCBco_codigo.Text & "' "
''    CryF01.Formulas(12) = "boletamontoTit = 'Monto Bs.:'"
''    CryF01.Formulas(13) = "boletamonto = '" & TDBNmontoBs & "' " 'TDBNnro_pagos & "' "
'  Else
'    CryF01.Formulas(8) = "boletaTit = ' '"
'    CryF01.Formulas(9) = "boletanumero = ' ' "
'    CryF01.Formulas(10) = "boletaCtaTit = ' '"
'    CryF01.Formulas(11) = "boletaCta = ' '"
'    CryF01.Formulas(12) = "boletamontoTit = ' '"
'    CryF01.Formulas(13) = "boletamonto = ' '"
'  End If
'  CryF01.ReportFileName = App.Path & "\FormulariosEntrada\S01_F01.rpt"
'  IResult = CryF01.PrintReport
'  If IResult <> 0 Then
'     MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical + vbOKOnly, "Error..."
'  End If
End Sub

Private Sub BtnAddDetalle_Click()
  marca1 = adosolicitud.Recordset.Bookmark
  If adosolicitud.Recordset("estado_enviado") = "N" Then
    'marca1 = adosolicitud.Recordset.BookMark
'    Call cerea
    swnuevo = 1
    Frmnavega.Enabled = False
    Frmnavega.Visible = False
    Frame10.Enabled = False
    FrmEditaDet.Visible = True
    FrmEditaDet.Enabled = True
    FrmGrabaDet.Visible = True
    FrmABMDet.Visible = False
    gestion1 = adosolicitud.Recordset("ges_gestion")
    uni_codigo1 = adosolicitud.Recordset("CODIGO_UNIDAD")
    COD_SOL = adosolicitud.Recordset("codigo_solicitud")
    'adosolicitud.Recordset.Move marca1 - 1
   'parametro = "ges_gestion" + " <> " + "'2000'"
   parametro2 = "Cod_marca = '" & DtcMarca.Text & "' "
   parametro = "ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
   Call ABRE_SOL_LISTA
   sstab1.Tab = 1
    sstab1.TabEnabled(1) = True
    sstab1.TabEnabled(0) = False
    
   ' Call ABRE_SOL_LISTA
    AdoAo_solicitud_Lista.Recordset.AddNew
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnModDetalle_Click()
    Frame3.Enabled = True
    Frame10.Enabled = False
    Frame10.Visible = False
'    txtnrosol.Enabled = False
    DTPfechasol.SetFocus
    DTPfechasol.CheckBox = True
    frmabm.Visible = False
    Frmnavega.Enabled = False
    frmgrabcabeza.Visible = True
    DtCUnidad.Enabled = False
    Dtcpaternobe.Enabled = False
    Option1.Visible = True
    Option2.Visible = True
    swgrabar = 2
    sstab1.Tab = 0
    sstab1.TabEnabled(0) = True
    sstab1.TabEnabled(1) = False
    lblUni_codigo.Caption = DtCUnidad.Text
    
    Set rstpoaAux = New ADODB.Recordset
    If rstpoaAux.State = 1 Then rstpoaAux.Close
    rstpoaAux.Open "select * from fc_Relacionador_poa_ppto where (codigo_unidad = '" & DtCUnidad.Text & "') and (nivel  = 5)  ", db, adOpenKeyset, adLockReadOnly
    If rstpoaAux.RecordCount > 0 Then
        If rstpoaAux!codigo_poa = "" Then
            DtcPOA.Text = "1.1.1.2.1"
            DtcPOADes.Text = "INSUMOS Y MATERIALES ODONTOLOGICOS"
        Else
            rstpoaAux.MoveLast
            DtcPOA.Text = rstpoaAux!codigo_poa
            DtcPOADes.Text = rstpoaAux!DESCRIPCION_POA
        End If
        Set AdoPOA.Recordset = rstpoaAux
        AdoPOA.Refresh
    End If
End Sub

Private Sub BtnEliminar_Click()
If adosolicitud.Recordset!estado_enviado = "N" Then
  Dim rsterr As New ADODB.Recordset
    sino = MsgBox("Está seguro de eliminar este registro", vbYesNo + vbQuestion, "Atención")
    If sino = vbYes Then
      Set rsterr = New ADODB.Recordset
      If rsterr.State = 1 Then rsterr.Close
      rsterr.Open "select * from Ao_solicitud where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
      If rsterr.RecordCount > 0 Then
        rsterr!ESTADO_APROBADO = "E"
        rsterr!estado_enviado = "E"
        rsterr.Update
      End If
      If rsterr.State = 1 Then rsterr.Close
      'marca1 = adosolicitud.Recordset.BookMark
      'rstAo_solicitud.Requery
      Call OptFilGral1_Click
      'Set adosolicitud.Recordset = rstAo_solicitud
      'adosolicitud.Refresh
      'Set adosolicitud.Recordset = rstAo_solicitud
'      If marca1 > 1 Then
'        adosolicitud.Recordset.Move marca1 - 1
'      End If
    End If
Else
    MsgBox "No se puede ANULAR. El registro ya esta Aprobado y Enviado ...", , "Atención"
End If
End Sub

Private Sub BtnSalir_Click()
    sino = MsgBox("Esta Seguro de Salir?", vbQuestion + vbYesNo, "Confirmando...")
    If sino = vbYes Then
'        adosolicitud.Recordset.Close
        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
        If rstpoa.State = 1 Then rstpoa.Close
        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
        If rstfc_partida_gasto.State = 1 Then rstfc_partida_gasto.Close
'        If rstao_solicitud_detalle.State = 1 Then rstao_solicitud_detalle.Close
'        If rstAo_solicitud.State = 1 Then rstAo_solicitud.Close
        Unload Me
    End If
End Sub

'Private Sub CmdDetCabeza_Click()
'    frmabm.Visible = False
'    frmdetalle.Visible = True
'    FraDetalle.Visible = True
'    Frmnavega.Enabled = False
'    If Not (Adodetallesolicitud.Recordset.BOF) Then Adodetallesolicitud.Recordset.MoveFirst
'
'End Sub

Private Sub BtnGrabar_Click()
    If DtCUnidad.Text = "" Then
        MsgBox "Error, Debe ELEGIR la Unidad Productiva ..."
        DtCUnidad.SetFocus
        Exit Sub
    End If
    If DtcTipo.Text = "" Then
        MsgBox "Error, Debe ELEGIR el Tipo de Solicitud ..."
        DtcTipoS.SetFocus
        Exit Sub
    End If
    If dtccisol.Text = "" Then
        MsgBox "Error, Debe ELEGIR el Responsable de la Solicitud ..."
        Dtcpaternosol.SetFocus
        Exit Sub
    End If
    If Dtccibe.Text = "" Then
        MsgBox "Error, Debe ELEGIR el Proveedor o Proponente ..."
        Dtcpaternobe.SetFocus
        Exit Sub
    End If
'    If DtcPOA.Text = "" Then
'        MsgBox "Error, Debe ELEGIR la Actividad del POA ..."
'        DtcPOADes.SetFocus
'        Exit Sub
'    End If
    If Txtcaracteristicas.Text = "" Then
        MsgBox "Error, Debe REGISTRAR las Caracteristicas Generales de la solicitud ..."
        Txtcaracteristicas.SetFocus
        Exit Sub
    End If
    Frame3.Enabled = False
    Call grabar
    Call proveedor
    Frame10.Enabled = True
    Frame10.Visible = True
    DataGrid1.Visible = True
    frmabm.Visible = True
    frmgrabcabeza.Visible = False
    Frmnavega.Enabled = True
    Frame3.Enabled = False
    FrmApertura.Visible = False
    Frasolic.Enabled = True
    DtCUnidad.Enabled = True
    Option1.Visible = False
    Option2.Visible = False
    swgrabar = 0
    sstab1.Tab = 0
    sstab1.TabEnabled(0) = True
    sstab1.TabEnabled(1) = True
End Sub

Private Sub BtnCancelar_Click()
    adosolicitud.Refresh
    frmabm.Visible = True
'    frmdetalle.Visible = False
    frmgrabcabeza.Visible = False
    If adosolicitud.Recordset.RecordCount > 0 Then
        adosolicitud.Recordset.CancelUpdate
    End If
    Call proveedor
    Frmnavega.Enabled = True
    Frame3.Enabled = False
    FrmApertura.Visible = False
    DataGrid1.Visible = True
    DtCUnidad.Enabled = True
    Option1.Visible = False
    Option2.Visible = False
    Frame10.Enabled = True
    Frame10.Visible = True
    swgrabar = 0
    adosolicitud.Refresh
    sstab1.Tab = 0
    sstab1.TabEnabled(0) = True
    sstab1.TabEnabled(1) = True
End Sub

Private Sub Dtc_UniMed_Click(Area As Integer)
    Dtccodbien.BoundText = Dtc_UniMed.BoundText
    Dtcdesbien.BoundText = Dtc_UniMed.BoundText
    DtcPrecioU.BoundText = Dtc_UniMed.BoundText
    DtcPrecioUV.BoundText = Dtc_UniMed.BoundText
    DtcdesAnt.BoundText = Dtc_UniMed.BoundText
    DtcCodAnt.BoundText = Dtc_UniMed.BoundText
    DtcCodUniv.BoundText = Dtc_UniMed.BoundText
    DtcPrecioC.BoundText = Dtc_UniMed.BoundText
    DtcCodGrupoP.BoundText = Dtc_UniMed.BoundText
    DtcSubgrupoP.BoundText = Dtc_UniMed.BoundText
End Sub

Private Sub Dtccibe_Click(Area As Integer)
    Dtcpaternobe.BoundText = Dtccibe.BoundText
'    Dtcmaternobe.BoundText = Dtccibe.BoundText
'    Dtcnombrebe.BoundText = Dtccibe.BoundText
End Sub

'Private Sub dtccisol_Change()
''  lblUni_codigo = IIf(IsNull(adopuestosol.Recordset("codigo_unidad")) = True, "", adopuestosol.Recordset("codigo_unidad"))
''  Call fbuscaunidad
'End Sub

Private Sub dtccisol_Click(Area As Integer)
    Dtcpaternosol.BoundText = dtccisol.BoundText
'    dtcmaternosol.BoundText = dtccisol.BoundText
'    dtcnombresol.BoundText = dtccisol.BoundText
End Sub

Private Sub DtcCodAnt_Click(Area As Integer)
    Dtccodbien.BoundText = DtcCodAnt.BoundText
    Dtcdesbien.BoundText = DtcCodAnt.BoundText
    DtcPrecioU.BoundText = DtcCodAnt.BoundText
    DtcPrecioUV.BoundText = DtcCodAnt.BoundText
    DtcdesAnt.BoundText = DtcCodAnt.BoundText
    Dtc_UniMed.BoundText = DtcCodAnt.BoundText
    DtcCodUniv.BoundText = DtcCodAnt.BoundText
    DtcPrecioC.BoundText = DtcCodAnt.BoundText
    DtcCodGrupoP.BoundText = DtcCodAnt.BoundText
    DtcSubgrupoP.BoundText = DtcCodAnt.BoundText
End Sub

Private Sub Dtccodbien_Click(Area As Integer)
    Dtcdesbien.BoundText = Dtccodbien.BoundText
    DtcPrecioU.BoundText = Dtccodbien.BoundText
    DtcPrecioUV.BoundText = Dtccodbien.BoundText
    Dtc_UniMed.BoundText = Dtccodbien.BoundText
    DtcdesAnt.BoundText = Dtccodbien.BoundText
    DtcCodAnt.BoundText = Dtccodbien.BoundText
    DtcCodUniv.BoundText = Dtccodbien.BoundText
    DtcPrecioC.BoundText = Dtccodbien.BoundText
    DtcCodGrupoP.BoundText = Dtccodbien.BoundText
    DtcSubgrupoP.BoundText = Dtccodbien.BoundText
End Sub

Private Sub Dtccodbien_LostFocus()
  'If swnuevo = 1 Then
    Txtrazon_s.Text = Trim(Dtcdesbien.Text) + "-" + Trim(DtcdesAnt.Text)
    TxtPrecioU.Text = DtcPrecioU.Text
    'TxtPrecioC.Text = DtcPrecioC.Text
    TxtPrecioC.Text = CDbl(TxtPrecioU) * (1 - Val(Txt_porcentaje) / 100)
  'End If
  DtcCodGrupo.Text = DtcCodGrupoP.Text
  DtcSubgrupo.Text = DtcSubgrupoP.Text
End Sub

Private Sub DtcCodGrupo_Click(Area As Integer)
    DtcGrupo.BoundText = DtcCodGrupo.BoundText
End Sub

Private Sub DtcCodGrupoP_Click(Area As Integer)
    Dtcdesbien.BoundText = DtcCodGrupoP.BoundText
    DtcPrecioU.BoundText = DtcCodGrupoP.BoundText
    DtcPrecioUV.BoundText = DtcCodGrupoP.BoundText
    Dtc_UniMed.BoundText = DtcCodGrupoP.BoundText
    DtcdesAnt.BoundText = DtcCodGrupoP.BoundText
    DtcCodAnt.BoundText = DtcCodGrupoP.BoundText
    DtcCodUniv.BoundText = DtcCodGrupoP.BoundText
    DtcPrecioC.BoundText = DtcCodGrupoP.BoundText
    Dtccodbien.BoundText = DtcCodGrupoP.BoundText
    DtcSubgrupoP.BoundText = DtcCodGrupoP.BoundText
End Sub

Private Sub DtcCodUniv_Click(Area As Integer)
    Dtccodbien.BoundText = DtcCodUniv.BoundText
    Dtcdesbien.BoundText = DtcCodUniv.BoundText
    DtcPrecioU.BoundText = DtcCodUniv.BoundText
    DtcPrecioUV.BoundText = DtcCodUniv.BoundText
    DtcdesAnt.BoundText = DtcCodUniv.BoundText
    Dtc_UniMed.BoundText = DtcCodUniv.BoundText
    DtcCodAnt.BoundText = DtcCodUniv.BoundText
    DtcPrecioC.BoundText = DtcCodUniv.BoundText
    DtcCodGrupoP.BoundText = DtcCodUniv.BoundText
    DtcSubgrupoP.BoundText = DtcCodUniv.BoundText
End Sub

Private Sub DtcdesAnt_Click(Area As Integer)
    Dtccodbien.BoundText = DtcdesAnt.BoundText
    Dtcdesbien.BoundText = DtcdesAnt.BoundText
    Dtc_UniMed.BoundText = DtcdesAnt.BoundText
    DtcPrecioU.BoundText = DtcdesAnt.BoundText
    DtcPrecioUV.BoundText = DtcdesAnt.BoundText
    DtcCodAnt.BoundText = DtcdesAnt.BoundText
    DtcCodUniv.BoundText = DtcdesAnt.BoundText
    DtcPrecioC.BoundText = DtcdesAnt.BoundText
    DtcCodGrupoP.BoundText = DtcdesAnt.BoundText
    DtcSubgrupoP.BoundText = DtcdesAnt.BoundText
End Sub

Private Sub Dtcdesbien_Click(Area As Integer)
    Dtccodbien.BoundText = Dtcdesbien.BoundText
    DtcPrecioU.BoundText = Dtcdesbien.BoundText
    DtcPrecioUV.BoundText = Dtcdesbien.BoundText
    Dtc_UniMed.BoundText = Dtcdesbien.BoundText
    DtcdesAnt.BoundText = Dtcdesbien.BoundText
    DtcCodAnt.BoundText = Dtcdesbien.BoundText
    DtcCodUniv.BoundText = Dtcdesbien.BoundText
    DtcPrecioC.BoundText = Dtcdesbien.BoundText
    DtcCodGrupoP.BoundText = Dtcdesbien.BoundText
    DtcSubgrupoP.BoundText = Dtcdesbien.BoundText
End Sub

Private Sub Dtcdesbien_LostFocus()
  If Dtcdesbien.Text <> "" Then
    Txtrazon_s.Text = Trim(Dtcdesbien.Text) + " - " + Trim(DtcdesAnt.Text)
    TxtPrecioU.Text = DtcPrecioC.Text   'DtcPrecioU.Text
    'TxtPrecioC.Text = DtcPrecioC.Text
    If Val(Txt_porcentaje) > 0 Then
        TxtPrecioC.Text = CDbl(TxtPrecioU) * (1 - Val(Txt_porcentaje) / 100)
    Else
        TxtPrecioC.Text = CDbl(TxtPrecioU)  '* (1 / 100)
    End If
    DtcCodGrupo.Text = DtcCodGrupoP.Text
    DtcSubgrupo.Text = DtcSubgrupoP.Text
  Else
    MsgBox "ERROR, Debe Elegir el PRODUCTO y/o registrar el Precio ..."
  End If
End Sub

Private Sub DtcGrupo_Click(Area As Integer)
    DtcCodGrupo.BoundText = DtcGrupo.BoundText
'    Call pSubGrupo(DtcCodGrupo.BoundText)
End Sub

Private Sub Dtcpaternobe_Click(Area As Integer)
    Dtccibe.BoundText = Dtcpaternobe.BoundText
'    Dtcmaternobe.BoundText = Dtcpaternobe.BoundText
'    Dtcnombrebe.BoundText = Dtcpaternobe.BoundText
End Sub

'Private Sub Dtcpaternosol_Change()
''  lblUni_codigo = IIf(IsNull(adopuestosol.Recordset("codigo_unidad")) = True, "", adopuestosol.Recordset("codigo_unidad"))
''  Call fbuscaunidad
'End Sub

Private Sub Dtcpaternosol_Click(Area As Integer)
'    dtcmaternosol.BoundText = Dtcpaternosol.BoundText
'    dtcnombresol.BoundText = Dtcpaternosol.BoundText
    dtccisol.BoundText = Dtcpaternosol.BoundText
End Sub

Private Sub DtcPOA_Click(Area As Integer)
    DtcPOADes.BoundText = DtcPOA.BoundText
End Sub

Private Sub DtcPOADes_Click(Area As Integer)
    DtcPOA.BoundText = DtcPOADes.BoundText
End Sub

Private Sub DtcPrecioC_Click(Area As Integer)
    Dtccodbien.BoundText = DtcPrecioC.BoundText
    Dtcdesbien.BoundText = DtcPrecioC.BoundText
    Dtc_UniMed.BoundText = DtcPrecioC.BoundText
    DtcdesAnt.BoundText = DtcPrecioC.BoundText
    DtcCodAnt.BoundText = DtcPrecioC.BoundText
    DtcCodUniv.BoundText = DtcPrecioC.BoundText
    DtcPrecioU.BoundText = DtcPrecioC.BoundText
    DtcPrecioUV.BoundText = DtcPrecioC.BoundText
    DtcCodGrupoP.BoundText = DtcPrecioC.BoundText
    DtcSubgrupoP.BoundText = DtcPrecioC.BoundText
End Sub

Private Sub DtcPrecioU_Click(Area As Integer)
    Dtccodbien.BoundText = DtcPrecioU.BoundText
    Dtcdesbien.BoundText = DtcPrecioU.BoundText
    Dtc_UniMed.BoundText = DtcPrecioU.BoundText
    DtcdesAnt.BoundText = DtcPrecioU.BoundText
    DtcCodAnt.BoundText = DtcPrecioU.BoundText
    DtcCodUniv.BoundText = DtcPrecioU.BoundText
    DtcPrecioC.BoundText = DtcPrecioU.BoundText
    DtcCodGrupoP.BoundText = DtcPrecioU.BoundText
    DtcSubgrupoP.BoundText = DtcPrecioU.BoundText
    DtcPrecioUV.BoundText = DtcPrecioU.BoundText
End Sub

Private Sub DtcPrecioUV_Click(Area As Integer)
    Dtccodbien.BoundText = DtcPrecioUV.BoundText
    Dtcdesbien.BoundText = DtcPrecioUV.BoundText
    Dtc_UniMed.BoundText = DtcPrecioUV.BoundText
    DtcdesAnt.BoundText = DtcPrecioUV.BoundText
    DtcCodAnt.BoundText = DtcPrecioUV.BoundText
    DtcCodUniv.BoundText = DtcPrecioUV.BoundText
    DtcPrecioC.BoundText = DtcPrecioUV.BoundText
    DtcCodGrupoP.BoundText = DtcPrecioUV.BoundText
    DtcSubgrupoP.BoundText = DtcPrecioUV.BoundText
    DtcPrecioU.BoundText = DtcPrecioUV.BoundText
End Sub

Private Sub DtcSubgrupo_Click(Area As Integer)
    DtcSubgrupoDes.BoundText = DtcSubgrupo.BoundText
End Sub

Private Sub DtcSubgrupoDes_Click(Area As Integer)
    DtcSubgrupo.BoundText = DtcSubgrupoDes.BoundText
    'Call pProducto(DtcSubgrupo.BoundText)
End Sub

Private Sub DtcSubgrupoP_Click(Area As Integer)
    Dtcdesbien.BoundText = DtcSubgrupoP.BoundText
    DtcPrecioU.BoundText = DtcSubgrupoP.BoundText
    DtcPrecioUV.BoundText = DtcSubgrupoP.BoundText
    Dtc_UniMed.BoundText = DtcSubgrupoP.BoundText
    DtcdesAnt.BoundText = DtcSubgrupoP.BoundText
    DtcCodAnt.BoundText = DtcSubgrupoP.BoundText
    DtcCodUniv.BoundText = DtcSubgrupoP.BoundText
    DtcPrecioC.BoundText = DtcSubgrupoP.BoundText
    Dtccodbien.BoundText = DtcSubgrupoP.BoundText
    DtcCodGrupoP.BoundText = DtcSubgrupoP.BoundText
End Sub

Private Sub dtcUnidad_Click(Area As Integer)
    DtcUnidadDes.BoundText = DtCUnidad.BoundText
End Sub

Private Sub DtcUnidad_LostFocus()

'    Set rstrc_personalSoli = New ADODB.Recordset
'    If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'    rstrc_personalSoli.Open "select * from unidad_responsable WHERE codigo_unidad='" & DtCUnidad.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    Set adopuestosol.Recordset = rstrc_personalSoli
'    adopuestosol.Refresh
    lblUni_codigo.Caption = DtCUnidad.Text
    'BEBECITA
    Set rstpoaAux = New ADODB.Recordset
    If rstpoaAux.State = 1 Then rstpoaAux.Close
    'rstpoaAux.Open "select * from fc_Relacionador_poa_ppto where (codigo_unidad = '" & DtCUnidad.Text & "') and (nivel  = 5)  ", db, adOpenKeyset, adLockReadOnly
    rstpoaAux.Open "select * from fc_Relacionador_poa_ppto where (codigo_unidad = '" & DtCUnidad.Text & "') and (nivel  = 5) and (Nuevo = 'S') ", db, adOpenKeyset, adLockReadOnly
    'rstpoa.Open "select par_codigo,* from fc_Relacionador_poa_ppto where (codigo_unidad = '" & adosolicitud.Recordset("codigo_unidad") & "') and (nivel  = 5) ORDER BY codigo_poa", db, adOpenKeyset, adLockReadOnly
'    If GlSistema = "Z" Then
'        rstpoaAux.Open "select * from fc_Relacionador_poa_ppto where (codigo_unidad = '" & DtCUnidad.Text & "') and (nivel  = 5)  ", db, adOpenKeyset, adLockReadOnly
'    Else
'        rstpoaAux.Open "select * from fc_Relacionador_poa_ppto where (codigo_unidad = '" & DtCUnidad.Text & "') and (nivel  = 5) and (ARCHIVO = '" & GlSistema & "') ", db, adOpenKeyset, adLockReadOnly
'    End If
    If rstpoaAux.RecordCount > 0 Then
        'If rstpoaAux!codigo_poa = "" Then
            DtcPOA.Text = rstpoaAux!codigo_poa
            DtcPOADes.Text = rstpoaAux!DESCRIPCION_POA
        Else
            DtcPOA.Text = "1.1.1.2.1"
            DtcPOADes.Text = "INSUMOS Y MATERIALES ODONTOLOGICOS"
        'End If
    End If
    Set AdoPOA.Recordset = rstpoaAux
    AdoPOA.Refresh
End Sub

Private Sub DtcUnidadDes_Click(Area As Integer)
    DtCUnidad.BoundText = DtcUnidadDes.BoundText
End Sub

Private Sub DTPfechasol_Change()
    txtGes_gestion = CStr(Year(DTPfechasol.Value))
End Sub

Private Sub Form_Load()
'jqa JUN/2008
   GlNombFor = "F04"
   Label7.Caption = GlUsuario
    
   Set rstFc_unidad_ejecutora = New ADODB.Recordset
   If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
   rstFc_unidad_ejecutora.Open "select * from Fc_unidad_ejecutora WHERE UNI_ACTIVO='S' ", db, adOpenKeyset, adLockReadOnly
   Set AdoUnidad.Recordset = rstFc_unidad_ejecutora
   AdoUnidad.Refresh
    
'   Set rstrc_personalSoli = New ADODB.Recordset
'   If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'   rstrc_personalSoli.Open "select * from unidad_responsable WHERE status='S' ORDER BY PATERNO ", db, adOpenKeyset, adLockReadOnly
'   Set adopuestosol.Recordset = rstrc_personalSoli
'   adopuestosol.Refresh
   
   Set rstrc_personalSoli = New ADODB.Recordset
   If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
   rstrc_personalSoli.Open "select * from fc_beneficiario where tipo_beneficiario=1 ORDER BY denominacion_beneficiario", db, adOpenKeyset, adLockReadOnly
   Set adopuestosol.Recordset = rstrc_personalSoli
   adopuestosol.Refresh
    
    
'   If (Not adosolicitud.Recordset.BOF) And (Not adosolicitud.Recordset.EOF) Then
'      dtccisol.Text = IIf(IsNull(adosolicitud.Recordset("ci")) = True, "-", adosolicitud.Recordset("ci"))
'        'Dtcpaternosol.Text = dtccisol.BoundText
'   End If
    
   Set rs_montador = New ADODB.Recordset
   If rs_montador.State = 1 Then rs_montador.Close
   rs_montador.Open "select * from Al_Montador order by descripcion ", db, adOpenKeyset, adLockReadOnly
   Set AdoMontador.Recordset = rs_montador
   AdoMontador.Refresh
  
   Set rsgrupo = New ADODB.Recordset
   If rsgrupo.State = 1 Then rsgrupo.Close
   rsgrupo.Open "select * from ALCLGrupo order by DescGrupo ", db, adOpenKeyset, adLockReadOnly
   Set AdoGRUPO.Recordset = rsgrupo
   AdoGRUPO.Refresh

    Set rstpoa = New ADODB.Recordset
    If rstpoa.State = 1 Then rstpoa.Close
    'rstpoa.Open "select * from fc_Relacionador_poa_ppto where (codigo_unidad = '" & DtcUnidad.Text & "') and (nivel  = 5) ORDER BY codigo_poa", db, adOpenKeyset, adLockReadOnly
    rstpoa.Open "select * from fc_Relacionador_poa_ppto where (nivel  = 5) ORDER BY codigo_poa", db, adOpenKeyset, adLockReadOnly
    Set AdoPOA.Recordset = rstpoa
    AdoPOA.Refresh

    'modi alb
    'FrmApertura.Visible = False
    ''''Lbltipo_bien_Cta_doc.Caption = ""
    Set rscc_parametros = New ADODB.Recordset
    If rscc_parametros.State = 1 Then rscc_parametros.Close
    rscc_parametros.Open " select * from cc_parametros where valor2 = 'F1A' order by valor1 ", db, adOpenKeyset, adLockReadOnly
    Set Adocc_parametros.Recordset = rscc_parametros
    Adocc_parametros.Refresh
    '
    Call proveedor

    Set rs_TipoSol = New ADODB.Recordset
    If rs_TipoSol.State = 1 Then rs_TipoSol.Close
    rs_TipoSol.Open "select * from ac_tipo_solicitud ", db, adOpenKeyset, adLockReadOnly
    Set AdoTipoSol.Recordset = rs_TipoSol
    AdoTipoSol.Refresh
    
   parametro = "ges_gestion" + " <> " + "'2012'"
   parametro2 = "cod_marca" + " <> " + "'0'"
   'parametro = "ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
   Call OptFilGral1_Click
   'Call ABRE_SOL_LISTA
   FrmEditaDet.Enabled = False
   swnuevo = 0
   swgrabar = 0
   sstab1.Tab = 0
   sstab1.TabEnabled(0) = True
   sstab1.TabEnabled(1) = False
	Call SeguridadSet(Me)
End Sub

Private Sub proveedor()
    Set rstrc_personalCargo = New ADODB.Recordset
    If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
    rstrc_personalCargo.Open "select * from fc_beneficiario where tipo_beneficiario=2 or tipo_beneficiario=22 ORDER BY denominacion_beneficiario", db, adOpenKeyset, adLockReadOnly
    Set adopuestobe.Recordset = rstrc_personalCargo
    adopuestobe.Refresh
End Sub

Private Sub ABRE_SOL_LISTA()
  
   Set rstao_solicitud_lista = New ADODB.Recordset
   If rstao_solicitud_lista.State = 1 Then rstao_solicitud_lista.Close
   rstao_solicitud_lista.Open "select * from ao_solicitud_lista where " & parametro & " order by CodGrupo, COD_MONTADOR, DescDetalle", db, adOpenKeyset, adLockOptimistic
   'rstao_solicitud_lista.Open "select * from ao_solicitud_lista where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' order by CodGrupo, COD_MONTADOR, profesion", db, adOpenKeyset, adLockOptimistic
   Set AdoAo_solicitud_Lista.Recordset = rstao_solicitud_lista
   Set DtGLista.DataSource = AdoAo_solicitud_Lista.Recordset
   AdoAo_solicitud_Lista.Refresh
   If AdoAo_solicitud_Lista.Recordset.RecordCount > 0 Then
        SOLISTA = "A"
        Set rstacumdet = New ADODB.Recordset
        If rstacumdet.State = 1 Then rstacumdet.Close
        rstacumdet.Open "select sum(precio_compra) as precuni, sum(precio_venta) as precSln, sum(total_compra) as precTot, sum(total_venta) as precTotV, sum(cantidad) as cantTot from ao_solicitud_lista where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' ", db, adOpenKeyset, adLockReadOnly
        'adoao_solicitud_lista.Caption = "TOTALES-->   Cantidad= " & CStr(rstacumdet!cantTot) & "    Precio Actual= " & CStr(rstacumdet!precSln) & "    Total Actual= " & CStr(rstacumdet!precTotV) & "    Precio c/Dscto= " & CStr(rstacumdet!precuni) & "    Total c/Dscto= " & CStr(rstacumdet!precTot) & " "
        'LblTotales.Caption = "TOTALES-->     |   " & CStr(rstacumdet!cantTot) & "       |   " & CStr(rstacumdet!precSln) & "       |   " & CStr(rstacumdet!precTotV) & "       |   " & CStr(rstacumdet!precuni) & "       |   " & CStr(rstacumdet!precTot) & "    |  "
        
        LblTotCant.Caption = "  " & CStr(rstacumdet!cantTot) & ""
        LblTotPrecAc.Caption = "  " & CStr(rstacumdet!precuni) & ""
        LblTotAct.Caption = "  " & CStr(rstacumdet!precTot) & ""
        LblTotPrecDsto.Caption = "  " & CStr(rstacumdet!precSln) & ""
        LblTotDscto.Caption = "  " & CStr(rstacumdet!precTotV) & ""
        
        If rstacumdet.State = 1 Then rstacumdet.Close
        DtGLista.Caption = "PRODUCTOS DE LA SOLICITUD Nro. " + Str(adosolicitud.Recordset("codigo_solicitud"))
        DtGLista.Visible = True
        Frame1.Caption = "ADICION/MODIFICACION DE PRODUCTOS DE LA SOLICITUD Nro. " + Str(adosolicitud.Recordset("codigo_solicitud"))
        AdoAo_solicitud_Lista.Visible = True
        Set DtGLista.DataSource = AdoAo_solicitud_Lista.Recordset
        'adosolicitud.Recordset!Lista_adjunta = "S"
        db.Execute " UPDATE AO_SOLICITUD SET Lista_adjunta='S' " & _
            "WHERE (ao_Solicitud.Ges_Gestion) = '" & adosolicitud.Recordset!ges_gestion & "' and " & _
            "(ao_Solicitud.codigo_unidad) = '" & adosolicitud.Recordset!codigo_unidad & "' and " & _
            "(ao_Solicitud.codigo_solicitud) =  " & adosolicitud.Recordset!codigo_solicitud & ""
   Else
        SOLISTA = "B"
        DtGLista.Caption = ""
        DtGLista.Visible = False
        'If adoao_solicitud_lista.Recordset.RecordCount > 0 Then
        AdoAo_solicitud_Lista.Visible = False
        'adosolicitud.Recordset!Lista_adjunta = "N"
        db.Execute " UPDATE AO_SOLICITUD SET Lista_adjunta='N' " & _
            "WHERE (ao_Solicitud.Ges_Gestion) = '" & adosolicitud.Recordset!ges_gestion & "' and " & _
            "(ao_Solicitud.codigo_unidad) = '" & adosolicitud.Recordset!codigo_unidad & "' and " & _
            "(ao_Solicitud.codigo_solicitud) =  " & adosolicitud.Recordset!codigo_solicitud & ""
        'End If
        LblTotCant.Caption = " "
        LblTotPrecAc.Caption = " "
        LblTotAct.Caption = " "
        LblTotPrecDsto.Caption = " "
        LblTotDscto.Caption = " "
   End If
   
   Set rsgrupo2 = New ADODB.Recordset
   If rsgrupo2.State = 1 Then rsgrupo2.Close
   rsgrupo2.Open "select * from ALCLGrupo where codigo_unidad = '" & DtCUnidad.Text & "' ", db, adOpenKeyset, adLockReadOnly
      
   Set rs_Bienes = New ADODB.Recordset
   If rs_Bienes.State = 1 Then rs_Bienes.Close
   If rsgrupo2.RecordCount > 0 Then
        rs_Bienes.Open "select * from AlClDetalle where cod_univ = '" & DtCUnidad & "' order by DescDetalle ", db, adOpenKeyset, adLockReadOnly
   Else
        rs_Bienes.Open "select * from AlClDetalle order by DescDetalle ", db, adOpenKeyset, adLockReadOnly
   End If
   Set ado_bienes.Recordset = rs_Bienes
   ado_bienes.Refresh
    
End Sub

Private Sub grabar()
'JQA JUN/2008
    db.BeginTrans
    If swgrabar = 1 Then
       'rstdestino.Open "select * from Ao_solicitud where formulario = '0'", db, adOpenDynamic, adLockOptimistic
       Dim correlsolic As Integer
       Dim rstao_solicitud_correl As New ADODB.Recordset
       Set rstao_solicitud_correl = New ADODB.Recordset
       If rstao_solicitud_correl.State = 1 Then rstao_solicitud_correl.Close
       rstao_solicitud_correl.Open "select * from fc_unidad_ejecutora where codigo_unidad = '" & Trim(DtCUnidad.Text) & "' ", db, adOpenDynamic, adLockOptimistic
       'rstao_solicitud_correl.Open "select * from ao_solicitud_correl where codigo_unidad = '" & Trim(DtcUnidad.Text) & "' ", db, adOpenDynamic, adLockOptimistic
        'rstao_solicitud_correl.Find "formulario = 'F03'", , adSearchForward
       If rstao_solicitud_correl.RecordCount > 0 Then
          If Not (rstao_solicitud_correl.BOF) Then rstao_solicitud_correl.MoveFirst
          rstao_solicitud_correl("correl_solicitud") = rstao_solicitud_correl("correl_solicitud") + 1
          correlsolic = rstao_solicitud_correl("correl_solicitud")
          rstao_solicitud_correl.Update
       Else
          rstao_solicitud_correl.AddNew
          rstao_solicitud_correl("codigo_unidad") = Trim(lblUni_codigo.Caption)
          rstao_solicitud_correl("correl_solicitud") = 1
          correlsolic = rstao_solicitud_correl("correl_solicitud")
          rstao_solicitud_correl.Update
       End If
       If rstao_solicitud_correl.State = 1 Then rstao_solicitud_correl.Close
       adosolicitud.Recordset("ges_gestion") = GlGestion        'CStr(Year(DTPfechasol.Value))
       adosolicitud.Recordset("codigo_solicitud") = correlsolic
       adosolicitud.Recordset("codigo_unidad") = DtCUnidad.Text
       adosolicitud.Recordset("Lista_adjunta") = "N"
    End If
    If DTPfechasol.Value = "01/01/1900" Then
        adosolicitud.Recordset("fecha_solicitud") = Format(Date, "dd/mm/yyyy")
    Else
        adosolicitud.Recordset("fecha_solicitud") = DTPfechasol.Value
    End If
        adosolicitud.Recordset("CI") = dtccisol.Text
        adosolicitud.Recordset("CI_aprueba") = Dtccibe.Text
        adosolicitud.Recordset("codigo_poa") = DtcPOA.Text
        adosolicitud.Recordset("codigo_bien") = Dtccodbien.Text
        adosolicitud.Recordset("caracteristicas") = Txtcaracteristicas.Text
        adosolicitud.Recordset("justificacion_solicitud") = Txtcaracteristicas.Text     'txtjustifica.Text
        adosolicitud.Recordset("observaciones") = txtobservaciones.Text
        adosolicitud.Recordset("tr_adjuntos") = IIf(IsNull(txtterref.Text), "N", txtterref.Text)
        adosolicitud.Recordset("TipoF1") = DtcTipo.Text     'PD=Pedido Directo al Proveedor CN=Cotizaciones     'jqa jun/2008 Cargo de Cuenta "CC"  'DtCvalor1.BoundText
      
        adosolicitud.Recordset("subcta2") = "02"     'JQA JUN/2008 Cargo de Cuenta Otros
        If Val(Txt_porcentaje.Text) > 0 Then
            adosolicitud.Recordset("por_tiempo") = Val(Txt_porcentaje.Text)     '/ 100
        Else
            adosolicitud.Recordset("por_tiempo") = 0
        End If
        adosolicitud.Recordset("formulario") = "F04"
        adosolicitud.Recordset("tipo_bien_Cta_doc") = "A"
        If adosolicitud.Recordset("Lista_adjunta") = "S" Then
            adosolicitud.Recordset("Lista_adjunta") = "S"
        Else
            adosolicitud.Recordset("Lista_adjunta") = "N"
        End If
        adosolicitud.Recordset("codigo_bien") = Dtccodbien.Text
        'adosolicitud.Recordset("nro_pagos") = 1     'IIf(IsNull(TxtCantPedi), 1, TxtCantPedi)
        adosolicitud.Recordset("usr_usuario") = GlUsuario '"xxx"
        adosolicitud.Recordset("fecha_registro") = Format(Date, "dd/mm/yyyy")
        adosolicitud.Recordset("hora_registro") = Format(Time, "hh:mm:ss") '"16:00:00"
        adosolicitud.Recordset("usuario_aprueba") = ""
        adosolicitud.Recordset("hora_aprueba") = ""
'        adosolicitud.Recordset("AUnidad") = "-"
'        adosolicitud.Recordset("APlanilla") = 0
'        adosolicitud.Recordset("Planilla_depto") = "-"
'        adosolicitud.Recordset("Bco_codigo") = "-"
'        adosolicitud.Recordset("Ges_Gestion_ant") = "-"
'        adosolicitud.Recordset("APlanilla_existe") = "N"
        adosolicitud.Recordset("estado_enviado") = "N"
        adosolicitud.Recordset("estado_aprobado") = "N"
      adosolicitud.Recordset.Update

    db.CommitTrans
    If adosolicitud.Recordset.RecordCount > 0 Then
       marca1 = adosolicitud.Recordset.Bookmark
       parametro = " ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
       Call OptFilGral1_Click
       If swgrabar = 1 Then
           adosolicitud.Recordset.MoveLast
       Else
           adosolicitud.Recordset.Move marca1 - 1
       End If
    End If
'JQA JUN 2008
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
    Else
      KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub Text1_LostFocus()
    If CDbl(Text1.Text) > 0 Then
        Txt_porcentaje.Text = CDbl(Text1.Text) / 100
    Else
        Txt_porcentaje.Text = 0
    End If
End Sub

Private Sub txtjustifica_KeyPress(KeyAscii As Integer)
   'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    'solo numeros a numero
'      If KeyAscii < 58 And KeyAscii > 47 Or KeyAscii = 8 Then
'
'      Else
'        KeyAscii = Asc(UCase(Chr(0)))
'      End If

End Sub

Private Sub TxtMonto_bolivianos_contra_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
    Else
      KeyAscii = Asc(UCase(Chr(0)))
    End If

End Sub

'Private Sub TxtMonto_bolivianos_contra_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If (Len(Trim(TxtMonto_bolivianos_contra.Text)) > 0) Then
'       Txtmonto_dolares_contra.Text = IIf(TxtMonto_bolivianos_contra.Text > 0, TxtMonto_bolivianos_contra.Text / TxtTipo_cambio, 0)
'    Else
'       Txtmonto_dolares_contra.Text = 0
'    End If
'  End If
'
'End Sub
'
'Private Sub TxtMonto_bolivianos_KeyPress(KeyAscii As Integer)
''solo numeros y , .
'    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'
'    Else
'      KeyAscii = Asc(UCase(Chr(0)))
'    End If
'End Sub
'
'Private Sub TxtMonto_bolivianos_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If (Len(Trim(TxtMonto_bolivianos.Text)) > 0) Then
'       Txtmonto_dolares.Text = IIf(TxtMonto_bolivianos.Text > 0, TxtMonto_bolivianos.Text / TxtTipo_cambio, 0)
'    Else
'       Txtmonto_dolares.Text = 0
'    End If
'  End If
'
'End Sub
'
'Private Sub Txtmonto_dolares_contra_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'    Else
'      KeyAscii = Asc(UCase(Chr(0)))
'    End If
'
'End Sub
'
'Private Sub Txtmonto_dolares_contra_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If Len(Trim(Txtmonto_dolares_contra.Text)) > 0 Then
'      TxtMonto_bolivianos_contra.Text = IIf(Txtmonto_dolares_contra.Text > 0, Txtmonto_dolares_contra * TxtTipo_cambio, 0)
'    Else
'      TxtMonto_bolivianos_contra.Text = 0
'    End If
'  End If
'
'End Sub
'
'Private Sub Txtmonto_dolares_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'    Else
'      KeyAscii = Asc(UCase(Chr(0)))
'    End If
'
'End Sub
'
'Private Sub Txtmonto_dolares_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If Len(Trim(Txtmonto_dolares.Text)) > 0 Then
'      TxtMonto_bolivianos.Text = IIf(Txtmonto_dolares.Text > 0, Txtmonto_dolares * TxtTipo_cambio, 0)
'    Else
'      TxtMonto_bolivianos.Text = 0
'    End If
'  End If
'
'End Sub

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

Private Sub TxtTipo_cambio_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
    Else
      KeyAscii = Asc(UCase(Chr(0)))
    End If

End Sub
Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros aprobados)
  Set rstAo_solicitud = New ADODB.Recordset
  'queryinicial = "select * from Ao_solicitud where formulario = 'F01' and estatus <> 'A' and estatus <> 'S' AND usr_usuario = '" & GlUsuario & "' "
  queryinicial = "select * from Ao_solicitud where formulario = 'F04' and estado_enviado = 'N'"
  If rstAo_solicitud.State = 1 Then rstAo_solicitud.Close
  rstAo_solicitud.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  'rstAo_solicitud.Open queryinicial, db, adOpenStatic, adLockReadOnly
  rstAo_solicitud.Requery
  If swgrabar = 1 Then
    'rstAo_solicitud.Open queryinicial & " order by fecha_registro , hora_registro", db, adOpenKeyset, adLockOptimistic
    rstAo_solicitud.Sort = "fecha_registro , hora_registro"
  Else
    'rstAo_solicitud.Open queryinicial & " order by codigo_unidad , codigo_solicitud", db, adOpenKeyset, adLockOptimistic
    rstAo_solicitud.Sort = "codigo_unidad , codigo_solicitud"
  End If
  Set adosolicitud.Recordset = rstAo_solicitud
  Set DtGLista.DataSource = rstAo_solicitud
  If rstAo_solicitud.RecordCount > 0 Then
    'Frame10.Enabled = True
    Frame10.Visible = True
    BtnImprimir.Enabled = True
    BtnBuscar.Enabled = True
    ' 06/03/2012 ADALID
    Call ABRE_SOL_LISTA
  Else
    'Frame10.Enabled = False
    Frame10.Visible = False
    BtnImprimir.Enabled = False
    BtnBuscar.Enabled = False
  End If
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
  Set rstAo_solicitud = New ADODB.Recordset
  'queryinicial = "select * from Ao_solicitud where formulario = 'F01' AND usr_usuario = '" & GlUsuario & "' "
  queryinicial = "select * from Ao_solicitud where formulario = 'F04' "
  If rstAo_solicitud.State = 1 Then rstAo_solicitud.Close
  'rstAo_solicitud.Open queryinicial & " order by codigo_unidad , codigo_solicitud", db, adOpenKeyset, adLockOptimistic
  rstAo_solicitud.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  rstAo_solicitud.Requery
  rstAo_solicitud.Sort = "codigo_unidad , codigo_solicitud"
  Set adosolicitud.Recordset = rstAo_solicitud
  Set DtGLista.DataSource = rstAo_solicitud
  If rstAo_solicitud.RecordCount > 0 Then
    Frame10.Enabled = True
    Frame10.Visible = True
    BtnImprimir.Enabled = True
    BtnBuscar.Enabled = True
  Else
    Frame10.Enabled = False
    Frame10.Visible = False
    BtnImprimir.Enabled = False
    BtnBuscar.Enabled = False
  End If
End Sub

'Private Sub fbuscaunidad()
'  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
'  'rstFc_unidad_ejecutora.Open "select * from Fc_unidad_ejecutora where uni_codigo = '" & Trim(adopuestosol.Recordset("codigo_unidad")) & "'", db, adOpenKeyset, adLockReadOnly
'  rstFc_unidad_ejecutora.Open "select * from Fc_unidad_ejecutora where uni_codigo = '" & Trim(adopuestosol.Recordset("codigo_unidad")) & "'", db, adOpenKeyset, adLockReadOnly
'  If rstFc_unidad_ejecutora.RecordCount > 0 Then
'    LblUni_descripcion_larga.Caption = rstFc_unidad_ejecutora("Uni_descripcion_larga")
'  Else
'    LblUni_descripcion_larga.Caption = ""
'  End If
'  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
'End Sub

Private Sub cerea()
  txtnrosol = ""
''  Optpasvia1.Value = False
'  Optpasvia2.Value = False
'  dtccodpoa.Text = ""
'  dtcdespoa.Text = dtccodpoa.BoundText
  dtccisol.Text = ""
  'Dtcpaternosol.Text = dtccisol.BoundText
  Dtcpaternosol.Text = ""

  Dtccibe.Text = ""
  Dtcpaternobe.Text = ""
'  Dtcpaternobe.Text = Dtccibe.BoundText
'  DtCcodigo_beneficiario = ""
'  DtCdenominacion_beneficiario = DtCcodigo_beneficiario.BoundText
  txtjustifica.Text = ""
  txtterref.Text = ""
'  TxtDurac_tiempo.Text = ""
'  DtCDenominacion_moneda.Text = ""
'  TxtTipo_cambio.Text = GlTipoCambioOficial
'  TxtMonto_bolivianos.Text = 0
'  Txtmonto_dolares.Text = 0
'  DtCOrg_descripcion.Text = ""
'  TxtMonto_bolivianos_contra.Text = 0
'  Txtmonto_dolares_contra.Text = 0
End Sub


Private Sub DtCvalor1_LostFocus()
  If DtCvalor1.BoundText = "CC" Then
    Label3.Visible = True
    cmbSubCta2.Visible = True
  Else
    Label3.Visible = False
    cmbSubCta2.Visible = False
    cmbSubCta2.Text = ""
  End If
End Sub

Private Sub APrueba2()
Dim rsAcum As New ADODB.Recordset
  Dim Acum As Double
  Dim Acumlimite As Double
  Set rstdestino = New ADODB.Recordset
  If rstdestino.State = 1 Then rstdestino.Close
  rstdestino.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adosolicitud.Recordset("ges_gestion") & "' and codigo_solicitud = " & adosolicitud.Recordset("codigo_solicitud") & " and codigo_unidad = '" & adosolicitud.Recordset("codigo_unidad") & "'", db, adOpenKeyset, adLockOptimistic
  If rstdestino.RecordCount < 1 Then
    MsgBox "No puede aprobar sin Detalle de Registro.", vbCritical + vbOKOnly, "Error al aprobar..."
    Exit Sub
  End If
  Dim swver_monto As Integer
  swver_monto = 1
  If Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "R" Or Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "C" Then
    '==== ini convenios ====
    Dim rstAo_solicitud_detalle_ant As New ADODB.Recordset
    Set rstAo_solicitud_detalle_ant = New ADODB.Recordset
    Dim swconv As Integer
    swconv = 1
    conv2 = " "
    conv1 = " "
    rstao_solicitud_detalle.MoveFirst
    While Not rstao_solicitud_detalle.EOF
      Call fBuscaConvenio(rstao_solicitud_detalle!codigo_poa)
      If rstAo_solicitud_detalle_ant.State = 1 Then rstAo_solicitud_detalle_ant.Close
      rstAo_solicitud_detalle_ant.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adosolicitud.Recordset!ges_gestion_ant & "' and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad_ant & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud_ant, db, adOpenKeyset, adLockReadOnly
      If rstAo_solicitud_detalle_ant.RecordCount > 0 Then
        While Not rstAo_solicitud_detalle_ant.EOF
          conv1 = rstAo_solicitud_detalle_ant!codigo_convenio
          If rstAo_solicitud_detalle_ant!codigo_convenio <> conv2 Then
            MsgBox "No puede aprobar una RENDICION o un CIERRE " & vbCrLf & "con un convenio diferente al de la APERTURA." & vbCrLf & vbCrLf & _
                   "       Convenio Apertura  : " & rstAo_solicitud_detalle_ant!codigo_convenio & vbCrLf & _
                   "       Convenio Solicitud  : " & conv2 & " (poa: " & rstao_solicitud_detalle!codigo_poa & ") " & vbCrLf & _
            vbCrLf & "          Por favor corrija los codigos POA", vbCritical + vbOKOnly, "Error al aprobar..."
            swconv = 0
          End If
          rstAo_solicitud_detalle_ant.MoveNext
        Wend
      Else
        MsgBox "Error en el convenio de la APERTURA.", vbCritical + vbOKOnly, "Error al aprobar..."
        Exit Sub
      End If
      rstao_solicitud_detalle.MoveNext
    Wend
    '==== fin convenios ====
    If swconv = 0 Then Exit Sub
    '==== ini CTA ====
'    Dim rstfc_convenio As New ADODB.Recordset
'    Set rstfc_convenio = New ADODB.Recordset
'    If rstfc_convenio.State = 1 Then rstfc_convenio.Close
'    rstfc_convenio.Open "select * from fc_convenioS where codigo_convenio = '" & Conv1 & "' ", db, adOpenKeyset, adLockReadOnly
'    If rstfc_convenio.RecordCount > 0 Then
'      cta1 = rstfc_convenio!cta_codigo
'    Else
'      cta1 = " "
'    End If
'    If rstfc_convenio.State = 1 Then rstfc_convenio.Close
'    If cta1 <> Me.DtCBco_codigo.Text Then
'      MsgBox "La cuenta bancaria de la solicitud " & vbCrLf & "debe pertenecer al Convenio de Apertura" & vbCrLf & vbCrLf & _
'             "     Cuenta de Apertura : " & cta1 & vbCrLf & _
'             "     Cuenta de Solicitud : " & Me.DtCBco_codigo.Text, vbCritical + vbOKOnly, "Error en las Cuentas Bancarias..."
'      Exit Sub
'    End If
    '==== fin CTA ====
    
    '==== INI MONTOS ====
    Set rsAcum = New ADODB.Recordset
    If rsAcum.State = 1 Then rsAcum.Close
    rsAcum.Open "select sum (monto_bolivianos) + sum (monto_bolivianos_CONTRA)as acum from ao_solicitud_detalle where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud, db, adOpenKeyset, adLockReadOnly
    If rsAcum.RecordCount > 0 Then
      Acum = IIf(IsNull(rsAcum!Acum), 0, rsAcum!Acum)
    End If
    If rsAcum.State = 1 Then rsAcum.Close
    Set rsAcum = New ADODB.Recordset
    If rsAcum.State = 1 Then rsAcum.Close
    rsAcum.Open "select sum (monto_bolivianos) + sum (monto_bolivianos_CONTRA)as acum from ao_solicitud_detalle where ges_gestion = '" & adosolicitud.Recordset!ges_gestion_ant & "' and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad_ant & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud_ant, db, adOpenKeyset, adLockReadOnly
    If rsAcum.RecordCount > 0 Then
      Acumlimite = IIf(IsNull(rsAcum!Acum), 0, rsAcum!Acum)
    End If
    If rsAcum.State = 1 Then rsAcum.Close
    If Acum + Me.adosolicitud.Recordset!Nro_pagos > Acumlimite Then
      MsgBox "No puede aprobar una RENDICION o un CIERRE" & vbCrLf & "con monto mayor al de la APERTURA." & vbCrLf & _
             "     Monto de Apertura : " & Acumlimite & vbCrLf & _
             "     Monto Solicitud      : " & Acum + Me.adosolicitud.Recordset!Nro_pagos & vbCrLf & _
      vbCrLf & vbCrLf & "          Por favor corrija los MONTOS.", vbCritical + vbOKOnly, "Error al aprobar..."
      Exit Sub
    End If
    If Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "C" Then
      If (Acum + Me.adosolicitud.Recordset!Nro_pagos) < Acumlimite Then
        MsgBox "No puede aprobar un CIERRE con monto menor al de APERTURA." & vbCrLf & _
        "     Monto APERTURA : " & Acumlimite & vbCrLf & "     Monto CIERRE        : " & Acum + Me.adosolicitud.Recordset!Nro_pagos & vbCrLf & vbCrLf & _
        "          Por favor corrija los MONTOS.", vbCritical + vbOKOnly, "Error al aprobar..."
        Exit Sub
      End If
    End If
  End If
'  swver_monto = verifica_montos(Me.adosolicitud.Recordset!codigo_unidad, Me.adosolicitud.Recordset!codigo_solicitud)
  If swver_monto = 0 Then
    MsgBox "El registro tiene problemas en los montos, Por favor Verifique e intente aprobarlo luego. Gracias", vbCritical + vbOKOnly, "Error en los montos"
    Exit Sub
  End If
'==== MONTOS ====
  sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
  If sino = vbYes Then
    db.BeginTrans
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    rstdestino.Open "select * from Ao_solicitud where ges_gestion = '" & adosolicitud.Recordset("ges_gestion") & "' and formulario = '" & adosolicitud.Recordset("formulario") & "' and codigo_solicitud = " & adosolicitud.Recordset("codigo_solicitud") & " and codigo_unidad = '" & adosolicitud.Recordset("codigo_unidad") & "'", db, adOpenDynamic, adLockOptimistic
    If Not rstdestino.BOF Then rstdestino.MoveFirst
    If Not rstdestino.BOF And Not rstdestino.EOF Then
      rstdestino("aprobado") = 1
      rstdestino("estado_aprobacion") = "S"
      rstdestino.Update
      If rstdestino!tipo_bien_Cta_doc = "C" Then
        Set rsAcum = New ADODB.Recordset
'        rsAcum.CancelUpdate
        If rsAcum.State = 1 Then rsAcum.Close
        rsAcum.Open "select * from ao_solicitud where ges_gestion = '" & adosolicitud.Recordset!ges_gestion_ant & "' and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad_ant & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud_ant, db, adOpenKeyset, adLockOptimistic
        If rsAcum.RecordCount > 0 Then
          rsAcum!codigo_unidad_ant = "X"
          rsAcum.Update
        End If
        If rsAcum.State = 1 Then rsAcum.Close
      End If
    End If
    If rstdestino.State = 1 Then rstdestino.Close
    db.CommitTrans
    marca1 = adosolicitud.Recordset.Bookmark
    Set adosolicitud.Recordset = rstAo_solicitud
    adosolicitud.Refresh
    adosolicitud.Recordset.Move marca1 - 1
'    rstAo_solicitud.Requery
  End If
End Sub

Private Sub fBuscaConvenio(Poa)
  Dim rst_fc_relacionador_poa_ppto As New ADODB.Recordset
  Set rst_fc_relacionador_poa_ppto = New ADODB.Recordset
  If rst_fc_relacionador_poa_ppto.State = 1 Then rst_fc_relacionador_poa_ppto.Close
  rst_fc_relacionador_poa_ppto.Open "select * from fc_relacionador_poa_ppto where codigo_poa = '" & Poa & "'", db, adOpenKeyset, adLockReadOnly
  If rst_fc_relacionador_poa_ppto.RecordCount > 0 Then
    conv2 = rst_fc_relacionador_poa_ppto!codigo_convenio
  Else
    conv2 = "Err"
  End If
  If rst_fc_relacionador_poa_ppto.State = 1 Then rst_fc_relacionador_poa_ppto.Close
End Sub

Private Function fbusmuni(cod)
  Dim rstfc_unidad_educativa As New ADODB.Recordset
  Set rstfc_unidad_educativa = New ADODB.Recordset
  If rstfc_unidad_educativa.State = 1 Then rstfc_unidad_educativa.Close
  rstfc_unidad_educativa.Open "select * from fc_unidad_educativa where codigo = '" & cod & "'", db, adOpenKeyset, adLockReadOnly
  If rstfc_unidad_educativa.RecordCount > 0 Then
    fbusmuni = rstfc_unidad_educativa!denominacion
  Else
    fbusmuni = ""
  End If
  If rstfc_unidad_educativa.State = 1 Then rstfc_unidad_educativa.Close
End Function

'Public Sub val_presupF04(adoorigen, GlNombFor)
'  'If (GlNombFor <> "F01") And (GlNombFor <> "F06") Or (GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "R" Or Trim(adoorigen!tipo_bien_Cta_doc) = "C")) Then
'  If GlNombFor <> "F02" And GlNombFor <> "F06" And GlNombFor <> "F07" And GlNombFor <> "F08" Then
'    Set rstao_solicitud_detalle = New ADODB.Recordset
'    If rstao_solicitud_detalle.State = 1 Then rstao_solicitud_detalle.Close
'    rstao_solicitud_detalle.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockReadOnly
'    If rstao_solicitud_detalle.RecordCount > 0 Then
'      rectot = rstao_solicitud_detalle.RecordCount
'      Fte_contraparte1 = fBuscaFte(rstao_solicitud_detalle!org_codigo_contra)
'      Org_Contraparte1 = rstao_solicitud_detalle!org_codigo_contra
'      Dim v_EstPoa(50, 18)
'    End If
'    If Not (rstao_solicitud_detalle.BOF) Then rstao_solicitud_detalle.MoveFirst
'    For i = 1 To rstao_solicitud_detalle.RecordCount
'      Set rstfc_relacionador_poa_ppto = New ADODB.Recordset
'      If rstfc_relacionador_poa_ppto.State = 1 Then rstfc_relacionador_poa_ppto.Close
'      rstfc_relacionador_poa_ppto.Open "select * from fc_relacionador_poa_ppto where codigo_poa = '" & rstao_solicitud_detalle!codigo_poa & "'", db, adOpenKeyset, adLockReadOnly
'      If rstfc_relacionador_poa_ppto.RecordCount > 0 Then
'
'    'aqui se puede definir porcentaje
'        v_EstPoa(i, 1) = rstao_solicitud_detalle!codigo_poa
'        v_EstPoa(i, 2) = rstfc_relacionador_poa_ppto!da  'Dirección Administrativa JQA OCT-2009
'        v_EstPoa(i, 3) = rstfc_relacionador_poa_ppto!par_codigo 'Par_Codigo1
'        v_EstPoa(i, 4) = rstfc_relacionador_poa_ppto!fte_codigo 'fte_codigo1    JQA JUL-2005
'        v_EstPoa(i, 5) = rstfc_relacionador_poa_ppto!org_codigo 'Org_Codigo1
'        v_EstPoa(i, 6) = rstfc_relacionador_poa_ppto!pro_programa 'pro_Programa1
'        v_EstPoa(i, 7) = IIf(IsNull(rstfc_relacionador_poa_ppto!pro_subprograma), "00", rstfc_relacionador_poa_ppto!pro_subprograma) 'Pro_SubPrograma1   JQA OCT-2009
'        v_EstPoa(i, 8) = rstfc_relacionador_poa_ppto!pro_proyecto 'Pro_Proyecto1
'        v_EstPoa(i, 9) = rstfc_relacionador_poa_ppto!pro_actividad 'Pro_Actividad1
'        v_EstPoa(i, 10) = rstfc_relacionador_poa_ppto!codigo_unidad 'codigo_UNIDAD1
'        v_EstPoa(i, 11) = rstfc_relacionador_poa_ppto!codigo_Categoria  'codigo_categoria1    JQA OCT-2009
'        v_EstPoa(i, 12) = rstfc_relacionador_poa_ppto!codigo_convenio 'codigo_convenio1
'        por_fte_ext1 = rstfc_relacionador_poa_ppto!por_ext
'        por_fte_nal1 = rstfc_relacionador_poa_ppto!por_nal
'        If rstfc_relacionador_poa_ppto!por_ext = 100 Then           'CONTRAPARTE JQA JUL-2005
'            v_EstPoa(i, 13) = rstfc_relacionador_poa_ppto!fte_codigo 'fte_codigo2   JQA JUL-2005
'            v_EstPoa(i, 14) = rstfc_relacionador_poa_ppto!org_codigo 'Org_Codigo2   JQA JUL-2005
'            cat_nal = rstfc_relacionador_poa_ppto!codigo_Categoria  'codigo_categoria1    JQA JUL-2005
'            conv_nal = rstfc_relacionador_poa_ppto!codigo_convenio 'codigo_convenio1
'            tot_form = 1
'        Else
'            v_EstPoa(i, 13) = "10"              ' REVISAR !!!!!!!!!!!!!!!!!!!!!!!!!     JQA JUL-2005
'            v_EstPoa(i, 14) = "111"             ' REVISAR !!!!!!!!!!!!!!!!!!!!!!!!!     JQA JUL-2005
'            cat_nal = IIf(IsNull(rstfc_relacionador_poa_ppto!CODIGO_Categoria_nal), "S/C TGN", rstfc_relacionador_poa_ppto!CODIGO_Categoria_nal)  'codigo_categoria2    JQA JUL-2005
'            conv_nal = rstfc_relacionador_poa_ppto!codigo_convenio_nal 'codigo_convenio2
'            tot_form = 2
'        End If
'        v_EstPoa(i, 15) = IIf(IsNull(rstfc_relacionador_poa_ppto!Categoria), "-", rstfc_relacionador_poa_ppto!Categoria) 'categoria1 COMPLEMENTARIA   JQA OCT-2009
'        v_EstPoa(i, 16) = rstfc_relacionador_poa_ppto!uni_codigo 'uni_codigo1
'        v_EstPoa(i, 17) = rstao_solicitud_detalle!monto_bolivianos 'monto en Bs
'        v_EstPoa(i, 18) = rstao_solicitud_detalle!monto_dolares 'monto en SUS
'        'If rstao_solicitud_detalle!org_codigo_contra = "" Or rstao_solicitud_detalle!org_codigo_contra = "-" Then      'JQA JUL-2005
'        '  v_EstPoa(i, 13) = "10"
'        '  v_EstPoa(i, 14) = "111"
'        'Else
'        '  v_EstPoa(i, 13) = fBuscaFte(rstao_solicitud_detalle!org_codigo_contra)
'        '  v_EstPoa(i, 14) = rstao_solicitud_detalle!org_codigo_contra
'        'End If
'        If rstfc_relacionador_poa_ppto.State = 1 Then rstfc_relacionador_poa_ppto.Close
'        Dim rstfo_formulacion_gasto As New ADODB.Recordset
'        Set rstfo_formulacion_gasto = New ADODB.Recordset
'        If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
'        'rstfo_formulacion_gasto.Open "select * from fo_formulacion_gasto where pro_programa='" & pro_Programa1 & "' and pro_subprograma='" & Pro_SubPrograma1 & "' and pro_proyecto='" & Pro_Proyecto1 & "' and pro_actividad='" & Pro_Actividad1 & "' and par_codigo='" & Par_Codigo1 & "' and org_codigo= '" & Org_Codigo1 & "'", db, adOpenKeyset, adLockOptimistic
'        rstfo_formulacion_gasto.Open "select * from fo_formulacion_gasto where pro_programa='" & v_EstPoa(i, 6) & "' and pro_proyecto='" & v_EstPoa(i, 8) & "' and pro_actividad='" & v_EstPoa(i, 9) & "' and par_codigo='" & v_EstPoa(i, 3) & "' and org_codigo= '" & v_EstPoa(i, 5) & "'", db, adOpenKeyset, adLockOptimistic
'        If Not (rstfo_formulacion_gasto.EOF) Then
'          If (rstfo_formulacion_gasto!fgs_vigente - rstfo_formulacion_gasto!FGS_compromiso < rstao_solicitud_detalle!monto_bolivianos) Then  'adoorigen         'adoorigen.adosolicitud.Recordset!monto_dolares ) Then
'            'JQA 07/12/01
''            swSubir = "No existe Presup"
'            MsgBox "NO EXISTE Presupuesto para dar curso a la Solicitud, debe informar a Presupuestos ...", vbOKOnly, "ERROR"
''            swpresup = 0
''            Exit Sub
'            'JQA 07/12/01
'            swpresup = 1    'Borrar despues de habilitar JQA
'          Else
'            'JQA 07/12/01
'            'rstfo_formulacion_gasto!0  = rstfo_formulacion_gasto!fgs_precompromiso  + rstao_solicitud_detalle!monto_bolivianos
'            'rstfo_formulacion_gasto.Update
'            'JQA 07/12/01
'            swpresup = 1
'            swSubir = "SI correcto"
'          End If
'          If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
'            swpresup = 1
'          Else
'            'JQA 07/12/01
''          MsgBox "NO EXISTE Estructura presupuestaria...", vbOKOnly, "ERROR ..."
''          swSubir = "NO Error Estruc.Ppto"
''          swpresup = 0
''          Exit Sub
'            'JQA 07/12/01
'            swpresup = 1    'Borrar despues de habilitar JQA
'          End If
'      Else
'        MsgBox "NO Existe POA ... ", vbOKOnly, "ERROR ..."
'        swSubir = "No existe POA"
'        swpresup = 0
'        Exit Sub
'      End If
''          Else
''            swpresup = 1
''          End If
'      rstao_solicitud_detalle.MoveNext
'    Next
'
'    If swpresup = 1 Then
'      If GlNombFor <> "F02" And GlNombFor <> "F06" And GlNombFor <> "F07" And GlNombFor <> "F08" Then
'        If (rstao_solicitud_detalle.RecordCount > 0) And (Not rstao_solicitud_detalle.BOF) Then rstao_solicitud_detalle.MoveFirst
'        Set rstao_solicitud_recibido = New ADODB.Recordset
'        If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
'        rstao_solicitud_recibido.Open "SELECT * FROM ao_solicitud_recibido", db, adOpenKeyset, adLockOptimistic
'        db.BeginTrans
'        'For j = 1 To rstao_solicitud_detalle.RecordCount       ' REV. JQA OCT-2009
'        For j = 1 To tot_form       'tot_form=1 (100%) y tot_form>1 CON CONTRAPARTE
'          '-- 100%  'UN SOLO FINANCIADOR
'          v_por_fte(1, 1) = por_fte_ext1
'          v_por_fte(1, 2) = v_EstPoa(j, 4) 'fte_codigo1
'          v_por_fte(1, 3) = v_EstPoa(j, 5) 'Org_Codigo1
'          '__ con Contraparte
'          v_por_fte(2, 1) = por_fte_nal1
'          v_por_fte(2, 2) = v_EstPoa(j, 13) 'Fte_contraparte1
'          v_por_fte(2, 3) = v_EstPoa(j, 14) 'Org_Contraparte1
'          '__ Segunda Contraparte
'          v_por_fte(3, 1) = por_fte_nal1
'          v_por_fte(3, 2) = v_EstPoa(j, 13) 'Fte_contraparte1
'          v_por_fte(3, 3) = v_EstPoa(j, 14) 'Org_Contraparte1
'          Dim SwEsBase As Integer
'          Dim ValEsBase As Double
''          ValEsBase = v_por_fte(1, 1)
''          For I = 1 To tot_form
''            If v_por_fte(I, 1) > ValEsBase Then
''              SwEsBase = I
''              ValEsBase = v_por_fte(I, 1)
''            End If
''          Next
''          For i = 1 To rectot       'tot_form(ya no)
'          Set rstpagos = New ADODB.Recordset
'          If rstpagos.State = 1 Then rstpagos.Close
'            'If i = 1 Then              ' REV JQA OCT-2009
'            '  rstpagos.Open "select * from pagos_espera where codigo_pago = ''", db, adOpenKeyset, adLockOptimistic
'            'Else
'          rstpagos.Open "select * from pagos where codigo_pago = ''", db, adOpenKeyset, adLockOptimistic
'            'End If                 ' REV JQA OCT-2009
'          rstpagos.AddNew
'            '==== ini generación de correlativo ====
'          Set rscorrelativo = New ADODB.Recordset
'          If rscorrelativo.State = 1 Then rscorrelativo.Close
''            If i = 1 Then              ' REV JQA OCT-2009
''              'rscorrelativo.Open "select * from fc_correlativos_espera", db, adOpenKeyset, adLockOptimistic
''              '======== ini GENERA EL CODIGO DE COMPROBANTE ========
''                Set rscorrelativo = New ADODB.Recordset
''                rscorrelativo.CursorLocation = adUseClient
''                If rscorrelativo.State = 1 Then rscorrelativo.Close
''                rscorrelativo.Open "select * from fc_Correlativos_espera  where org_codigo = '" & v_por_fte(1, 3) & "' ", db, adOpenDynamic, adLockOptimistic
''                If rscorrelativo.RecordCount > 0 Then
''                  codigo_pago1 = Val(rscorrelativo!correlativo)
''                  codigo_pago1 = codigo_pago1 + 1
''                  rscorrelativo!correlativo = Trim(Str(codigo_pago1))
''                  rscorrelativo.Update
''                End If
''                If rscorrelativo.State = 1 Then rscorrelativo.Close
''                '======== fin TERMINA GENERACION DE COMPROBANTE ========
''            Else               ' REV JQA OCT-2009
'              'rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
'              '======== ini GENERA EL CODIGO DE COMPROBANTE ========
'          Set rscorrelativo = New ADODB.Recordset
'          rscorrelativo.CursorLocation = adUseClient
'          If rscorrelativo.State = 1 Then rscorrelativo.Close
'          'rscorrelativo.Open "select * from fc_Correlativos  where org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenDynamic, adLockOptimistic
'          rscorrelativo.Open "select * from fc_organismo_financiamiento where org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenDynamic, adLockOptimistic
'          If rscorrelativo.RecordCount > 0 Then
'                  codigo_pago1 = Val(rscorrelativo!correlativo)
'                  codigo_pago1 = codigo_pago1 + 1
'                  rscorrelativo!correlativo = Trim(Str(codigo_pago1))
'                  rscorrelativo.Update
'          End If
'          If rscorrelativo.State = 1 Then rscorrelativo.Close
'                '======== fin TERMINA GENERACION DE COMPROBANTE ========
''            End If                 ' REV JQA OCT-2009
'            '==== fin generación de correlativo ====
'          MsgBox "Comprobante : " & codigo_pago1 & vbCrLf & "Organismo :     " & v_por_fte(i, 3), vbInformation + vbOKOnly, " Generando el Comprobante..."
'          rstpagos!codigo_pago = codigo_pago1     'Cont_Comp
'          rstpagos!org_codigo = v_por_fte(j, 3)       'v_por_fte(i, 3)
'            'If i = 1 Then
'          rstpagos!uni_codigo = v_EstPoa(j, 16) 'v_EstPoa(I, 10) 'uni_codigo1   ' REV JQA OCT-2009
'          rstpagos!codigo_Categoria = v_EstPoa(j, 11) 'v_EstPoa(I, 11) 'codigo_categoria1
'          rstpagos!codigo_convenio = v_EstPoa(j, 12) 'codigo_convenio1
'          CONVE = v_EstPoa(j, 12)
'          CATEG = v_EstPoa(j, 11)
'            'End If
'          rstpagos!Codigo_orden = adoorigen!codigo_solicitud         'documento de respaldo
'          rstpagos!codigo_documento = "D13"         'documento de respaldo
'          rstpagos!codigo_solicitud = adoorigen!codigo_solicitud
'          rstpagos!codigo_unidad = adoorigen!codigo_unidad 'nuevo
'          rstpagos!fte_codigo = v_por_fte(j, 2)       'v_por_fte(i, 2)
'          rstpagos!justificacion = adoorigen!justificacion_solicitud   'adoorigen.txtjustifica
'          rstpagos!tipo_moneda = rstao_solicitud_detalle!tipo_moneda 'adoorigen!tipo_moneda   'DtCDenominacion_moneda.bounttext  '"Bs." 'DtCTipoMoneda.Text
'            'If i = 1 Then
''              If rstao_solicitud_detalle!por_fte_nal = 100 Or rstao_solicitud_detalle!por_fte_ext = 100 Then
'          rstpagos!monto_bolivianos = IIf(IsNull(rstpagos!monto_bolivianos), 0, rstpagos!monto_bolivianos) + (rstao_solicitud_detalle!monto_bolivianos)  'adoorigen!Monto_bolivianos   '- adoorigen!monto_bolivianos_contra  '* por_fte_ext1   'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_ext1
'          rstpagos!monto_dolares = IIf(IsNull(rstpagos!monto_dolares), 0, rstpagos!monto_dolares) + rstao_solicitud_detalle!monto_dolares 'adoorigen!monto_dolares   '- adoorigen!monto_dolares_contra  '* por_fte_ext1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_ext1
'          rstpagos!liquido_pagar = IIf(IsNull(rstpagos!monto_bolivianos), 0, rstpagos!monto_bolivianos) + (rstao_solicitud_detalle!monto_bolivianos)  'adoorigen!Monto_bolivianos   '- adoorigen!monto_bolivianos_contra  '* por_fte_ext1   'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_ext1
''              End If
'          If j = 1 Then
'              'If rstao_solicitud_detalle!monto_bolivianos > 0 Then
'                rstpagos!es_base = "S"
'          Else
'                rstpagos!es_base = "N"
'          End If
'            'End If
'            'rstpagos!liquido_pagar  = "0" 'Val(TxtLiquido.Text)
'          rstpagos!formulario = GlNombFor
'          If GlNombFor = "F04" Or GlNombFor = "F11" Then
'              rstpagos!es_licitacion = "S"
'              rstpagos!duracion_estimada_tiempo = adoorigen!duracion_estimada_tiempo
'              rstpagos!duracion_estimada_numero = adoorigen!duracion_estimada_numero
'              rstpagos!por_tiempo = adoorigen!por_tiempo
'              rstpagos!fecha_estimada_inicio = IIf(IsNull(adoorigen!fecha_estimada_inicio), Date, Format(adoorigen!fecha_estimada_inicio, "dd/mm/yyyy"))
'              rstpagos!Lista_adjunta = adoorigen!Lista_adjunta
'              rstpagos!periodo_de_trabajo = ""
'              rstpagos!tipo_formulario = "COM"
'              rstpagos!tipo_comp = "DAC"
'              rstpagos!estado_aprobacion = "S"
'              rstpagos!estado_compromiso = "S"
'              rstpagos!estado_devengado = ""
'          End If
'          If GlNombFor = "F05" Or GlNombFor = "F10" Then
'              rstpagos!es_licitacion = "N"
'              rstpagos!duracion_estimada_tiempo = adoorigen!duracion_estimada_tiempo
'              rstpagos!duracion_estimada_numero = adoorigen!duracion_estimada_numero
'              rstpagos!por_tiempo = adoorigen!por_tiempo
'              rstpagos!fecha_estimada_inicio = IIf(IsNull(adoorigen!fecha_estimada_inicio), Date, Format(adoorigen!fecha_estimada_inicio, "dd/mm/yyyy"))
'              rstpagos!Lista_adjunta = adoorigen!Lista_adjunta
'              rstpagos!periodo_de_trabajo = ""
'              rstpagos!tipo_formulario = "COM"
'              rstpagos!tipo_comp = "DAC"
'              rstpagos!estado_aprobacion = "S"
'              rstpagos!estado_compromiso = "S"
'              rstpagos!estado_devengado = ""
'          End If
'          If GlNombFor = "F03" Or GlNombFor = "F12" Then
'              rstpagos!tipo_formulario = "CYD"
'              rstpagos!tipo_comp = "DAC"
'              rstpagos!estado_aprobacion = "S"
'              rstpagos!estado_compromiso = "S"
'              rstpagos!estado_devengado = "S"
'          End If
'          If (GlNombFor = "F01") Then     'And i = 2
'              rstpagos!tipo_formulario = "REG"
'              rstpagos!estado_contabilidad = "P"
'              rstpagos!estado_aprobacion = "S"
'              rstpagos!estado_compromiso = ""
'              rstpagos!estado_devengado = ""
'              rstpagos!estado_pagado = "N"
'              rstpagos!estado_pagado = "N"
'              rstpagos!tipo_comp = "PCE"
'          End If
'          rstpagos!fecha_egreso = Format(Date, "dd/mm/yyyy") 'CDate(adoorigen!fecha_recepcion)   ', "dd/mm/aaaa
'          rstpagos!ges_gestion = Year(Date)
'          ges_gestion1 = Year(Date)
'          rstpagos!usr_usuario = GlUsuario
'          rstpagos!fecha_registro = Date  ' Format(Date, "dd/mm/aaaa
'          rstpagos!hora_registro = Format(Time, "hh:mm:ss")
'          rstpagos.Update
'          If rstpagos.State = 1 Then rstpagos.Close
'            '======== fin graba pagos ========
'          For i = 1 To rectot       'tot_form(ya no)
'            '======== ini graba pago_detalle ========
'            Set rstpago_detalle = New ADODB.Recordset
'            If rstpago_detalle.State = 1 Then rstpago_detalle.Close
'            'If GlNombFor = "F04" Or GlNombFor = "F05" Or GlNombFor = "F10" Or GlNombFor = "F11" Or GlNombFor = "F01" Then  '
'            'If i = 1 Then           ' REV JQA OCT-2009
'            '  rstpago_detalle.Open "select * from pago_detalle_espera where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenKeyset, adLockOptimistic
'            'Else
'            rstpago_detalle.Open "select * from pago_detalle where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenKeyset, adLockOptimistic
'            'End If                  ' REV JQA OCT-2009
'            'If rstpago_detalle.RecordCount > 0 Then
'            '  rstpago_detalle.MoveFirst
'            'Else
'              rstpago_detalle.AddNew
'            'End If
'            rstpago_detalle!codigo_pago = codigo_pago1      'Cont_Comp
'            rstpago_detalle!ges_gestion = ges_gestion1
'            rstpago_detalle!org_codigo = v_por_fte(j, 3)        'v_por_fte(i, 3)
'            rstpago_detalle!codigo_pago_detalle = rstpago_detalle.RecordCount
'            rstpago_detalle!par_codigo = v_EstPoa(i, 3) 'Par_Codigo1
'            rstpago_detalle!pro_programa = v_EstPoa(i, 6) 'pro_Programa1
''            rstpago_detalle!pro_subprograma = v_EstPoa(j, 7) 'Pro_SubPrograma1
'            rstpago_detalle!pro_proyecto = v_EstPoa(i, 8) 'Pro_Proyecto1
'            rstpago_detalle!pro_actividad = v_EstPoa(i, 9) 'Pro_Actividad1
'            rstpago_detalle!codigo_beneficiario = adoorigen!CI_aprueba
'            '==== ini porcentajes ====
'            rstpago_detalle!codigo_poa = v_EstPoa(i, 1)            'rstao_solicitud_detalle!codigo_poa
'            'If i = 1 Then          ' REV JQA OCT-2009
'              'rstpago_detalle!monto_total = Val(rstao_solicitud_detalle!monto_bolivianos)  'adoorigen!Monto_bolivianos   '- adoorigen!monto_bolivianos_contra  '* por_fte_ext1   'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_ext1
'              'rstpago_detalle!monto_bolivianos = Val(rstao_solicitud_detalle!monto_bolivianos)  'adoorigen!Monto_bolivianos   '- adoorigen!monto_bolivianos_contra  '* por_fte_ext1   'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_ext1
'              'rstpago_detalle!monto_dolares = Val(rstao_solicitud_detalle!monto_dolares)  'adoorigen!monto_dolares   '- adoorigen!monto_dolares_contra  '* por_fte_ext1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_ext1
'              'rstpago_detalle!monto_dolares_dev = Val(rstao_solicitud_detalle!monto_dolares)  'adoorigen!monto_dolares   '- adoorigen!monto_dolares_contra  '* por_fte_ext1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_ext1
'            rstpago_detalle!monto_total = v_EstPoa(i, 17)
'            rstpago_detalle!monto_bolivianos = v_EstPoa(i, 17)
'            rstpago_detalle!monto_dolares = v_EstPoa(i, 18)
'            rstpago_detalle!monto_dolares_dev = v_EstPoa(i, 18)
'            If j = 1 Then
'                 rstpago_detalle!Porcentaje = CDbl(rstao_solicitud_detalle!por_fte_ext)
'            Else
'                rstpago_detalle!Porcentaje = por_fte_nal1      'verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 1)
'            End If                 ' REV JQA OCT-2009
''            End If
'            '==== fin porcentajes ====
'            rstpago_detalle!Deducciones = 1             'Val(TxtDeducciones.Text)
'            rstpago_detalle!saldo_bolivianos = 0        'Val(TxtSaldo.Text)
'            rstpago_detalle!tipo_cambio = Val(rstao_solicitud_detalle!tipo_cambio) 'adoorigen!tipo_caMBIO    'adoorigen.adosolicitud.Recordset!tipo_cambio
'            rstpago_detalle!tipo_cambio_dev = Val(rstao_solicitud_detalle!tipo_cambio)
'            rstpago_detalle!estado_aprobacion = "N"
'            rstpago_detalle!fecha_pago = Format(Date, "DD/MM/YYYY")  ', "dd/mm/aaaa
'            rstpago_detalle!fecha_registro = Format(Date, "DD/MM/YYYY")
'            rstpago_detalle!usr_usuario = GlUsuario
'            rstpago_detalle!hora_registro = Format(Time, "hh:mm:ss")
'            rstpago_detalle.Update
'            '======== fin graba pago_detalle
'          Next
'          Set rstdestino = New ADODB.Recordset
'          If rstdestino.State = 1 Then rstdestino.Close
'          rstdestino.Open "select * from ao_solicitud where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
'          If rstdestino.RecordCount > 0 Then
'            rstdestino!estado_enviado = "S"
'            rstdestino!Status = "A"
'            rstdestino.Update
'            rstao_solicitud_recibido.AddNew
'            rstao_solicitud_recibido!ges_gestion = IIf(IsNull(adoorigen!ges_gestion), "", adoorigen!ges_gestion)
'            rstao_solicitud_recibido!codigo_solicitud = IIf(IsNull(adoorigen!codigo_solicitud), CStr(0), CStr(adoorigen!codigo_solicitud))
'            rstao_solicitud_recibido!formulario = IIf(IsNull(adoorigen!formulario), "", adoorigen!formulario)
'            rstao_solicitud_recibido!codigo_unidad = IIf(IsNull(adoorigen!codigo_unidad), "", adoorigen!codigo_unidad)
'            rstao_solicitud_recibido!justificacion_solicitud = IIf(IsNull(adoorigen!justificacion_solicitud), "", adoorigen!justificacion_solicitud)
'            rstao_solicitud_recibido!fecha_solicitud = Format(adoorigen!fecha_solicitud, "dd/mm/yyyy")
'            rstao_solicitud_recibido!swSubir = swSubir
'            rstao_solicitud_recibido!usr_usuario = GlUsuario
'            rstao_solicitud_recibido!fecha_registro = Format(Date, "dd/mm/yyyy")
'            rstao_solicitud_recibido!hora_registro = Format(Time, "hh:mm:ss")
'            rstao_solicitud_recibido.Update
'          End If
'          If rstdestino.State = 1 Then rstdestino.Close
'          rstao_solicitud_detalle.MoveNext
'        Next
'        db.CommitTrans
'        If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
'      End If
'    End If
'  End If
'
'  '======== tipo de formulario F01 CONTABILIZA========
'  'If GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "A") Then
'  If GlNombFor = "F01" Then
'    tot_reg = 0
'    Dim rsAuxDetalle As New ADODB.Recordset
'    Dim rstdetalle As New ADODB.Recordset
'    Set rsAuxDetalle = New ADODB.Recordset
'    If rsAuxDetalle.State = 1 Then rsAuxDetalle.Close
'    rsAuxDetalle.Open "select sum(monto_bolivianos) as AuxTotBs, sum(monto_dolares) as AuxTotSus, sum(monto_bolivianos_contra) as AuxTotBsCn, sum(monto_dolares_contra) as AuxTotSusCn from ao_Solicitud_detalle where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockReadOnly
'    Set rstdetalle = New ADODB.Recordset
'    If rstdetalle.State = 1 Then rstdetalle.Close
'    rstdetalle.Open "select * from ao_Solicitud_detalle where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockReadOnly
'    If rstdetalle.RecordCount < 1 Then
'      MsgBox "No se puede generar el asiento contable," & vbCrLf & "debido a que el registro no tiene el detalle de montos.", vbOKOnly + vbCritical, "Error al generar el asiento contablE..."
'      If rstdetalle.State = 1 Then rstdetalle.Close
'      Exit Sub
'    Else
'      tot_reg = 0
'      If rstdetalle!monto_bolivianos > 0 Then tot_reg = tot_reg + 1
'      If rstdetalle!monto_bolivianos_contra > 0 Then tot_reg = tot_reg + 1
'    End If
'    'INI JQA JUL-2005 **
'      Set rstfc_relacionador_poa_ppto = New ADODB.Recordset
'      If rstfc_relacionador_poa_ppto.State = 1 Then rstfc_relacionador_poa_ppto.Close
'      rstfc_relacionador_poa_ppto.Open "select * from fc_relacionador_poa_ppto where codigo_poa = '" & rstdetalle!codigo_poa & "'", db, adOpenKeyset, adLockReadOnly
'      If rstfc_relacionador_poa_ppto.RecordCount > 0 Then
'         ConvExt = rstfc_relacionador_poa_ppto!codigo_convenio
'         ConvNAl = rstfc_relacionador_poa_ppto!codigo_convenio_nal
'         CatExt = rstfc_relacionador_poa_ppto!codigo_Categoria
'         CatNal = IIf(IsNull(rstfc_relacionador_poa_ppto!CODIGO_Categoria_nal = True), "", rstfc_relacionador_poa_ppto!CODIGO_Categoria_nal)
'         'JQA OCT-2009 AQUI
'      Else
'        MsgBox "No se puede generar el asiento contable," & vbCrLf & "debido a problemas en el registro POA.", vbOKOnly + vbCritical, "Error al generar el asiento contable..."
'        Exit Sub
'      End If
'    'INI JQA JUL-2005 **
'    Set rstao_solicitud_recibido = New ADODB.Recordset
'    If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
'    rstao_solicitud_recibido.Open "SELECT * FROM ao_solicitud_recibido", db, adOpenKeyset, adLockOptimistic
'    db.BeginTrans
'    '======== ini registro de co_comprobante_M ========
'    Dim rstCodComp As New ADODB.Recordset
'    Set rstdestino = New ADODB.Recordset
'    For i = 1 To 2 'tot_reg
'      If rstdetalle!monto_bolivianos <= 0 And i = 1 Then
'        GoTo etiq
'      End If
'      If rstdetalle!monto_bolivianos_contra <= 0 And i = 2 Then
'        GoTo etiq
'      End If
'      '======== ini GENERA EL CODIGO DE COMPROBANTE ========
'      Set rstCodComp = New ADODB.Recordset
'      rstCodComp.CursorLocation = adUseClient
'      If rstCodComp.State = 1 Then rstCodComp.Close
'      rstCodComp.Open "select * from fc_Correl  where tipo_tramite = 'cmbte'", db, adOpenDynamic, adLockOptimistic
'      If rstCodComp.RecordCount > 0 Then
'        Cont_Comp = Val(rstCodComp!numero_correlativo)
'        Cont_Comp = Cont_Comp + 1
'        rstCodComp!numero_correlativo = Trim(Str(Cont_Comp))
'        rstCodComp.Update
'      End If
'      If rstCodComp.State = 1 Then rstCodComp.Close
'      '======== fin TERMINA GENERACION DE COMPROBANTE ========
'
'      '======== ini registro co_comprobantre_m ========
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from co_comprobante_m where Cod_Comp = 0", db, adOpenKeyset, adLockOptimistic
'      If rstdestino.RecordCount > 0 Then
'      End If
'      rstdestino.AddNew
'      rstdestino!Cod_Comp = Cont_Comp
'      rstdestino!cod_trans = codigo_pago1
'      If i = 1 Then
'        rstdestino!org_codigo = v_por_fte(i, 3) 'adoorigen!org_codigo_ext
'      End If
'
'      If i = 2 Then
'        rstdestino!org_codigo = v_por_fte(i, 3) '"999"  'rstdestino!org_codigo = adoorigen!org_codigo_contra
'      End If
'
'      rstdestino!cod_trans_detalle = 1
'      'rstdestino!Num_respaldo = adoorigen!codigo_unidad & "/" & Str(adoorigen!codigo_solicitud)
'      rstdestino!Num_respaldo = Str(adoorigen!codigo_solicitud)     'Respaldo
'      rstdestino!codigo_solicitud = (adoorigen!codigo_solicitud) 'adoorigen!codigo_unidad '& "/" & Str(adoorigen!codigo_solicitud)
'      rstdestino!codigo_unidad = (adoorigen!codigo_unidad)
'      rstdestino!Fecha_A = Format(Date, "dd/mm/yyyy")         'Format(adoorigen!fecha_solicitud, "dd/mm/yyyy")
'      rstdestino!codigo_beneficiario = adoorigen!CI_aprueba
'      rstdestino!Origen = "1"
'  'aqui fBuscaFteCorta(fte_1)
'      If i = 1 Then
'        'rstdestino!glosa = Trim(adoorigen!justificacion_solicitud) & " " & fBuscaOrgCorta(rstdetalle!org_codigo_ext) & ": " & Round((rstdetalle!monto_Bolivianos * 100 / (rstdetalle!monto_Bolivianos + rstdetalle!monto_bolivianos_contra)), 2) & "%"
'        rstdestino!glosa = Trim(adoorigen!justificacion_solicitud)
'      End If
'      If i = 2 Then
'        'rstdestino!glosa = Trim(adoorigen!justificacion_solicitud) & " " & fBuscaOrgCorta(rstdetalle!org_codigo_contra) & ": " & Round((rstdetalle!monto_bolivianos_contra * 100 / (rstdetalle!monto_Bolivianos + rstdetalle!monto_bolivianos_contra)), 2) & "%"
'        rstdestino!glosa = Trim(adoorigen!justificacion_solicitud)
'      End If
'      rstdestino!Status = "S"
'      rstdestino!ges_gestion = adoorigen!ges_gestion
'      rstdestino!codigo_documento = "D22"
'      rstdestino!tipo_comp = "PCE" 'IIf(adoorigen!codigo_tipo = "DEV", "CAD", IIf(adoorigen!codigo_tipo = "REC", "CAR", v_Tipo_Comp(i)))
'  '        rstdestino!tipo_moneda = adoorigen!tipo_moneda
'      rstdestino!usr_usuario = GlUsuario
'      rstdestino!fecha_registro = Date
'      rstdestino!hora_registro = Format(Time, "hh:mm:ss")
'      rstdestino!tipo_moneda = rstdetalle!tipo_moneda
'      rstdestino.Update
'      '======== fin registro co_comprobantre_m ========
'
'      '======== ini registra CO_diaRIO ========
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from co_diario where Cod_Comp = " & Cont_Comp, db, adOpenKeyset, adLockOptimistic
'      If rstdestino.RecordCount > 0 Then
'        rstdestino.MoveFirst
'      Else
'        rstdestino.AddNew
'        rstdestino!Cod_Comp = Cont_Comp
'        rstdestino!Cod_Comp_C = codigo_pago1
'        rstdestino!cod_trans_detalle = codigo_pago1
'      End If
'
'      rstdestino!tipo_comp = "PCE"
'      rstdestino!d_cuenta = "1127"
'  'y        rstdestino!D_Nombre = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
'      rstdestino!d_subcta1 = "02"
'      Select Case adoorigen!subcta2
'        Case "01" '"Regulares" 'Cargos de Cuenta Regulares
'          rstdestino!d_SubCta2 = "01"
'          rstdestino!d_Aux3 = "00"
'        Case "02" '"Otros" 'Cargos de Cuenta Otros
'          rstdestino!d_SubCta2 = "02"
'          rstdestino!d_Aux3 = "00"
'        Case "03"  '"PASE" 'Cargos de Cuenta PASE
'          rstdestino!d_SubCta2 = "03"
'          rstdestino!d_Aux3 = "10"
'
'      End Select
'      rstdestino!d_Aux1 = "01"
'      rstdestino!d_Aux2 = "09"
''      rstdestino!d_Aux3 = "00"
'      rstdestino!d_cta_larga = adoorigen!CI_aprueba
'      rstdestino!d_des_Larga = "-" ' CAMPO PARA ELIMINAR
'      If i = 1 Then
'        'rstdestino!d_montoBs = rstdetalle!monto_bolivianos
'        'rstdestino!d_montoDl = rstdetalle!monto_dolares
'        rstdestino!D_MontoBs = rsAuxDetalle!AuxTotBs
'        rstdestino!D_MontoDl = rsAuxDetalle!AuxTotSus
'        'rstdestino!d_ctaaux2 = rstdetalle!org_codigo_ext   'JQA NOV-
'      End If
'      If i = 2 Then
'        'rstdestino!d_montoBs = rstdetalle!monto_bolivianos_contra
'        'rstdestino!d_montoDl = rstdetalle!MONTO_DOLARES_CONTRA
'        rstdestino!D_MontoBs = rsAuxDetalle!AuxTotBsCn
'        rstdestino!D_MontoDl = rsAuxDetalle!AuxTotSusCn
'        'rstdestino!d_ctaaux2 = rstdetalle!org_codigo_contra  'JQA NOV-
'      End If
'      rstdestino!D_Cambio = rstdetalle!tipo_cambio
'      rstdestino!h_cuenta = "2116"
'  'Y        rstdestino!H_Nombre = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
'      rstdestino!h_subcta1 = "02"
'      rstdestino!h_subcta2 = "00"
'      rstdestino!h_Aux1 = "01"
'      rstdestino!h_Aux2 = "09"   'Y
'      rstdestino!h_Aux3 = "00"
'      rstdestino!h_cta_larga = adoorigen!CI_aprueba
'      rstdestino!h_des_Larga = "-"   ' CAMPO PARA ELIMINAR
'      If i = 1 Then
'        rstdestino!h_MontoBs = rsAuxDetalle!AuxTotBs
'        rstdestino!h_MontoDl = rsAuxDetalle!AuxTotSus
'        'INI JQA JUL-2005 **
'        rstdestino!h_ctaaux2 = ConvExt
'        rstdestino!D_CtaAux2 = ConvExt
'        rstdestino!d_ctaaux3 = IIf(IsNull(rstdetalle!aux3), "", rstdetalle!aux3)
'        'rsCo_diario!d_Aux3 = "10"
'        'rsCo_diario!d_ctaaux3 = DtCCodigo.Text
'      End If
'      If i = 2 Then
'        rstdestino!h_MontoBs = rstdetalle!monto_bolivianos_contra
'        rstdestino!h_MontoDl = rstdetalle!MONTO_DOLARES_CONTRA
'        rstdestino!h_ctaaux2 = ConvExt
'        rstdestino!D_CtaAux2 = ConvExt
'        'rstdestino!h_ctaaux2 = "S/C TGN" 'rstdetalle!codigo_convenio       'JQA OCT-2009
'        'rstdestino!D_CtaAux2 = "S/C TGN" 'rstdetalle!codigo_convenio
'        rstdestino!d_ctaaux3 = IIf(IsNull(rstdetalle!aux3), "", rstdetalle!aux3)
'      End If
'      rstdestino!h_Cambio = rstdetalle!tipo_cambio
'      'grabar convenios
'      'en h_ctaaux2 y en d_ctaaux2
''      rstdestino!h_ctaaux2 = rstdetalle!codigo_convenio
''      rstdestino!d_ctaaux2 = rstdetalle!codigo_convenio
'
'      rstdestino!usr_usuario = GlUsuario
'      rstdestino!fecha_registro = Date
'      rstdestino!hora_registro = Format(Time, "hh:mm:ss")
'      rstdestino.Update
'      If rstdestino.State = 1 Then rstdestino.Close
'      '======== fin registra co_diario ========
'etiq:
'    Next i
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from ao_solicitud where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
'    If rstdestino.RecordCount > 0 Then
'      rstdestino!estatus = "A"
'      rstdestino!estado_enviado = "S"
'      rstdestino.Update
'      rstao_solicitud_recibido.AddNew
'      rstao_solicitud_recibido!ges_gestion = IIf(IsNull(adoorigen!ges_gestion), "", adoorigen!ges_gestion)
'      rstao_solicitud_recibido!codigo_solicitud = IIf(IsNull(adoorigen!codigo_solicitud), 0, adoorigen!codigo_solicitud)
'      rstao_solicitud_recibido!formulario = IIf(IsNull(adoorigen!formulario), "", adoorigen!formulario)
'      rstao_solicitud_recibido!codigo_unidad = IIf(IsNull(adoorigen!codigo_unidad), "", adoorigen!codigo_unidad)
'      rstao_solicitud_recibido!justificacion_solicitud = IIf(IsNull(adoorigen!justificacion_solicitud), "", adoorigen!justificacion_solicitud)
'      'rstao_solicitud_recibido!swSubir = swSubir
'      rstao_solicitud_recibido!usr_usuario = GlUsuario
'      rstao_solicitud_recibido!fecha_registro = Format(Date, "dd/mm/yyyy")
'      rstao_solicitud_recibido!hora_registro = Format(Time, "hh:mm:ss")
'      rstao_solicitud_recibido.Update
'    End If
'    If rstdestino.State = 1 Then rstdestino.Close
'    If rstdetalle.State = 1 Then rstdetalle.Close
'    db.CommitTrans
'    If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
'  End If
'  '---- fin formulario f01 ----
'
'  '---- ini formulario F06 ----
'  If GlNombFor = "F06" Then
'
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from lo_pagos_conformidad where ges_gestion = '0' ", db, adOpenKeyset, adLockOptimistic
'    rstdestino.AddNew
'    rstdestino!ges_gestion = adoorigen!ges_gestion
'    rstdestino!codigo_unidad = adoorigen!codigo_unidad
'    rstdestino!codigo_grupo = adoorigen!codigo_solicitud_ant
'    rstdestino!NUMERO_PAGO = adoorigen!Nro_pagos
'    rstdestino!codigo_beneficiario = adoorigen!CI_aprueba
'
''    rstdestino!ges_gestion = adoorigen!ges_gestion
''    rstdestino!ges_gestion = adoorigen!ges_gestion
''    rstdestino!ges_gestion = adoorigen!ges_gestion
''    rstdestino!ges_gestion = adoorigen!ges_gestion
''    rstdestino!ges_gestion = adoorigen!ges_gestion
'
'    Set rstorigen = New ADODB.Recordset
'    If rstorigen.State = 1 Then rstorigen.Close
'    rstorigen.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
'    If rstorigen.RecordCount > 0 Then
'      rstdestino!tipo_moneda = rstorigen!tipo_moneda
'      rstdestino!monto_bs_ext = CDbl(rstorigen!monto_bolivianos)
'      rstdestino!monto_dol_ext = CDbl(rstorigen!monto_dolares)
'      rstdestino!conformidad = "S"
'      rstdestino!enviadoaudapre = "S"
'      rstdestino!confo_procesada = "N"
'    End If
'    rstdestino!usr_usuario = GlUsuario
'    rstdestino!fecha_registro = Format(Date, "dd/mm/yyyy")
'    rstdestino!hora_registro = Format(Time, "hh:mm:ss")
'    rstdestino.Update
'
'
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from ao_solicitud where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
'    If rstdestino.RecordCount > 0 Then
'      rstdestino!estatus = "A"
'      rstdestino.Update
'      Set rstao_solicitud_recibido = New ADODB.Recordset
'      If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
'      rstao_solicitud_recibido.Open "select * from ao_solicitud_recibido", db, adOpenKeyset, adLockOptimistic
'      rstao_solicitud_recibido.AddNew
'      rstao_solicitud_recibido!ges_gestion = IIf(IsNull(adoorigen!ges_gestion), "", adoorigen!ges_gestion)
'      rstao_solicitud_recibido!codigo_solicitud = IIf(IsNull(adoorigen!codigo_solicitud), 0, adoorigen!codigo_solicitud)
'      rstao_solicitud_recibido!formulario = IIf(IsNull(adoorigen!formulario), "", adoorigen!formulario)
'      rstao_solicitud_recibido!codigo_unidad = IIf(IsNull(adoorigen!codigo_unidad), "", adoorigen!codigo_unidad)
'      rstao_solicitud_recibido!justificacion_solicitud = IIf(IsNull(adoorigen!justificacion_solicitud), "", adoorigen!justificacion_solicitud)
'      'rstao_solicitud_recibido!swSubir = swSubir
'      rstao_solicitud_recibido!usr_usuario = GlUsuario
'      rstao_solicitud_recibido!fecha_registro = Format(Date, "dd/mm/yyyy")
'      rstao_solicitud_recibido!hora_registro = Format(Time, "hh:mm:ss")
'      rstao_solicitud_recibido.Update
'    End If
'    If rstdestino.State = 1 Then rstdestino.Close
'
''captu    '  0 descripcion_grupo varchar 50  0 0 0 ('')  0     0
''captu    '  0 concepto  varchar 250 0 0 0 ('')  0     0
'
'    '  0 antecedente varchar 250 0 0 0 ('')  0     0
'    '  0 nombre_proveedor  varchar 30  0 0 0 ('')  0     0
'    '  0 idBeneficiario varchar 15  0 0 0 ('')  0     0
'
'    '  0 fecha_envio datetime  8 0 0 0 (getdate()) 0     0
'    '  0 NCite_conformidad char  15  0 0 0 ('')  0     0
'    '  0 FCite_conformidad datetime  8 0 0 1 (getdate()) 0     0
'    '  0 migrado char  1 0 0 0 ('N') 0     0
'    '  0 Usr_Usuario varchar 15  0 0 0 ('')  0     0
'    '  0 Fecha_Registro  datetime  8 0 0 0 (getdate()) 0     0
'    '  0 Hora_Registro varchar 8 0 0 0 ('')  0     0
''captu    '  0 Emite_Factura char  1 0 0 0 ('N') 0     0
'    '  0 Sesion  int 4 10  0 0 (0) 0     0
'    '  0 porcentaje_pago numeric 9 18  2 0 (100) 0     0
'  End If
'  '---- fin formulario f06 ----
'
'
'End Sub


Private Sub val_presupF04(adoorigen, GlNombFor)
        
  'If (GlNombFor <> "F02") And (GlNombFor <> "F06") Or (GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "R" Or Trim(adoorigen!tipo_bien_Cta_doc) = "C")) Then
  If (GlNombFor = "F04" And Trim(adoorigen!tipo_bien_Cta_doc) = "A") Then
    Set rstao_solicitud_detalle = New ADODB.Recordset
    If rstao_solicitud_detalle.State = 1 Then rstao_solicitud_detalle.Close
    rstao_solicitud_detalle.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockReadOnly
    If rstao_solicitud_detalle.RecordCount > 0 Then
      rectot = rstao_solicitud_detalle.RecordCount
      Fte_contraparte1 = fBuscaFte(rstao_solicitud_detalle!org_codigo_contra)
      Org_Contraparte1 = rstao_solicitud_detalle!org_codigo_contra
      Dim v_EstPoa(50, 14)
    End If
    If Not (rstao_solicitud_detalle.BOF) Then rstao_solicitud_detalle.MoveFirst
    For i = 1 To rstao_solicitud_detalle.RecordCount            ' primer i
      Set rstfc_relacionador_poa_ppto = New ADODB.Recordset
      If rstfc_relacionador_poa_ppto.State = 1 Then rstfc_relacionador_poa_ppto.Close
      rstfc_relacionador_poa_ppto.Open "select * from fc_relacionador_poa_ppto where codigo_poa = '" & rstao_solicitud_detalle!codigo_poa & "'", db, adOpenKeyset, adLockReadOnly
      If rstfc_relacionador_poa_ppto.RecordCount > 0 Then
        'aqui se puede definir porcentaje
        v_EstPoa(i, 1) = rstao_solicitud_detalle!codigo_poa
        v_EstPoa(i, 2) = rstfc_relacionador_poa_ppto!da          'Dirección Administrativa JGCA 10/08/2007
        v_EstPoa(i, 3) = rstfc_relacionador_poa_ppto!par_codigo             'Par_Codigo1
        v_EstPoa(i, 4) = rstfc_relacionador_poa_ppto!fte_codigo             'fte_codigo1    JQA JUL-2005
        v_EstPoa(i, 5) = rstfc_relacionador_poa_ppto!org_codigo             'Org_Codigo1
        v_EstPoa(i, 6) = rstfc_relacionador_poa_ppto!pro_programa           'pro_Programa1
        v_EstPoa(i, 7) = rstfc_relacionador_poa_ppto!uni_codigo       'Pro_SubPrograma1
        v_EstPoa(i, 8) = rstfc_relacionador_poa_ppto!pro_proyecto           'Pro_Proyecto1
        v_EstPoa(i, 9) = rstfc_relacionador_poa_ppto!pro_actividad          'Pro_Actividad1
        v_EstPoa(i, 10) = rstfc_relacionador_poa_ppto!codigo_unidad            'uni_codigo1
        v_EstPoa(i, 11) = rstfc_relacionador_poa_ppto!codigo_categoria  'IIf(IsNull(rstfc_relacionador_poa_ppto!Categoria), rstfc_relacionador_poa_ppto!codigo_Categoria, rstfc_relacionador_poa_ppto!Categoria) 'codigo_categoria1    JQA JUL-2005
        v_EstPoa(i, 12) = rstfc_relacionador_poa_ppto!codigo_convenio       'codigo_convenio1
        por_fte_ext1 = rstfc_relacionador_poa_ppto!por_ext
        por_fte_nal1 = rstfc_relacionador_poa_ppto!por_nal
        If rstfc_relacionador_poa_ppto!por_ext = 100 Then                   'JQA JUL-2005
            v_EstPoa(i, 13) = rstfc_relacionador_poa_ppto!fte_codigo        'fte_codigo2   JQA JUL-2005
            v_EstPoa(i, 14) = rstfc_relacionador_poa_ppto!org_codigo        'Org_Codigo2   JQA JUL-2005
            cat_nal = IIf(IsNull(rstfc_relacionador_poa_ppto!Categoria), rstfc_relacionador_poa_ppto!codigo_categoria, rstfc_relacionador_poa_ppto!Categoria) 'codigo_categoria1    JQA JUL-2005
            conv_nal = rstfc_relacionador_poa_ppto!codigo_convenio          'codigo_convenio1
            tot_form = 1
        Else
            v_EstPoa(i, 13) = rstfc_relacionador_poa_ppto!fte_codigo        'fte_codigo2   JQA NOV-2008
            v_EstPoa(i, 14) = rstfc_relacionador_poa_ppto!org_codigo        'Org_Codigo2   JQA NOV-2008
            cat_nal = IIf(IsNull(rstfc_relacionador_poa_ppto!CODIGO_Categoria_nal), "FIN_PROPIO", rstfc_relacionador_poa_ppto!CODIGO_Categoria_nal)  'codigo_categoria2    JQA JUL-2005
            conv_nal = IIf(IsNull(rstfc_relacionador_poa_ppto!codigo_convenio_nal), "FIN_PROPIO", rstfc_relacionador_poa_ppto!codigo_convenio_nal)      'codigo_convenio2
            tot_form = 2
        End If
'        If rstao_solicitud_detalle!org_codigo_contra = "" Or rstao_solicitud_detalle!org_codigo_contra = "-" Then
'          v_EstPoa(i, 13) = "10"
'          v_EstPoa(i, 14) = "111"
'        Else
'          v_EstPoa(i, 13) = fBuscaFte(rstao_solicitud_detalle!org_codigo_contra)
'          v_EstPoa(i, 14) = rstao_solicitud_detalle!org_codigo_contra
'        End If
        If rstfc_relacionador_poa_ppto.State = 1 Then rstfc_relacionador_poa_ppto.Close
        Dim rstfo_formulacion_gasto As New ADODB.Recordset
        Set rstfo_formulacion_gasto = New ADODB.Recordset
        If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
        'rstfo_formulacion_gasto.Open "select * from fo_formulacion_gasto where pro_programa='" & pro_Programa1 & "' and pro_subprograma='" & Pro_SubPrograma1 & "' and pro_proyecto='" & Pro_Proyecto1 & "' and pro_actividad='" & Pro_Actividad1 & "' and par_codigo='" & Par_Codigo1 & "' and org_codigo= '" & Org_Codigo1 & "'", db, adOpenKeyset, adLockOptimistic
        rstfo_formulacion_gasto.Open "select * from fo_formulacion_gasto where pro_programa='" & v_EstPoa(i, 6) & "' and pro_proyecto='" & v_EstPoa(i, 8) & "' and pro_actividad='" & v_EstPoa(i, 9) & "' and par_codigo='" & v_EstPoa(i, 3) & "' and org_codigo= '" & v_EstPoa(i, 5) & "'", db, adOpenKeyset, adLockOptimistic
        If Not (rstfo_formulacion_gasto.EOF) Then
          If (rstfo_formulacion_gasto!FGS_VIGENTE - rstfo_formulacion_gasto!FGS_compromiso < rstao_solicitud_detalle!monto_bolivianos) Then  'adoorigen         'adoorigen.adosolicitud.Recordset!monto_dolares ) Then
            'JQA 07/12/01
'            swSubir = "No existe Presup"
'            MsgBox "NO EXISTE Presupuesto para dar curso a la Solicitud ...", vbOKOnly, "ERROR"
'            swpresup = 0
'            Exit Sub
            'JQA 07/12/01
            swpresup = 1    'Borrar despues de habilitar JQA
          Else
            'JQA 07/12/01
            'rstfo_formulacion_gasto!0  = rstfo_formulacion_gasto!fgs_precompromiso  + rstao_solicitud_detalle!monto_bolivianos
            'rstfo_formulacion_gasto.Update
            'JQA 07/12/01
            swpresup = 1
            swSubir = "SI correcto"
          End If
          If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
            swpresup = 1
          Else
            'JQA 07/12/01
'          MsgBox "NO EXISTE Estructura presupuestaria...", vbOKOnly, "ERROR ..."
'          swSubir = "NO Error Estruc.Ppto"
'          swpresup = 0
'          Exit Sub
            'JQA 07/12/01
            swpresup = 1    'Borrar despues de habilitar JQA
          End If
      Else
        MsgBox "NO Existe POA ... ", vbOKOnly, "ERROR ..."
        swSubir = "No existe POA"
        swpresup = 0
        Exit Sub
      End If
'          Else
'            swpresup = 1
'          End If
      rstao_solicitud_detalle.MoveNext
    Next            'fin del primer i

    If swpresup = 1 Then        'ini swpresup
      If GlNombFor = "F11" Or GlNombFor = "F04" Then '(GlNombFor = "F04" And (Trim(adoorigen!tipo_bien_Cta_doc) = "R" Or Trim(adoorigen!tipo_bien_Cta_doc) = "C")) Then
        If (rstao_solicitud_detalle.RecordCount > 0) And (Not rstao_solicitud_detalle.BOF) Then rstao_solicitud_detalle.MoveFirst
        Set rstao_solicitud_recibido = New ADODB.Recordset
        If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
        rstao_solicitud_recibido.Open "SELECT * FROM ao_solicitud_recibido", db, adOpenKeyset, adLockOptimistic
        db.BeginTrans
        'por_fte_ext
        'por_fte_nal
        For j = 1 To rstao_solicitud_detalle.RecordCount        ' del j
          'j = 2
          v_por_fte(1, 1) = 100 'por_fte_ext1
          v_por_fte(1, 2) = v_EstPoa(j, 4) 'fte_codigo1
          v_por_fte(1, 3) = v_EstPoa(j, 5) 'Org_Codigo1

          v_por_fte(2, 1) = 0 'por_fte_nal1
          v_por_fte(2, 2) = v_EstPoa(j, 13) 'Fte_contraparte1
          v_por_fte(2, 3) = v_EstPoa(j, 14) 'Org_Contraparte1

          v_por_fte(3, 1) = 0 'por_fte_nal1
          v_por_fte(3, 2) = v_EstPoa(j, 13) 'Fte_contraparte1
          v_por_fte(3, 3) = v_EstPoa(j, 14) 'Org_Contraparte1

          Dim SwEsBase As Integer
          Dim ValEsBase As Double
'          ValEsBase = v_por_fte(1, 1)
'          For I = 1 To tot_form
'            If v_por_fte(I, 1) > ValEsBase Then
'              SwEsBase = I
'              ValEsBase = v_por_fte(I, 1)
'            End If
'          Next
'AQUI UN SOLO FINANCIADOR
'la variable "j" distribuye al Preventivo y al Comprometido cada reg de ao_solicitud_detalle
        'k = 1
'        While (k <= 2)
'        begin
         prev_dev = 1     ' 1 p/pagos_espera y 2 p/pagos
         'prev_dev = 1       ' 1 solo p/pagos_espera
         For k = 1 To prev_dev
            '        Print rstpagos!monto_bolivianos
          'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
          '          Set rstpagos = New ADODB.Recordset
'          If rstpagos.State = 1 Then rstpagos.Close
'          rstpagos.Open "select * from pagos where codigo_pago = ''", db, adOpenKeyset, adLockOptimistic
'          rstpagos.AddNew
'            '==== ini generación de correlativo ====
'          Set rscorrelativo = New ADODB.Recordset
'          If rscorrelativo.State = 1 Then rscorrelativo.Close
'              'rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
'              '======== ini GENERA EL CODIGO DE COMPROBANTE ========
'          Set rscorrelativo = New ADODB.Recordset
'          rscorrelativo.CursorLocation = adUseClient
'          If rscorrelativo.State = 1 Then rscorrelativo.Close
'          'rscorrelativo.Open "select * from fc_Correlativos  where org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenDynamic, adLockOptimistic
'          rscorrelativo.Open "select * from fc_organismo_financiamiento where org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenDynamic, adLockOptimistic
'          If rscorrelativo.RecordCount > 0 Then
'                  codigo_pago1 = Val(rscorrelativo!correlativo)
'                  codigo_pago1 = codigo_pago1 + 1
'                  rscorrelativo!correlativo = Trim(Str(codigo_pago1))
'                  rscorrelativo.Update
'          End If
'          If rscorrelativo.State = 1 Then rscorrelativo.Close
'                '======== fin TERMINA GENERACION DE COMPROBANTE ========
'            '==== fin generación de correlativo ====
'          MsgBox "Comprobante : " & codigo_pago1 & vbCrLf & "Organismo :     " & v_por_fte(i, 3), vbInformation + vbOKOnly, " Generando el Comprobante..."
'          rstpagos!codigo_pago = codigo_pago1     'Cont_Comp
          
          'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
          For i = 1 To tot_form         'dos (segundo i)
            If k = 1 Then
                Set rstpagos = New ADODB.Recordset
                If rstpagos.State = 1 Then rstpagos.Close
                rstpagos.Open "select * from pagos where codigo_pago = ''", db, adOpenKeyset, adLockOptimistic
                'db.Execute " EXEC edGeneraCodigoPago @org_codigo_ida, @GCODIGO_PAGO OUT "
                'db.Execute " EXEC edGeneraCodigoPago v_por_fte(i, 3), @GCODIGO_PAGO OUT "
                'codigo_pago1 = GCODIGO_PAGO
'                Set rscorrelativo = New ADODB.Recordset
'                If rscorrelativo.State = 1 Then rscorrelativo.Close
'                rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
            End If
            rstpagos.AddNew
            '======== ini GENERA EL CODIGO DE COMPROBANTE ========
          Set rscorrelativo = New ADODB.Recordset
          rscorrelativo.CursorLocation = adUseClient
          If rscorrelativo.State = 1 Then rscorrelativo.Close
          'rscorrelativo.Open "select * from fc_Correlativos  where org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenDynamic, adLockOptimistic
          rscorrelativo.Open "select * from fc_organismo_financiamiento where org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenDynamic, adLockOptimistic
          If rscorrelativo.RecordCount > 0 Then
                  codigo_pago1 = Val(rscorrelativo!correlativo)
                  codigo_pago1 = codigo_pago1 + 1
                  rscorrelativo!correlativo = Trim(Str(codigo_pago1))
                  rscorrelativo.Update
          End If
          If rscorrelativo.State = 1 Then rscorrelativo.Close
                '======== fin TERMINA GENERACION DE COMPROBANTE ========
'            End If                 ' REV JQA OCT-2009
            '==== fin generación de correlativo ===='            'MsgBox "Comprobante : " & codigo_pago1 & vbCrLf & "Organismo :     " & rstpagos!org_codigo, vbInformation + vbOKOnly, " Generando el Comprobante..."
            'MsgBox "Comprobante : " & codigo_pago1 & vbCrLf & "Organismo :     " & v_por_fte(i, 3), vbInformation + vbOKOnly, " Generando el Comprobante..."
            rstpagos!org_codigo = v_por_fte(j, 3)
            rstpagos!codigo_pago = codigo_pago1
            If i = 1 Then
              rstpagos!da = IIf((v_EstPoa(j, 2) = ""), "02", v_EstPoa(j, 2))        'Dirección Administrativa JGCA 10/08/2007
              rstpagos!uni_codigo = IIf((v_EstPoa(j, 7) = ""), "PLAN", v_EstPoa(j, 7))        'uni_codigo1
              rstpagos!codigo_categoria = v_EstPoa(j, 11) 'v_EstPoa(I, 11) 'codigo_categoria1
              rstpagos!codigo_convenio = v_EstPoa(j, 12) 'codigo_convenio1
              CONVE = v_EstPoa(j, 12)
              CATEG = v_EstPoa(j, 11)
              rstpagos!es_base = "S"
              CODPAG = codigo_pago1
              rstpagos!nro_comprobante_anterior = Val(CODPAG)
            End If
            If i = 2 Then
               rstpagos!da = IIf((v_EstPoa(i - 1, 2) = ""), "02", v_EstPoa(i - 1, 2))      'Dirección Administrativa JGCA 10/08/2007
               rstpagos!uni_codigo = IIf((v_EstPoa(i - 1, 7) = ""), "PLAN", v_EstPoa(i - 1, 7))     'uni_codigo1
'              If rstao_solicitud_detalle!por_fte_nal = 100 Or rstao_solicitud_detalle!por_fte_ext = 100 Then
'                rstpagos!codigo_categoria = v_EstPoa(I - 1, 11)
'                rstpagos!codigo_convenio = v_EstPoa(I - 1, 12)
'              Else
'                v_por_fte(I, 3) = "S/C TGNP"
                'rstpagos!codigo_convenio = fbusCatConv(v_por_fte(i, 3), 1)  '"S/C TGN"  'v_EstPoa(i - 1, 12) 'codigo_convenio1 JQA JUL-2005
'                Print v_por_fte(j, 3)
                'rstpagos!codigo_Categoria = fbusCatConv(v_por_fte(i, 3), 2) '"S/C TGN" 'v_EstPoa(i - 1, 11) 'codigo_categoria1  JQA JUL-2005
'              End If
                rstpagos!codigo_convenio = conv_nal     'codigo_convenio2  JQA JUL-2005
                rstpagos!codigo_categoria = cat_nal    'codigo_categoria1  JQA JUL-2005
                rstpagos!nro_comprobante_anterior = Val(CODPAG)
                rstpagos!es_base = "N"
            End If

            If i = 3 Then
              rstpagos!da = v_EstPoa(1, 2)          'Dirección Administrativa JGCA 10/08/2007
              rstpagos!uni_codigo = v_EstPoa(1, 7) 'uni_codigo1
'              rstpagos!codigo_categoria = "S/C TGN" 'v_EstPoa(i - 1, 11) 'codigo_categoria1
'              rstpagos!codigo_convenio = "S/C TGN"  'v_EstPoa(i - 1, 12) 'codigo_convenio1
              'rstpagos!codigo_convenio = fbusCatConv(v_por_fte(i, 3), 1)  '"S/C TGN"  'v_EstPoa(i - 1, 12) 'codigo_convenio1
              'rstpagos!codigo_Categoria = fbusCatConv(v_por_fte(i, 3), 2) '"S/C TGN" 'v_EstPoa(i - 1, 11) 'codigo_categoria1
              rstpagos!codigo_convenio = conv_nal     'codigo_convenio2  JQA JUL-2005
              rstpagos!codigo_categoria = cat_nal    'codigo_categoria1  JQA JUL-2005
              rstpagos!es_base = "N"
            End If
            rstpagos!codigo_solicitud = adoorigen!codigo_solicitud
            rstpagos!codigo_unidad = adoorigen!codigo_unidad
            rstpagos!fte_codigo = v_por_fte(i, 2)
            rstpagos!Codigo_orden = adoorigen!codigo_solicitud
            rstpagos!codigo_documento = "D20"
            rstpagos!Deducciones = 1
            rstpagos!justificacion = adoorigen!justificacion_solicitud   'adoorigen.txtjustifica
            rstpagos!observaciones = adoorigen!caracteristicas
            rstpagos!tipo_moneda = rstao_solicitud_detalle!tipo_moneda 'adoorigen!tipo_moneda   'DtCDenominacion_moneda.bounttext  '"Bs." 'DtCTipoMoneda.Text
            If i = 1 Then
                rstpagos!monto_bolivianos = IIf(IsNull(rstpagos!monto_bolivianos), 0, rstpagos!monto_bolivianos) + (rstao_solicitud_detalle!monto_bolivianos)
                rstpagos!monto_dolares = IIf(IsNull(rstpagos!monto_dolares), 0, rstpagos!monto_dolares) + rstao_solicitud_detalle!monto_dolares
            End If
            If i = 2 Then
              If v_EstPoa(j, 12) <> "FIN_PROPIO" Then
                'ext1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 1)
                'tgn1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 2)
                ext1 = por_fte_ext1
                tgn1 = por_fte_nal1
                'abel 2004
                If IsNull(rstpagos!monto_bolivianos) Then
                    rstpagos!monto_bolivianos = 0
                Else
                    rstpagos!monto_bolivianos = IIf(IsNull(rstpagos!monto_bolivianos), 0, rstpagos!monto_bolivianos) + ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
                End If
                If IsNull(rstpagos!monto_dolares) Then
                    rstpagos!monto_dolares = 0
                Else
                    rstpagos!monto_dolares = IIf(IsNull(rstpagos!monto_dolares), 0, rstpagos!monto_dolares) + ((rstao_solicitud_detalle!monto_dolares_contra * tgn1) / (100 - ext1)) 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
                End If
              Else
                rstpagos!monto_bolivianos = Val(rstao_solicitud_detalle!monto_bolivianos_contra)
                rstpagos!monto_dolares = Val(rstao_solicitud_detalle!monto_dolares_contra)
              End If
              If rstao_solicitud_detalle!monto_bolivianos > 0 Then
                rstpagos!es_base = "N"
              Else
                rstpagos!es_base = "S"
              End If
            End If
            If i = 3 Then
              'tgn1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 3)
              tgn1 = por_fte_nal1
              rstpagos!monto_bolivianos = IIf(IsNull(rstpagos!monto_bolivianos), 0, rstpagos!monto_bolivianos) + ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
              rstpagos!monto_dolares = IIf(IsNull(rstpagos!monto_dolares), 0, rstpagos!monto_dolares) + ((rstao_solicitud_detalle!monto_dolares_contra * tgn1) / (100 - ext1)) 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
              rstpagos!es_base = "I"
            End If

            If k = 1 Then
              'rstpago_detalle.Open "select * from pago_detalle_espera where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenKeyset, adLockOptimistic
              rstpagos!tipo_comp = "DAC"
            Else
              'rstpago_detalle.Open "select * from pago_detalle where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenKeyset, adLockOptimistic
              'rstpago_detalle.Open "select * from pago_detalle where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '999' ", db, adOpenKeyset, adLockOptimistic
              'rstpagos!tipo_comp = "PCE"       ' DIC-2008
              rstpagos!tipo_comp = "DAC"
            End If
            rstpagos!liquido_pagar = IIf(IsNull(rstpagos!monto_bolivianos), 0, rstpagos!monto_bolivianos) + ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
            'rstpagos!tipo_formulario = GlNombFor
            rstpagos!formulario = GlNombFor
            'rstpagos!estado_aprobacion  = "X"
            rstpagos!estado_devengado = ""
'            If GlNombFor = "F04" Then
'              rstpagos!estado_compromiso = "S"
'              rstpagos!es_licitacion = "S"
'            'AQUI ULTIMO
'            Else
'              rstpagos!estado_compromiso = "N"
''              rstpagos!codigo_poa = rstao_solicitud_detalle!codigo_poa 'adoorigen!codigo_poa
'            End If
            If GlNombFor = "F11" Or GlNombFor = "F04" Then
              rstpagos!es_licitacion = "D"
              rstpagos!tipo_formulario = "COM"
              rstpagos!estado_compromiso = "S"
              rstpagos!estado_devengado = ""
              rstpagos!estado_pagado = ""
            End If
            If GlNombFor = "F05" Or GlNombFor = "F10" Then
              rstpagos!duracion_estimada_tiempo = adoorigen!duracion_estimada_tiempo
              rstpagos!duracion_estimada_numero = adoorigen!duracion_estimada_numero
              rstpagos!por_tiempo = adoorigen!por_tiempo
              rstpagos!estado_compromiso = "S"
              rstpagos!estado_devengado = ""
              rstpagos!fecha_estimada_inicio = IIf(IsNull(adoorigen!fecha_estimada_inicio), Date, Format(adoorigen!fecha_estimada_inicio, "dd/mm/yyyy"))
              rstpagos!Lista_adjunta = adoorigen!Lista_adjunta
              rstpagos!periodo_de_trabajo = ""
            End If
            'rstpagos!estado_devengado  = ""
            'rstpagos!estado_pagado  = ""
            If GlNombFor = "F03" Or GlNombFor = "F12" Then
              rstpagos!tipo_formulario = "CYD"
              rstpagos!estado_compromiso = "N"
              rstpagos!estado_devengado = "N"
            'Else
            '  rstpagos!tipo_formulario = "COM"
            End If
            'If (GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "R" Or Trim(adoorigen!tipo_bien_Cta_doc) = "C")) Then
            rstpagos!fecha_egreso = Format(Date, "dd/mm/yyyy") 'CDate(adoorigen!fecha_recepcion)   ', "dd/mm/aaaa
            rstpagos!ges_gestion = Year(Date)
            ges_gestion1 = Year(Date)
            rstpagos!usr_usuario = GlUsuario
            rstpagos!fecha_registro = Date  ' Format(Date, "dd/mm/aaaa
            rstpagos!hora_registro = Format(Time, "hh:mm:ss")
            
            rstpagos.Update
            If rstpagos.State = 1 Then rstpagos.Close
            '======== fin graba pagos ========

            '======== ini graba pago_detalle ========
            Set rstpago_detalle = New ADODB.Recordset
            If rstpago_detalle.State = 1 Then rstpago_detalle.Close
            'If GlNombFor = "F04" Or GlNombFor = "F05" Or GlNombFor = "F10" Or GlNombFor = "F11" Then  'Or GlNombFor = "F06"
            If k = 1 Then 'Or GlNombFor = "F06"
              rstpago_detalle.Open "select * from pago_detalle where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenKeyset, adLockOptimistic
              'rstpago_detalle.Open "select * from pago_detalle where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '999' ", db, adOpenKeyset, adLockOptimistic
            End If

            If rstpago_detalle.RecordCount > 0 Then
              rstpago_detalle.MoveFirst
            Else
              rstpago_detalle.AddNew
            End If
            rstpago_detalle!codigo_pago = codigo_pago1
            rstpago_detalle!ges_gestion = ges_gestion1
            If k = 1 Then
              rstpago_detalle!org_codigo = v_por_fte(i, 3)
            Else
              'rstpago_detalle!org_codigo = "999"           ' VERIFICAR DIC-2008
              rstpago_detalle!org_codigo = v_por_fte(i, 3)
            End If
            rstpago_detalle!codigo_pago_detalle = rstpago_detalle.RecordCount

            rstpago_detalle!par_codigo = v_EstPoa(j, 3) 'Par_Codigo1
            rstpago_detalle!pro_programa = v_EstPoa(j, 6) 'pro_Programa1
'            rstpago_detalle!pro_subprograma = v_EstPoa(j, 7) 'Pro_SubPrograma1
            rstpago_detalle!pro_proyecto = v_EstPoa(j, 8) 'Pro_Proyecto1
            rstpago_detalle!pro_actividad = v_EstPoa(j, 9) 'Pro_Actividad1
            rstpago_detalle!codigo_beneficiario = adoorigen!CI_aprueba

            '==== ini porcentajes ====

            rstpago_detalle!codigo_poa = rstao_solicitud_detalle!codigo_poa 'adoorigen!codigo_poa
            If i = 1 Then
              rstpago_detalle!monto_total = CDbl(rstao_solicitud_detalle!monto_bolivianos)  'adoorigen!Monto_bolivianos   '- adoorigen!monto_bolivianos_contra  '* por_fte_ext1   'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_ext1
              rstpago_detalle!monto_bolivianos = CDbl(rstao_solicitud_detalle!monto_bolivianos)  'adoorigen!Monto_bolivianos   '- adoorigen!monto_bolivianos_contra  '* por_fte_ext1   'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_ext1
              rstpago_detalle!monto_dolares = CDbl(rstao_solicitud_detalle!monto_dolares)  'adoorigen!monto_dolares   '- adoorigen!monto_dolares_contra  '* por_fte_ext1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_ext1
              rstpago_detalle!monto_dolares_dev = CDbl(rstao_solicitud_detalle!monto_dolares)
              If GlNombFor = "F05" Or GlNombFor = "F04" Or GlNombFor = "F10" Or GlNombFor = "F11" Or GlNombFor = "F01" Then
                rstpago_detalle!Porcentaje = CDbl(rstao_solicitud_detalle!por_fte_ext)
                'rstpago_detalle!Porcentaje = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 1)
              End If
            End If
            If i = 2 Then
'              rstpago_detalle!monto_total = rstao_solicitud_detalle!monto_bolivianos_contra  'adoorigen!monto_bolivianos_contra   'adoorigen!monto_bolivianos  * por_fte_nal1 'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_nal1
'              rstpago_detalle!monto_dolares = rstao_solicitud_detalle!monto_dolares_contra 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
              If v_EstPoa(j, 12) <> "FIN_PROPIO" Then
                'ext1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 1)
                'tgn1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 2)
                ext1 = por_fte_ext1
                tgn1 = por_fte_nal1
                'abel 2004
                If rstao_solicitud_detalle!monto_bolivianos_contra <> 0 Then
                    rstpago_detalle!monto_total = ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
                End If
                'abel 2004
                If rstao_solicitud_detalle!monto_dolares_contra <> 0 Then
                    rstpago_detalle!monto_dolares = ((rstao_solicitud_detalle!monto_dolares_contra * tgn1) / (100 - ext1)) 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
                End If
                If GlNombFor = "F05" Or GlNombFor = "F04" Or GlNombFor = "F10" Or GlNombFor = "F11" Or GlNombFor = "F01" Then
                  rstpago_detalle!Porcentaje = CDbl(rstao_solicitud_detalle!por_fte_nal)
                End If
              Else
                rstpago_detalle!monto_total = Val(rstao_solicitud_detalle!monto_bolivianos_contra)
                rstpago_detalle!monto_dolares = Val(rstao_solicitud_detalle!monto_dolares_contra)
                If GlNombFor = "F05" Or GlNombFor = "F04" Or GlNombFor = "F10" Or GlNombFor = "F11" Or GlNombFor = "F01" Then
                  rstpago_detalle!Porcentaje = CDbl(rstao_solicitud_detalle!por_fte_nal)
                  'rstpago_detalle!Porcentaje = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 2)
                End If
              End If
'''              rstpago_detalle!monto_total = rstao_solicitud_detalle!monto_bolivianos_contra  'adoorigen!monto_bolivianos_contra   'adoorigen!monto_bolivianos  * por_fte_nal1 'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_nal1
'''              rstpago_detalle!monto_dolares = rstao_solicitud_detalle!monto_dolares_contra 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
            End If
            If i = 3 Then
              'tgn1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 3)
              tgn1 = por_fte_nal1
              rstpago_detalle!monto_total = ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
              rstpago_detalle!monto_dolares = ((rstao_solicitud_detalle!monto_dolares_contra * tgn1) / (100 - ext1)) 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
              If GlNombFor = "F05" Or GlNombFor = "F04" Or GlNombFor = "F10" Or GlNombFor = "F11" Then
                rstpago_detalle!Porcentaje = tgn1
              End If
            End If
            '==== fin porcentajes ====

            rstpago_detalle!Deducciones = 1     'Val(TxtDeducciones.Text)
            rstpago_detalle!saldo_bolivianos = CDbl(rstao_solicitud_detalle!monto_bolivianos)
            rstpago_detalle!tipo_cambio = CDbl(rstao_solicitud_detalle!tipo_cambio)
            rstpago_detalle!tipo_cambio_dev = CDbl(rstao_solicitud_detalle!tipo_cambio)
            rstpago_detalle!estado_aprobacion = "N"
            rstpago_detalle!fecha_pago = Format(Date, "DD/MM/YYYY")  ', "dd/mm/aaaa
            rstpago_detalle!fecha_registro = Format(Date, "DD/MM/YYYY")
            rstpago_detalle!usr_usuario = GlUsuario
            rstpago_detalle!hora_registro = Format(Time, "hh:mm:ss")
            rstpago_detalle.Update
            '======== fin graba pago_detalle
          Next          'del segundo i      Para nor. comprobantes por cada "pagos_espera" o cada "pagos"
          'k = 1
          'k = k + 1
          If k = 1 Then
'            Call contabPCE(adosolicitud.Recordset, GlNombFor)
          End If
         Next           'del k          Para cambiar de "pagos_espera" a "pagos"
         'End
         Set rstdestino = New ADODB.Recordset
         If rstdestino.State = 1 Then rstdestino.Close
         rstdestino.Open "select * from ao_solicitud where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
         If rstdestino.RecordCount > 0 Then
            rstdestino!estado_enviado = "S"
            rstdestino.Update
            rstao_solicitud_recibido.AddNew
            rstao_solicitud_recibido!ges_gestion = IIf(IsNull(adoorigen!ges_gestion), "", adoorigen!ges_gestion)
            rstao_solicitud_recibido!codigo_solicitud = IIf(IsNull(adoorigen!codigo_solicitud), CStr(0), CStr(adoorigen!codigo_solicitud))
            rstao_solicitud_recibido!formulario = IIf(IsNull(adoorigen!formulario), "", adoorigen!formulario)
            rstao_solicitud_recibido!codigo_unidad = IIf(IsNull(adoorigen!codigo_unidad), "", adoorigen!codigo_unidad)
            rstao_solicitud_recibido!justificacion_solicitud = IIf(IsNull(adoorigen!justificacion_solicitud), "", adoorigen!justificacion_solicitud)
            rstao_solicitud_recibido!fecha_solicitud = Format(adoorigen!fecha_solicitud, "dd/mm/yyyy")
            rstao_solicitud_recibido!swSubir = swSubir
            rstao_solicitud_recibido!usr_usuario = GlUsuario
            rstao_solicitud_recibido!fecha_registro = Format(Date, "dd/mm/yyyy")
            rstao_solicitud_recibido!hora_registro = Format(Time, "hh:mm:ss")
            rstao_solicitud_recibido.Update
         End If
         If rstdestino.State = 1 Then rstdestino.Close
         rstao_solicitud_detalle.MoveNext
        Next        'del j
        db.CommitTrans
        If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
      End If
    End If  'fin swpresup
  End If
End Sub

Private Sub contabPCE(adoorigen, GlNombFor)
  '======== tipo de formualrio F01 ========
  If GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "A") Then
    tot_reg = 0
    Dim rstdetalle As New ADODB.Recordset
    Set rstdetalle = New ADODB.Recordset
    If rstdetalle.State = 1 Then rstdetalle.Close
    rstdetalle.Open "select * from ao_Solicitud_detalle where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockReadOnly
    If rstdetalle.RecordCount < 1 Then
      MsgBox "No se puede generar el asiento contable," & vbCrLf & "debido a que el registro no tiene el detalle de montos.", vbOKOnly + vbCritical, "Error al generar el asiento contabl..."
      If rstdetalle.State = 1 Then rstdetalle.Close
      Exit Sub
    Else
      tot_reg = 0
      If rstdetalle!monto_bolivianos > 0 Then tot_reg = tot_reg + 1
      If rstdetalle!monto_bolivianos_contra > 0 Then tot_reg = tot_reg + 1
    End If

    Set rstao_solicitud_recibido = New ADODB.Recordset
    If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
    rstao_solicitud_recibido.Open "SELECT * FROM ao_solicitud_recibido", db, adOpenKeyset, adLockOptimistic
    'db.BeginTrans
    '======== ini registro de co_comprobante_M ========
    Dim rstCodComp As New ADODB.Recordset
    Set rstdestino = New ADODB.Recordset
    For i = 1 To 1  'tot_reg
      If rstdetalle!monto_bolivianos <= 0 And i = 1 Then
        GoTo etiq
      End If
      If rstdetalle!monto_bolivianos_contra <= 0 And i = 2 Then
        GoTo etiq
      End If
      '======== ini GENERA EL CODIGO DE COMPROBANTE ========
      Set rstCodComp = New ADODB.Recordset
      rstCodComp.CursorLocation = adUseClient
      If rstCodComp.State = 1 Then rstCodComp.Close
      rstCodComp.Open "select * from fc_Correl  where tipo_tramite = 'cmbte'", db, adOpenDynamic, adLockOptimistic
      If rstCodComp.RecordCount > 0 Then
        Cont_Comp = Val(rstCodComp!numero_correlativo)
        Cont_Comp = Cont_Comp + 1
        rstCodComp!numero_correlativo = Trim(Str(Cont_Comp))
        rstCodComp.Update
      End If
      If rstCodComp.State = 1 Then rstCodComp.Close
      '======== fin TERMINA GENERACION DE COMPROBANTE ========

      '======== ini registro co_comprobantre_m ========
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from co_comprobante_m where Cod_Comp = 0", db, adOpenKeyset, adLockOptimistic
      If rstdestino.RecordCount > 0 Then
      End If
      rstdestino.AddNew
      rstdestino!Cod_Comp = Cont_Comp
      rstdestino!cod_trans = "0"
      If i = 1 Then
        rstdestino!org_codigo = "999" 'adoorigen!org_codigo_ext
      End If
      If i = 2 Then
        rstdestino!org_codigo = "999"  'rstdestino!org_codigo = adoorigen!org_codigo_contra
      End If
      rstdestino!cod_trans_detalle = 1
      rstdestino!num_respaldo = adoorigen!codigo_unidad & "/" & Str(adoorigen!codigo_solicitud)
      rstdestino!codigo_solicitud = (adoorigen!codigo_solicitud) 'adoorigen!codigo_unidad '& "/" & Str(adoorigen!codigo_solicitud)
      rstdestino!codigo_unidad = (adoorigen!codigo_unidad)
      rstdestino!fecha_A = Format(Date, "dd/mm/yyyy")         'Format(adoorigen!fecha_solicitud, "dd/mm/yyyy")
      rstdestino!codigo_beneficiario = adoorigen!CI_aprueba
      rstdestino!Origen = "1"
      'aqui fBuscaFteCorta(fte_1)
      If i = 1 Then
        rstdestino!glosa = Trim(adoorigen!justificacion_solicitud) & " " & fBuscaOrgCorta(rstdetalle!org_codigo_ext) & ": " & Round((rstdetalle!monto_bolivianos * 100 / (rstdetalle!monto_bolivianos + rstdetalle!monto_bolivianos_contra)), 2) & "%"
      End If
      If i = 2 Then
        rstdestino!glosa = Trim(adoorigen!justificacion_solicitud) & " " & fBuscaOrgCorta(rstdetalle!org_codigo_contra) & ": " & Round((rstdetalle!monto_bolivianos_contra * 100 / (rstdetalle!monto_bolivianos + rstdetalle!monto_bolivianos_contra)), 2) & "%"
      End If
      rstdestino!Status = "S"
      rstdestino!ges_gestion = adoorigen!ges_gestion
      rstdestino!codigo_documento = "D13"
      rstdestino!tipo_comp = "PCE" 'IIf(adoorigen!codigo_tipo = "DEV", "CAD", IIf(adoorigen!codigo_tipo = "REC", "CAR", v_Tipo_Comp(i)))
      '        rstdestino!tipo_moneda = adoorigen!tipo_moneda
      rstdestino!usr_usuario = GlUsuario
      rstdestino!fecha_registro = Date
      rstdestino!hora_registro = Format(Time, "hh:mm:ss")
      rstdestino!tipo_moneda = rstdetalle!tipo_moneda
      rstdestino.Update
      '======== fin registro co_comprobantre_m ========
      '======== ini registra CO_diaRIO ========
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from co_diario where Cod_Comp = " & Cont_Comp, db, adOpenKeyset, adLockOptimistic
      If rstdestino.RecordCount > 0 Then
        rstdestino.MoveFirst
      Else
        rstdestino.AddNew
        rstdestino!Cod_Comp = Cont_Comp
      End If

      rstdestino!tipo_comp = "PCE"
      rstdestino!d_cuenta = "1127"
  'y        rstdestino!D_Nombre = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
      rstdestino!d_subcta1 = "02"
      Select Case adoorigen!subcta2
        Case "01" '"Regulares" 'Cargos de Cuenta Regulares
          rstdestino!d_subcta2 = "01"
          rstdestino!d_Aux3 = "00"
        Case "02" '"Otros" 'Cargos de Cuenta Otros
          rstdestino!d_subcta2 = "02"
          rstdestino!d_Aux3 = "00"
        Case "03"  '"PASE" 'Cargos de Cuenta PASE
          rstdestino!d_subcta2 = "03"
          rstdestino!d_Aux3 = "10"
      End Select
      rstdestino!d_Aux1 = "01"
      rstdestino!d_Aux2 = "09"
'      rstdestino!d_Aux3 = "00"
      rstdestino!d_cta_larga = adoorigen!CI_aprueba
      rstdestino!D_Des_Larga = "-" ' CAMPO PARA ELIMINAR
      If i = 1 Then
        rstdestino!d_montoBs = rstdetalle!monto_bolivianos
        rstdestino!d_montoDl = rstdetalle!monto_dolares
        rstdestino!d_ctaaux2 = rstdetalle!org_codigo_ext   'GABY
      End If
      If i = 2 Then
        rstdestino!d_montoBs = rstdetalle!monto_bolivianos_contra
        rstdestino!d_montoDl = rstdetalle!monto_dolares_contra
        rstdestino!d_ctaaux2 = rstdetalle!org_codigo_contra  'GABY
      End If
      rstdestino!d_Cambio = rstdetalle!tipo_cambio
      rstdestino!h_cuenta = "2116"
  'Y        rstdestino!H_Nombre = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
      rstdestino!h_subcta1 = "02"
      rstdestino!h_subcta2 = "00"
      rstdestino!h_Aux1 = "01"
      rstdestino!h_Aux2 = "09"   'Y
      rstdestino!h_Aux3 = "00"
      rstdestino!h_cta_larga = adoorigen!CI_aprueba
      rstdestino!H_Des_Larga = "-"   ' CAMPO PARA ELIMINAR
      If i = 1 Then
        rstdestino!h_montoBs = rstdetalle!monto_bolivianos
        rstdestino!h_montoDl = rstdetalle!monto_dolares
        rstdestino!h_ctaaux2 = rstdetalle!codigo_convenio
        rstdestino!d_ctaaux2 = rstdetalle!codigo_convenio
        rstdestino!d_CtaAux3 = IIf(IsNull(rstdetalle!aux3), "", rstdetalle!aux3)

        'rsCo_diario!d_Aux3 = "10"
        'rsCo_diario!d_ctaaux3 = DtCCodigo.Text

      End If
      If i = 2 Then
        rstdestino!h_montoBs = rstdetalle!monto_bolivianos_contra
        rstdestino!h_montoDl = rstdetalle!monto_dolares_contra
        rstdestino!h_ctaaux2 = "FIN_PROPIO" 'rstdetalle!codigo_convenio
        rstdestino!d_ctaaux2 = "FIN_PROPIO" 'rstdetalle!codigo_convenio
        rstdestino!d_CtaAux3 = IIf(IsNull(rstdetalle!aux3), "", rstdetalle!aux3)
      End If
      rstdestino!h_Cambio = rstdetalle!tipo_cambio
      'grabar convenios
      'en h_ctaaux2 y en d_ctaaux2
'      rstdestino!h_ctaaux2 = rstdetalle!codigo_convenio
'      rstdestino!d_ctaaux2 = rstdetalle!codigo_convenio

      rstdestino!usr_usuario = GlUsuario
      rstdestino!fecha_registro = Date
      rstdestino!hora_registro = Format(Time, "hh:mm:ss")
      rstdestino.Update
      If rstdestino.State = 1 Then rstdestino.Close
      '======== fin registra co_diario ========
etiq:
    Next i
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from ao_solicitud where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
'    If rstdestino.RecordCount > 0 Then
'      rstdestino!estado_enviado = "S"
'      rstdestino.Update
'      rstao_solicitud_recibido.AddNew
'      rstao_solicitud_recibido!ges_gestion = IIf(IsNull(adoorigen!ges_gestion), "", adoorigen!ges_gestion)
'      rstao_solicitud_recibido!codigo_solicitud = IIf(IsNull(adoorigen!codigo_solicitud), 0, adoorigen!codigo_solicitud)
'      rstao_solicitud_recibido!formulario = IIf(IsNull(adoorigen!formulario), "", adoorigen!formulario)
'      rstao_solicitud_recibido!codigo_unidad = IIf(IsNull(adoorigen!codigo_unidad), "", adoorigen!codigo_unidad)
'      rstao_solicitud_recibido!justificacion_solicitud = IIf(IsNull(adoorigen!justificacion_solicitud), "", adoorigen!justificacion_solicitud)
'      'rstao_solicitud_recibido!swSubir = swSubir
'      rstao_solicitud_recibido!usr_usuario = GlUsuario
'      rstao_solicitud_recibido!fecha_registro = Format(Date, "dd/mm/yyyy")
'      rstao_solicitud_recibido!hora_registro = Format(Time, "hh:mm:ss")
'      rstao_solicitud_recibido.Update
'    End If
'    If rstdestino.State = 1 Then rstdestino.Close
    If rstdetalle.State = 1 Then rstdetalle.Close
    'db.CommitTrans
'    If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
  End If
'  '---- fin formulario f01 ----
End Sub

Private Sub GRABADET()
    Dim rstcodigo_detalle As New ADODB.Recordset
    Set rstcodigo_detalle = New ADODB.Recordset
    If rstcodigo_detalle.State = 1 Then rstcodigo_detalle.Close
    rstcodigo_detalle.Open "select sum(ao_solicitud_LISTA.total_venta) as monto_sol_bs from ao_solicitud_LISTA where codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
    'ADICIONAR PARA MAS PARTIDAS POR GRUPOS PRODUCTOS (CODIGO POA)
    'rstcodigo_detalle.Open "select sum(ao_solicitud_LISTA.monto_solicitud_dl) as monto_sol_bs from ao_solicitud_LISTA where codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
    'rstcodigo_detalle.Open "select sum(monto_solicitud_dl) as monto_sol_bs from ao_solicitud_LISTA where codigo_unidad = '" & lblcodigo_unidad & "' and codigo_solicitud = " & lblcodigo_solicitud, db, adOpenKeyset, adLockOptimistic
    If rstcodigo_detalle.RecordCount > 0 Then
    'db.Execute "select sum(monto_solicitud_dl) as monto_sol_bs from ao_solicitud_LISTA where codigo_unidad = '" & lblcodigo_unidad & "' and codigo_solicitud = " & lblcodigo_solicitud & " "
    Set rsdetalle = New ADODB.Recordset
    If rsdetalle.State = 1 Then rsdetalle.Close
    db.BeginTrans
      rsdetalle.Open "select * from ao_solicitud_detalle where codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
      If rsdetalle.RecordCount > 0 Then
        db.Execute "Update ao_solicitud_detalle Set monto_bolivianos= " & rstcodigo_detalle!monto_sol_bs & ", fecha_registro = '" & Format(Date, "dd/mm/yyyy") & "' , monto_DOLARES= " & (rstcodigo_detalle!monto_sol_bs / GlTipoCambioOficial) & ", ESTADO_APROBACION='S' Where codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " "
      Else
        db.Execute "insert into ao_solicitud_detalle(ges_gestion, codigo_unidad, codigo_solicitud, codigo_detalle, codigo_poa, por_fte_ext, por_fte_nal, Tipo_cambio, monto_bolivianos, monto_DOLARES, org_codigo_contra, org_codigo_EXT, por_fte_nal, monto_bolivianos_contra, monto_dolares_contra, tipo_moneda, codigo_convenio, aux3, formulario, usr_usuario, fecha_registro, hora_registro)" & _
                  "values ('" & gestion1 & "','" & adosolicitud.Recordset!codigo_unidad & "','" & adosolicitud.Recordset!codigo_solicitud & "', 1, '" & adosolicitud.Recordset!codigo_poa & "', 100, 0, " & GlTipoCambioOficial & ", " & rstcodigo_detalle!monto_sol_bs & ", " & (rstcodigo_detalle!monto_sol_bs / GlTipoCambioOficial) & ", '111', '111', 0, 0, 0, 'Bs', 'FIN_PROPIO', 'FIN_PROPIO', 'F04', '" & GlUsuario & "', '" & Format(Date, "dd/mm/yyyy") & "', '09:00' "
      End If
    db.CommitTrans
    End If
'  rsdetalle.Requery
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    Call GRABADET
'  '  Call SalePantalla
'  sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
'  If sino = vbYes Then
'    Dim rstAo_solicitud As New ADODB.Recordset
'    Set rstAo_solicitud = New ADODB.Recordset
'    If rstAo_solicitud.State = 1 Then rstAo_solicitud.Close
'    rstAo_solicitud.Open "select * from ao_solicitud where ges_gestion = '" & Trim(lblges_gestion) & "' and codigo_unidad = '" & Trim(lblcodigo_unidad) & "' and codigo_solicitud = " & lblcodigo_solicitud, db, adOpenKeyset, adLockOptimistic
'    If rstAo_solicitud.RecordCount > 0 Then
'      If rstAo_solicitud.RecordCount > 0 Then
'        rstAo_solicitud!Lista_adjunta = "S"
'      Else
'        rstAo_solicitud!Lista_adjunta = "N"
'      End If
'      rstAo_solicitud.Update
'    End If
'    If rstAo_solicitud.State = 1 Then rstAo_solicitud.Close
''    If rstAc_departamentos.State = 1 Then rstAc_departamentos.Close
'    If rstao_solicitud_lista.State = 1 Then rstao_solicitud_lista.Close
'    If rstdestino.State = 1 Then rstdestino.Close
'    Unload Me
'  End If
'
'End Sub


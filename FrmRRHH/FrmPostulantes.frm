VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmPostulantes 
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   14865
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport crCurriculum 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   1
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton btnImprimirCV 
      Caption         =   "Imprimir CV"
      Height          =   615
      Left            =   12840
      TabIndex        =   27
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame frameFormacion 
      Caption         =   "FORMACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6960
      TabIndex        =   25
      Top             =   5640
      Width           =   7335
      Begin MSDataGridLib.DataGrid dgFormacion 
         Bindings        =   "FrmPostulantes.frx":0000
         Height          =   1095
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1931
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "CarreraOCurso"
            Caption         =   "Estudio"
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
            DataField       =   "CentroEducativo"
            Caption         =   "Institucion"
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
            DataField       =   "TituloObtenido"
            Caption         =   "Titulo"
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
            DataField       =   "tipoDescripcion"
            Caption         =   "Educacion"
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
            DataField       =   "pais_descripcion"
            Caption         =   "Pais"
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
            DataField       =   "FechaInicio"
            Caption         =   "Inicio"
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
            DataField       =   "FechaFinalizacion"
            Caption         =   "Fin"
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
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoFormacion 
         Height          =   330
         Left            =   240
         Top             =   1440
         Width           =   6855
         _ExtentX        =   12091
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
   Begin VB.Frame frameExperiencia 
      Caption         =   "EXPERIENCIA LABORAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   6960
      TabIndex        =   23
      Top             =   3480
      Width           =   7335
      Begin MSDataGridLib.DataGrid dgExperiencia 
         Bindings        =   "FrmPostulantes.frx":001B
         Height          =   1215
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2143
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Cargo"
            Caption         =   "Cargo"
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
            DataField       =   "FuncionGeneral"
            Caption         =   "Funcion"
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
            DataField       =   "NombreInstitucion"
            Caption         =   "Empresa"
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
            DataField       =   "pais_descripcion"
            Caption         =   "Pais"
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
            DataField       =   "FechaInicio"
            Caption         =   "Inicio"
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
            DataField       =   "FechaFinalizacion"
            Caption         =   "Fin"
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
            BeginProperty Column03 
               ColumnWidth     =   1305,071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1005,165
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoExperienca 
         Height          =   330
         Left            =   240
         Top             =   1560
         Width           =   6855
         _ExtentX        =   12091
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
   Begin VB.Frame frameDatos 
      Caption         =   "DATOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   6960
      TabIndex        =   2
      Top             =   840
      Width           =   7335
      Begin VB.Image imgPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1700
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1700
      End
      Begin VB.Label lblNacionPostulante 
         Caption         =   "Bolivia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   22
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblNacion 
         Caption         =   "Nacionalidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblGeneroPostulante 
         Caption         =   "Masculino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblGenero 
         Caption         =   "Genero:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblDireccionPostulante 
         Caption         =   "Av. Camacho Esq. Litoral Nro. 288"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   2160
         Width           =   5775
      End
      Begin VB.Label lblDireccion 
         Caption         =   "Direccion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblFechaNacPostulante 
         Caption         =   "21-04-2002"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   16
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblFechaNac 
         Caption         =   "Fecha de nacimiento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label lblEstadoCivilPostulante 
         Caption         =   "Soltero(a)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblEstadoCivil 
         Caption         =   "Estado civil:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblEmailPostulante 
         Caption         =   "postulante@gmail.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblCiPostulante 
         Caption         =   "13968194 LP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblCi 
         Caption         =   "CI:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblTelefonoPostulante 
         Caption         =   "2226457"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   8
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblTelefono 
         Caption         =   "Tel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblCelularPostulante 
         Caption         =   "73577592"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblCelular 
         Caption         =   "Celular:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblNombrePostulante 
         Caption         =   "Halkyer Camacho Estela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc adoListaPostuantes 
      Height          =   330
      Left            =   240
      Top             =   7680
      Width           =   6495
      _ExtentX        =   11456
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
   Begin MSDataGridLib.DataGrid dgListaPostulantes 
      Bindings        =   "FrmPostulantes.frx":0037
      Height          =   6735
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   11880
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "PrimerApellido"
         Caption         =   "Paterno"
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
         DataField       =   "SegundoApellido"
         Caption         =   "Materno"
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
         DataField       =   "Nombres"
         Caption         =   "Nombres"
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
         DataField       =   "puestoDescripcion"
         Caption         =   "Puesto postulado"
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
         DataField       =   "FechaRegistro"
         Caption         =   "Registro"
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
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1590,236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   959,811
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitulo 
      Caption         =   "POSTULANTES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "FrmPostulantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsListaPostulantes As New ADODB.Recordset
Dim rsListaExperiencia As New ADODB.Recordset
Dim rsListaFormacion As New ADODB.Recordset

Private Sub adoListaPostuantes_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Not rsListaPostulantes.BOF And Not rsListaPostulantes.EOF Then
        lblNombrePostulante.Caption = adoListaPostuantes.Recordset!PrimerApellido & " " & adoListaPostuantes.Recordset!SegundoApellido
        lblNombrePostulante.Caption = lblNombrePostulante.Caption & " " & adoListaPostuantes.Recordset!Nombres
        lblCiPostulante.Caption = adoListaPostuantes.Recordset!DocumentoIdentidad
        lblCelularPostulante.Caption = adoListaPostuantes.Recordset!TelefonoCelular
        lblTelefonoPostulante.Caption = adoListaPostuantes.Recordset!TelefonoFijo
        lblEmailPostulante.Caption = adoListaPostuantes.Recordset!EmailPersonal
        lblEstadoCivilPostulante.Caption = adoListaPostuantes.Recordset!estado_civil_descripcion
        lblGeneroPostulante.Caption = adoListaPostuantes.Recordset!generoDescripcion
        lblFechaNacPostulante.Caption = adoListaPostuantes.Recordset!FechaNacimiento
        lblNacionPostulante.Caption = adoListaPostuantes.Recordset!pais_descripcion
        lblDireccionPostulante.Caption = adoListaPostuantes.Recordset!DomicilioLegal
        
        Call leerExperiencia(adoListaPostuantes.Recordset!idBeneficiario)
        Call leerFormacion(adoListaPostuantes.Recordset!idBeneficiario)
    End If
End Sub

Private Sub btnImprimirCV_Click()
    db.Execute "UPDATE rrhh.Beneficiario SET Denominacion = PrimerApellido + ' ' + SegundoApellido + ' ' + Nombres"

    Dim result As Integer
    '------------------SEGUNDO REPORTE--------------
    crCurriculum.Reset
    crCurriculum.WindowState = crptMaximized
    crCurriculum.WindowShowSearchBtn = True
    crCurriculum.WindowShowRefreshBtn = True
    crCurriculum.WindowShowPrintSetupBtn = True
    crCurriculum.ReportFileName = App.Path & "\Reportes\RRHH\rrPostulanteExperiencia.rpt"
    crCurriculum.StoredProcParam(0) = adoListaPostuantes.Recordset!idBeneficiario
    result = crCurriculum.PrintReport
    If result <> 0 Then
        MsgBox crCurriculum.LastErrorNumber & " : " & crCurriculum.LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
    '------------------PRIMER REPORTE--------------
    crCurriculum.Reset
    crCurriculum.WindowState = crptMaximized
    crCurriculum.WindowShowSearchBtn = True
    crCurriculum.WindowShowRefreshBtn = True
    crCurriculum.WindowShowPrintSetupBtn = True
    crCurriculum.ReportFileName = App.Path & "\Reportes\RRHH\rrPostulanteCurriculum.rpt"
    crCurriculum.StoredProcParam(0) = adoListaPostuantes.Recordset!idBeneficiario
    result = crCurriculum.PrintReport
    If result <> 0 Then
        MsgBox crCurriculum.LastErrorNumber & " : " & crCurriculum.LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
End Sub

Private Sub Form_Load()
    Call leerPostulantes
End Sub

Private Sub leerPostulantes()
    Set rsListaPostulantes = New ADODB.Recordset
    If rsListaPostulantes.State = 1 Then rsListaPostulantes.Close
    rsListaPostulantes.Open "SELECT * FROM rvPostulantes ORDER BY FechaRegistro DESC", db, adOpenStatic
    Set adoListaPostuantes.Recordset = rsListaPostulantes
End Sub

Private Sub leerExperiencia(beneficiarioId As Integer)
    Set rsListaExperiencia = New ADODB.Recordset
    If rsListaExperiencia.State = 1 Then rsListaExperiencia.Close
    rsListaExperiencia.Open "SELECT * FROM rvPostulanteExperiencia WHERE BeneficiarioId = " & beneficiarioId & " ORDER BY FechaInicio DESC", db, adOpenStatic
    Set adoExperienca.Recordset = rsListaExperiencia
End Sub

Private Sub leerFormacion(beneficiarioId As Integer)
    Set rsListaFormacion = New ADODB.Recordset
    If rsListaFormacion.State = 1 Then rsListaFormacion.Close
    rsListaFormacion.Open "SELECT * FROM rvPostulanteFormacion WHERE BeneficiarioId = " & beneficiarioId & " ORDER BY tipoDescripcion DESC, FechaInicio DESC", db, adOpenStatic
    Set adoFormacion.Recordset = rsListaFormacion
End Sub

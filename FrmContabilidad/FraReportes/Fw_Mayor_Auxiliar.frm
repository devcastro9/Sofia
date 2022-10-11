VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Fw_Mayor_Auxiliar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Clasificadores - Gerencia General"
   ClientHeight    =   11325
   ClientLeft      =   990
   ClientTop       =   -105
   ClientWidth     =   18990
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11325
   ScaleWidth      =   18990
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame Fra_Aux3 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   2055
      Left            =   11520
      TabIndex        =   18
      Top             =   2520
      Width           =   5775
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2520
         TabIndex        =   59
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1200
         TabIndex        =   27
         Top             =   480
         Width           =   375
      End
      Begin VB.PictureBox Buscar3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3840
         Picture         =   "Fw_Mayor_Auxiliar.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo txt_Aux3 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":07B5
         Height          =   315
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "Aux3"
         BoundColumn     =   "correl"
         Text            =   "0000"
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":07D0
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "desc3"
         BoundColumn     =   "codigo3"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":07EA
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "codigo3"
         BoundColumn     =   "codigo3"
         Text            =   "Todos"
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Auxiliar - 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2295
         TabIndex        =   23
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.Frame Fra_Aux2 
      BackColor       =   &H00C0C0C0&
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
      Height          =   2055
      Left            =   5880
      TabIndex        =   42
      Top             =   2520
      Width           =   5580
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2520
         TabIndex        =   58
         Top             =   960
         Width           =   375
      End
      Begin VB.PictureBox Buscar2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3600
         Picture         =   "Fw_Mayor_Auxiliar.frx":0804
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   54
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1185
         TabIndex        =   49
         Top             =   480
         Width           =   375
      End
      Begin MSDataListLib.DataCombo txt_Aux2 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":0FB9
         Height          =   315
         Left            =   225
         TabIndex        =   50
         Top             =   480
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "Aux2"
         BoundColumn     =   "correl"
         Text            =   "0000"
      End
      Begin MSDataListLib.DataCombo dtc_codigo9 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":0FD4
         Height          =   315
         Left            =   240
         TabIndex        =   52
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "codigo2"
         BoundColumn     =   "codigo2"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc9 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":0FED
         Height          =   315
         Left            =   240
         TabIndex        =   53
         Top             =   1440
         Visible         =   0   'False
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "desc2"
         BoundColumn     =   "codigo2"
         Text            =   "Todos"
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Auxiliar - 2"
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
         Height          =   240
         Left            =   2280
         TabIndex        =   51
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.Frame Fra_Aux1 
      BackColor       =   &H00C0C0C0&
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
      Height          =   2055
      Left            =   120
      TabIndex        =   41
      Top             =   2520
      Width           =   5700
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   57
         Top             =   960
         Width           =   375
      End
      Begin VB.PictureBox Buscar1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3720
         Picture         =   "Fw_Mayor_Auxiliar.frx":1006
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   48
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   43
         Top             =   480
         Width           =   375
      End
      Begin MSDataListLib.DataCombo txt_Aux1 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":17BB
         Height          =   315
         Left            =   240
         TabIndex        =   44
         Top             =   480
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "Aux1"
         BoundColumn     =   "correl"
         Text            =   "0000"
      End
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":17D6
         Height          =   315
         Left            =   240
         TabIndex        =   46
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "codigo1"
         BoundColumn     =   "codigo1"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":17EF
         Height          =   315
         Left            =   240
         TabIndex        =   47
         Top             =   1440
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "desc1"
         BoundColumn     =   "codigo1"
         Text            =   "Todos"
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Auxiliar - 1"
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
         Height          =   240
         Index           =   0
         Left            =   2280
         TabIndex        =   45
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.Frame Fra_op_aux 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Resultado de la Selección"
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
      Height          =   5775
      Left            =   120
      TabIndex        =   37
      Top             =   4680
      Width           =   17175
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":1808
         Height          =   5250
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   16920
         _ExtentX        =   29845
         _ExtentY        =   9260
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
            Weight          =   700
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
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "fecha"
            Caption         =   "FECHA"
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
            DataField       =   "TC"
            Caption         =   "T.C."
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
            DataField       =   "comp"
            Caption         =   "CMPBTE"
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
            DataField       =   "debe"
            Caption         =   "DEBE Bs."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "haber"
            Caption         =   "HABER Bs."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "MovSus"
            Caption         =   "Dolares"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "SaldoBs"
            Caption         =   "Saldo.Bs."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "SaldoSus"
            Caption         =   "Saldo.Dol."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "glosa"
            Caption         =   "DETALLE"
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
            DataField       =   "org"
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
         BeginProperty Column10 
            DataField       =   "cte"
            Caption         =   "TRAMITE"
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
            DataField       =   "tipo"
            Caption         =   "TIPO"
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
         BeginProperty Column12 
            DataField       =   "AuxDes1"
            Caption         =   "Descripcion Aux1"
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
            DataField       =   "AuxDes2"
            Caption         =   "Descripcion Aux2"
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
            DataField       =   "AuxDes3"
            Caption         =   "Descripcion Aux3"
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
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Object.Visible         =   0   'False
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   8625.261
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   -1  'True
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1709.858
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1725.165
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1709.858
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton opt_aux1 
         BackColor       =   &H00000000&
         Caption         =   " AUXILIAR 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   390
         Left            =   120
         TabIndex        =   40
         Top             =   4440
         Width           =   1980
      End
      Begin VB.OptionButton opt_aux3 
         BackColor       =   &H00000000&
         Caption         =   " AUXILIAR 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   390
         Left            =   9600
         TabIndex        =   39
         Top             =   4440
         Width           =   1980
      End
      Begin VB.OptionButton opt_aux2 
         BackColor       =   &H00000000&
         Caption         =   " AUXILIAR 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   390
         Left            =   4920
         TabIndex        =   38
         Top             =   4440
         Width           =   1620
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   735
      Left            =   120
      TabIndex        =   31
      Top             =   720
      Width           =   17175
      Begin VB.OptionButton opt_bs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bolivianos"
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
         Height          =   390
         Left            =   12840
         TabIndex        =   55
         Top             =   240
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.OptionButton opt_dol 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dólares"
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
         Height          =   390
         Left            =   14880
         TabIndex        =   36
         Top             =   240
         Visible         =   0   'False
         Width           =   1260
      End
      Begin MSComCtl2.DTPicker stp_fecha_inicio 
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
         Left            =   1680
         TabIndex        =   34
         Top             =   240
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   84017153
         CurrentDate     =   42736
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker stp_fecha_final 
         DataField       =   "beneficiario_fecha_nacimiento"
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
         Left            =   5640
         TabIndex        =   35
         Top             =   240
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   84017153
         CurrentDate     =   43070
         MinDate         =   2
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
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
         Left            =   480
         TabIndex        =   33
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final"
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
         Left            =   4440
         TabIndex        =   32
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.Frame Fra_ABM1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buscar por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   17175
      Begin VB.TextBox txt_CtaHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   65
         Top             =   120
         Visible         =   0   'False
         Width           =   5415
      End
      Begin MSDataListLib.DataCombo txt_Cuenta_tot 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":1820
         Height          =   315
         Left            =   480
         TabIndex        =   64
         Top             =   480
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "CuentaAux"
         BoundColumn     =   "correl"
         Text            =   "0000"
      End
      Begin MSDataListLib.DataCombo txt_cuenta_des 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":183B
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   8400
         TabIndex        =   63
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "NombreCtaAux"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   8070
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         Height          =   315
         Left            =   1200
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSDataListLib.DataCombo txt_NombreCta 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":1856
         Height          =   315
         Left            =   2280
         TabIndex        =   10
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "NombreCta"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo txt_SubCta2 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":1871
         Height          =   315
         Left            =   2520
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "subcta2"
         BoundColumn     =   "correl"
         Text            =   "0000"
      End
      Begin MSDataListLib.DataCombo txt_SubCta1 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":188C
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "subcta1"
         BoundColumn     =   "correl"
         Text            =   "0000"
      End
      Begin MSDataListLib.DataCombo txt_Cuenta 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":18A7
         Height          =   315
         Left            =   480
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "Cuenta"
         BoundColumn     =   "correl"
         Text            =   "0000"
      End
      Begin MSDataListLib.DataCombo txt_correl 
         Bindings        =   "Fw_Mayor_Auxiliar.frx":18C2
         Height          =   315
         Left            =   0
         TabIndex        =   56
         Top             =   480
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483631
         ListField       =   "correl"
         BoundColumn     =   "correl"
         Text            =   "0000"
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre de la Cuenta"
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
         Left            =   2280
         TabIndex        =   16
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo y Nombre de la Cuenta"
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
         Left            =   8400
         TabIndex        =   15
         Top             =   240
         Width           =   2760
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cuenta"
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
         Left            =   480
         TabIndex        =   14
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.PictureBox Fra_ABM 
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
      TabIndex        =   6
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnMigrar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
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
         ForeColor       =   &H00FFFF00&
         Height          =   615
         Left            =   5160
         Picture         =   "Fw_Mayor_Auxiliar.frx":18DD
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   66
         ToolTipText     =   "Imprime Mayor"
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2040
         Picture         =   "Fw_Mayor_Auxiliar.frx":23B8
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   62
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   600
         Picture         =   "Fw_Mayor_Auxiliar.frx":2B6D
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   60
         ToolTipText     =   "Carga en Resultado de la Selección"
         Top             =   0
         Width           =   1335
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
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
         ForeColor       =   &H00FFFF00&
         Height          =   615
         Left            =   15960
         Picture         =   "Fw_Mayor_Auxiliar.frx":335B
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   17
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
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
         ForeColor       =   &H00FFFF00&
         Height          =   615
         Left            =   3720
         Picture         =   "Fw_Mayor_Auxiliar.frx":3B1D
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   8
         ToolTipText     =   "Imprime Mayor"
         Top             =   0
         Width           =   1400
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MAYOR AUXILIAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   9315
         TabIndex        =   7
         Top             =   195
         Width           =   2085
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
      ScaleWidth      =   18990
      TabIndex        =   0
      Top             =   11325
      Width           =   18990
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos 
      Height          =   330
      Left            =   240
      Top             =   10440
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
      Caption         =   "Ado_DE"
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
      Left            =   4680
      Top             =   10440
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
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
      Top             =   10440
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
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   120
      Top             =   10800
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
      Left            =   2400
      Top             =   10800
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
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   4680
      Top             =   10800
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSComCtl2.DTPicker DTP_Fecha1 
      DataField       =   "beneficiario_fecha_nacimiento"
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
      Left            =   1980
      TabIndex        =   28
      Top             =   -240
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   84017153
      CurrentDate     =   40179
      MinDate         =   2
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   7440
      Top             =   10560
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
   Begin Crystal.CrystalReport CryComp_Manual 
      Left            =   7920
      Top             =   10560
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Nombre de la Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   2205
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Fecha de Nacimiento"
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
      Left            =   0
      TabIndex        =   29
      Top             =   -240
      Width           =   1920
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Correlativo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1155
   End
End
Attribute VB_Name = "Fw_Mayor_Auxiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
Dim rs_datos1 As New ADODB.Recordset
Dim rs_detalle1 As New ADODB.Recordset
Dim rs_detalle2 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset

'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String
Dim VAR_VAL As String
Dim VAR_SW As String

Dim VAR_TABLA, VAR_CODIGO, VAR_DES As String
Dim VAR_AUX1, VAR_AUX2, VAR_AUX3 As String
Dim VAR_FECHA1, VAR_FECHA2 As String
Dim VAR_AUXD1, VAR_AUXD2, VAR_AUXD3 As String

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

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

Private Sub valida_campos()
  If opt_bs.Value = False Or opt_dol.Value = False Then
    MsgBox "Debe Seleccionar Bolivianos o Dolares " + opt_bs.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
   Exit Sub
  End If
   If opt_aux1.Value = False Or opt_aux2.Value = False Or opt_aux3.Value = False Then
    MsgBox "Debe Seleccionar un Auxiliar " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnGrabar_Click()
    If txt_Cuenta.Text = "" Then
        MsgBox "Debe Seleccionar una Cuenta para obtener un resultado, Vuelva a intentar ...", vbCritical + vbExclamation, "Validación de datos"
        Exit Sub
    Else
        Fra_op_aux.Visible = True
        BtnBuscar.Visible = True
        VAR_FECHA1 = Str(stp_fecha_inicio.Value)
        VAR_FECHA2 = Str(stp_fecha_final.Value)
        If dtc_codigo8.Text = "" Then
            VAR_AUXD1 = "%"
        Else
            VAR_AUXD1 = dtc_codigo8
        End If
        If dtc_codigo9.Text = "" Then
            VAR_AUXD2 = "%"
        Else
            VAR_AUXD2 = dtc_codigo9
        End If
        If dtc_codigo10.Text = "" Then
            VAR_AUXD3 = "%"
        Else
            VAR_AUXD3 = dtc_codigo10
        End If
        'db.Execute "EXEC cp_LMayorAux1_2_3  VAR_FECHA1, VAR_FECHA2, txt_Cuenta, txt_SubCta1, txt_SubCta2, VAR_AUXD1, VAR_AUXD2, VAR_AUXD3, txt_Aux1, txt_Aux2, txt_Aux3"
        db.Execute "EXEC cp_LMayorAux  '" & VAR_FECHA1 & "', '" & VAR_FECHA2 & "', '" & txt_Cuenta & "', '" & txt_SubCta1 & "', '" & txt_SubCta2 & "', '" & VAR_AUXD1 & "', '" & VAR_AUXD2 & "', '" & VAR_AUXD3 & "', '" & txt_Aux1 & "', '" & txt_Aux2 & "', '" & txt_Aux3 & "' "
    End If
    '@FFInicio varchar(10), @FFFinal varchar(10), @cuenta  varchar  (5) ,@subcta1 varchar (3) ,@subcta2 varchar (3) ,@busca1 varchar(40),@busca2 varchar(40),@busca3 varchar(40),@aux1 varchar(3),@aux2 varchar(3),     @aux3 varchar(3)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "select * From cv_LMayorAux "
    'WHERE estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "'
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "Fecha, comp"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset

End Sub

Private Sub BtnImprimir_Click()
  If txt_Cuenta.Text = "" Then
    MsgBox "Debe Seleccionar una Cuenta para obtener un resultado, Vuelva a intentar ...", vbCritical + vbExclamation, "Validación de datos"
    Exit Sub
  Else
    cr01.Reset
'valida_campos
'If (Ado_detalle1.Recordset.RecordCount > 0) Then
'And (Ado_datos8.Recordset.RecordCount > 0)
'And (Ado_datos9.Recordset.RecordCount > 0) Or (Ado_datos10.Recordset.RecordCount > 0)
        Dim iResult As Integer
        If GlBaseDatos = "ADMIN_EMPRESA" Then
            cr01.ReportFileName = App.Path & "\REPORTES\Contabilidad\cr_mayor_auxiliar_bs.rpt"
        Else
            cr01.ReportFileName = App.Path & "\REPORTES\Contabilidad\cr_mayor_auxiliar_prueba.rpt"
        End If
        cr01.WindowState = crptMaximized
        cr01.WindowShowSearchBtn = True
        cr01.WindowShowPrintSetupBtn = True
        cr01.WindowShowRefreshBtn = True
        cr01.StoredProcParam(0) = "%"
        cr01.StoredProcParam(1) = "%"
        cr01.StoredProcParam(2) = "%"
        cr01.StoredProcParam(3) = "%"
        cr01.StoredProcParam(4) = "%"
        cr01.StoredProcParam(5) = "%"
        cr01.StoredProcParam(6) = "%"
        cr01.StoredProcParam(7) = "%"
        cr01.StoredProcParam(8) = "%"
        cr01.StoredProcParam(9) = "%"
        cr01.StoredProcParam(10) = "%"
        
        If stp_fecha_inicio.Value = "" Then     '@FFInicio varchar(10),
             cr01.StoredProcParam(0) = "%"
         Else
            cr01.StoredProcParam(0) = Format(stp_fecha_inicio.Value, "dd/mm/yyyy")
        End If
        If stp_fecha_final.Value = "" Then      '@FFFinal varchar(10) ,
             cr01.StoredProcParam(1) = "%"
         Else
            cr01.StoredProcParam(1) = Format(stp_fecha_final.Value, "dd/mm/yyyy")
        End If
        If txt_Cuenta.Text = "" Then            '@cuenta  varchar  (5) ,
            cr01.StoredProcParam(2) = "%"
         Else
            cr01.StoredProcParam(2) = Trim(txt_Cuenta.Text)
        End If
        
        If txt_SubCta1.Text = "" Then           '@subcta1 varchar (3) ,
             cr01.StoredProcParam(3) = "%"
         Else
            cr01.StoredProcParam(3) = Trim(txt_SubCta1.Text)
        End If
        
        If txt_SubCta2.Text = "" Then           '@subcta2 varchar (3) ,
             cr01.StoredProcParam(4) = "%"
         Else
            cr01.StoredProcParam(4) = Trim(txt_SubCta2.Text)
        End If
'
        If dtc_codigo8.Text = "" Then            '@busca1 varchar(40),
            cr01.StoredProcParam(5) = "%"
         Else
            cr01.StoredProcParam(5) = Trim(dtc_codigo8.Text)
        End If
        If dtc_codigo9.Text = "" Then           '@busca2 varchar(40),
            cr01.StoredProcParam(6) = "%"
         Else
            cr01.StoredProcParam(6) = Trim(dtc_codigo9.Text)
        End If
        If dtc_codigo10.Text = "" Then          '@busca3 varchar(40),
            cr01.StoredProcParam(7) = "%"
         Else
            cr01.StoredProcParam(7) = Trim(dtc_codigo10.Text)
        End If
        If txt_Aux1.Text = "" Then              '@aux1 varchar(3),
             cr01.StoredProcParam(8) = "%"
         Else
            cr01.StoredProcParam(8) = Trim(txt_Aux1.Text)
        End If
        If txt_Aux2.Text = "" Then              '@aux2 varchar(3),
             cr01.StoredProcParam(9) = "%"
         Else
            cr01.StoredProcParam(9) = Trim(txt_Aux2.Text)
        End If
        If txt_Aux3.Text = "" Then              '@aux3 varchar(3)
            cr01.StoredProcParam(10) = "%"
         Else
            cr01.StoredProcParam(10) = Trim(txt_Aux3.Text)
        End If

        cr01.Formulas(0) = "FFechaFinal = '" & Format(stp_fecha_final.Value, "dd/mm/yyyy") & "'"
        cr01.Formulas(1) = "FFechaInicio ='" & Format(stp_fecha_inicio.Value, "dd/mm/yyyy") & "'"
        cr01.Formulas(2) = "nomaux1 = '" & Trim(dtc_desc8.Text) & "' "
        cr01.Formulas(3) = "nomaux2 = '" & Trim(dtc_desc9.Text) & "' "
        cr01.Formulas(4) = "nomaux3 = '" & Trim(dtc_desc10.Text) & "' "
        cr01.Formulas(5) = "nomcta = '" & Trim(txt_NombreCta.Text) & "' "
        
'        CR01.StoredProcParam(2) = "%"
'        CR01.StoredProcParam(3) = "%"
'        CR01.StoredProcParam(4) = "%"
'        CR01.StoredProcParam(5) = "%"
'        CR01.StoredProcParam(6) = "%"
'        CR01.StoredProcParam(7) = "%"
'        CR01.StoredProcParam(8) = "%"
'        CR01.StoredProcParam(9) = "%"
'        CR01.StoredProcParam(10) = "%"
        cr01.Formulas(6) = "Vcuenta = '" & cr01.StoredProcParam(2) & "'"
        cr01.Formulas(7) = "VSub1 ='" & cr01.StoredProcParam(3) & "'"
        cr01.Formulas(8) = "VSub2 = '" & cr01.StoredProcParam(4) & "' "
        cr01.Formulas(9) = "Vbusca1 = '" & cr01.StoredProcParam(5) & "' "
        cr01.Formulas(10) = "Vbusca1 = '" & cr01.StoredProcParam(6) & "' "
        cr01.Formulas(11) = "Vbusca1 = '" & cr01.StoredProcParam(7) & "' "

'            If dtc_desc8.Text = "" Then
'         cr01.StoredProcParam(5) = "%"
'         Else
'          cr01.StoredProcParam(5) = dtc_desc8.Text
'        End If
'            If dtc_desc9.Text = "" Then
'         cr01.StoredProcParam(8) = "%"
'         Else
'          cr01.StoredProcParam(8) = dtc_desc9.Text
'        End If

'        If dtc_desc10.Text = "" Then
'         cr01.StoredProcParam(11) = "%"
'         Else
'          cr01.StoredProcParam(11) = dtc_desc10.Text
'        End If
        '
        '
        'nomaux2
        'nomaux3
        
        iResult = cr01.PrintReport
        If iResult <> 0 Then MsgBox cr01.LastErrorNumber & " : " & cr01.LastErrorString, vbCritical, "Error de impresión"
        cr01.WindowState = crptMaximized
 ' Else
 '   MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
  
End Sub

Private Sub BtnMigrar_Click()
  If txt_Cuenta.Text = "" Then
    MsgBox "Debe Seleccionar una Cuenta para obtener un resultado, Vuelva a intentar ...", vbCritical + vbExclamation, "Validación de datos"
    Exit Sub
  Else
    cr01.Reset
'valida_campos
'If (Ado_detalle1.Recordset.RecordCount > 0) Then
'And (Ado_datos8.Recordset.RecordCount > 0)
'And (Ado_datos9.Recordset.RecordCount > 0) Or (Ado_datos10.Recordset.RecordCount > 0)
        Dim iResult As Integer
        cr01.ReportFileName = App.Path & "\REPORTES\Contabilidad\cr_mayor_auxiliar_bs_migrar.rpt"
        cr01.WindowState = crptMaximized
        cr01.WindowShowSearchBtn = True
        cr01.WindowShowPrintSetupBtn = True
        cr01.WindowShowRefreshBtn = True
        cr01.StoredProcParam(0) = "%"
        cr01.StoredProcParam(1) = "%"
        cr01.StoredProcParam(2) = "%"
        cr01.StoredProcParam(3) = "%"
        cr01.StoredProcParam(4) = "%"
        cr01.StoredProcParam(5) = "%"
        cr01.StoredProcParam(6) = "%"
        cr01.StoredProcParam(7) = "%"
        cr01.StoredProcParam(8) = "%"
        cr01.StoredProcParam(9) = "%"
        cr01.StoredProcParam(10) = "%"
        
        If stp_fecha_inicio.Value = "" Then     '@FFInicio varchar(10),
             cr01.StoredProcParam(0) = "%"
         Else
            cr01.StoredProcParam(0) = Format(stp_fecha_inicio.Value, "dd/mm/yyyy")
        End If
        If stp_fecha_final.Value = "" Then      '@FFFinal varchar(10) ,
             cr01.StoredProcParam(1) = "%"
         Else
            cr01.StoredProcParam(1) = Format(stp_fecha_final.Value, "dd/mm/yyyy")
        End If
        If txt_Cuenta.Text = "" Then            '@cuenta  varchar  (5) ,
            cr01.StoredProcParam(2) = "%"
         Else
            cr01.StoredProcParam(2) = Trim(txt_Cuenta.Text)
        End If
        
        If txt_SubCta1.Text = "" Then           '@subcta1 varchar (3) ,
             cr01.StoredProcParam(3) = "%"
         Else
            cr01.StoredProcParam(3) = Trim(txt_SubCta1.Text)
        End If
        
        If txt_SubCta2.Text = "" Then           '@subcta2 varchar (3) ,
             cr01.StoredProcParam(4) = "%"
         Else
            cr01.StoredProcParam(4) = Trim(txt_SubCta2.Text)
        End If
'
        If dtc_codigo8.Text = "" Then            '@busca1 varchar(40),
            cr01.StoredProcParam(5) = "%"
         Else
            cr01.StoredProcParam(5) = Trim(dtc_codigo8.Text)
        End If
        If dtc_codigo9.Text = "" Then           '@busca2 varchar(40),
            cr01.StoredProcParam(6) = "%"
         Else
            cr01.StoredProcParam(6) = Trim(dtc_codigo9.Text)
        End If
        If dtc_codigo10.Text = "" Then          '@busca3 varchar(40),
            cr01.StoredProcParam(7) = "%"
         Else
            cr01.StoredProcParam(7) = Trim(dtc_codigo10.Text)
        End If
        If txt_Aux1.Text = "" Then              '@aux1 varchar(3),
             cr01.StoredProcParam(8) = "%"
         Else
            cr01.StoredProcParam(8) = Trim(txt_Aux1.Text)
        End If
        If txt_Aux2.Text = "" Then              '@aux2 varchar(3),
             cr01.StoredProcParam(9) = "%"
         Else
            cr01.StoredProcParam(9) = Trim(txt_Aux2.Text)
        End If
        If txt_Aux3.Text = "" Then              '@aux3 varchar(3)
            cr01.StoredProcParam(10) = "%"
         Else
            cr01.StoredProcParam(10) = Trim(txt_Aux3.Text)
        End If

        cr01.Formulas(0) = "FFechaFinal = '" & Format(stp_fecha_final.Value, "dd/mm/yyyy") & "'"
        cr01.Formulas(1) = "FFechaInicio ='" & Format(stp_fecha_inicio.Value, "dd/mm/yyyy") & "'"
        cr01.Formulas(2) = "nomaux1 = '" & Trim(dtc_desc8.Text) & "' "
        cr01.Formulas(3) = "nomaux2 = '" & Trim(dtc_desc9.Text) & "' "
        cr01.Formulas(4) = "nomaux3 = '" & Trim(dtc_desc10.Text) & "' "
        cr01.Formulas(5) = "nomcta = '" & Trim(txt_NombreCta.Text) & "' "
        
'        CR01.StoredProcParam(2) = "%"
'        CR01.StoredProcParam(3) = "%"
'        CR01.StoredProcParam(4) = "%"
'        CR01.StoredProcParam(5) = "%"
'        CR01.StoredProcParam(6) = "%"
'        CR01.StoredProcParam(7) = "%"
'        CR01.StoredProcParam(8) = "%"
'        CR01.StoredProcParam(9) = "%"
'        CR01.StoredProcParam(10) = "%"
        cr01.Formulas(6) = "Vcuenta = '" & cr01.StoredProcParam(2) & "'"
        cr01.Formulas(7) = "VSub1 ='" & cr01.StoredProcParam(3) & "'"
        cr01.Formulas(8) = "VSub2 = '" & cr01.StoredProcParam(4) & "' "
        cr01.Formulas(9) = "Vbusca1 = '" & cr01.StoredProcParam(5) & "' "
        cr01.Formulas(10) = "Vbusca1 = '" & cr01.StoredProcParam(6) & "' "
        cr01.Formulas(11) = "Vbusca1 = '" & cr01.StoredProcParam(7) & "' "

'            If dtc_desc8.Text = "" Then
'         cr01.StoredProcParam(5) = "%"
'         Else
'          cr01.StoredProcParam(5) = dtc_desc8.Text
'        End If
'            If dtc_desc9.Text = "" Then
'         cr01.StoredProcParam(8) = "%"
'         Else
'          cr01.StoredProcParam(8) = dtc_desc9.Text
'        End If

'        If dtc_desc10.Text = "" Then
'         cr01.StoredProcParam(11) = "%"
'         Else
'          cr01.StoredProcParam(11) = dtc_desc10.Text
'        End If
        '
        '
        'nomaux2
        'nomaux3
        
        iResult = cr01.PrintReport
        If iResult <> 0 Then MsgBox cr01.LastErrorNumber & " : " & cr01.LastErrorString, vbCritical, "Error de impresión"
        cr01.WindowState = crptMaximized
 ' Else
 '   MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If

End Sub

Private Sub BtnSalir_Click()
  Unload Me
End Sub

Private Sub limpiar()
 dtc_codigo8.Visible = False
 dtc_desc8.Visible = False
 dtc_codigo9.Visible = False
 dtc_desc9.Visible = False
 dtc_codigo10.Visible = False
 dtc_desc10.Visible = False
 
 dtc_codigo8.Text = ""
 dtc_desc8.Text = ""
 dtc_codigo9.Text = ""
 dtc_desc9.Text = ""
 dtc_codigo10.Text = ""
 dtc_desc10.Text = ""
End Sub

Private Sub DtcUE_Click(Area As Integer)
    DtcUE_Des.BoundText = DtcUE.BoundText
End Sub

Private Sub DtcUE_Des_Click(Area As Integer)
    DtcUE.BoundText = DtcUE_Des.BoundText
End Sub


Private Sub Buscar1_Click()
VAR_AUX1 = txt_Aux1
If txt_Aux1.Text <> "" Then
Call ABRIR_AUX_TABLA

    If VAR_TABLA = "NN" And txt_Aux1 = "00" Then
        dtc_codigo8.Text = "0"
        dtc_desc8.Text = "NO ASIGNADO"
        MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
    Else
         
        dtc_codigo8.Visible = True
        dtc_desc8.Visible = True
        Set rs_datos8 = New ADODB.Recordset
        If rs_datos8.State = 1 Then rs_datos8.Close
            If VAR_AUX1 = "02" Then
                If txt_SubCta1.Text = "01" Then
                    rs_datos8.Open "Select " + VAR_CODIGO + " as codigo1 , " + VAR_DES + " as desc1 from " + VAR_TABLA + " WHERE tipo_moneda = 'BOB' " + " order by " + VAR_DES, db, adOpenStatic
                Else
                    rs_datos8.Open "Select " + VAR_CODIGO + " as codigo1 , " + VAR_DES + " as desc1 from " + VAR_TABLA + " WHERE tipo_moneda = 'USD' " + " order by " + VAR_DES, db, adOpenStatic
                End If
            Else
                rs_datos8.Open "Select " + VAR_CODIGO + " as codigo1 , " + VAR_DES + " as desc1 from " + VAR_TABLA + " order by " + VAR_DES, db, adOpenStatic
            End If
            Set Ado_datos8.Recordset = rs_datos8
            dtc_desc8.BoundText = dtc_codigo8.BoundText
    End If
        Else
    MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
  End If
  dtc_codigo8.Text = ""
  dtc_desc8.Text = ""
End Sub

Private Sub Buscar2_Click()
 VAR_AUX1 = txt_Aux2
 If txt_Aux2.Text <> "" Then
Call ABRIR_AUX_TABLA

    If VAR_TABLA = "NN" And txt_Aux2 = "00" Then
        dtc_codigo9.Text = "0"
        dtc_desc9.Text = "NO ASIGNADO"
        MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
        
    Else
  dtc_codigo9.Visible = True
 dtc_desc9.Visible = True
        Set rs_datos9 = New ADODB.Recordset
        If rs_datos9.State = 1 Then rs_datos9.Close
            rs_datos9.Open "Select " + VAR_CODIGO + " as codigo2 , " + VAR_DES + " as desc2 from " + VAR_TABLA + " order by " + VAR_DES, db, adOpenStatic
            Set Ado_datos9.Recordset = rs_datos9
            dtc_desc9.BoundText = dtc_codigo9.BoundText
    End If
    Else
    MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
  End If
    dtc_codigo9.Text = ""
  dtc_desc9.Text = ""
End Sub

Private Sub Buscar3_Click()
VAR_AUX1 = txt_Aux3
 If txt_Aux3.Text <> "" Then
Call ABRIR_AUX_TABLA

    If VAR_TABLA = "NN" And txt_Aux3 = "00" Then
        dtc_codigo10.Text = "0"
        dtc_desc10.Text = "NO ASIGNADO"
        MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
    Else
  dtc_codigo10.Visible = True
 dtc_desc10.Visible = True
        Set rs_datos10 = New ADODB.Recordset
        If rs_datos10.State = 1 Then rs_datos10.Close
            rs_datos10.Open "Select " + VAR_CODIGO + " as codigo3 , " + VAR_DES + " as desc3 from " + VAR_TABLA + " order by " + VAR_DES, db, adOpenStatic
            Set Ado_datos10.Recordset = rs_datos10
            dtc_desc10.BoundText = dtc_codigo10.BoundText
    End If
        Else
    MsgBox "No existe AUX para registrarlo ...", vbInformation, "informacion"
  End If
    dtc_codigo10.Text = ""
  dtc_desc10.Text = ""
End Sub

Private Sub dg_datos_DblClick()
   
    Dim recsetaux As ADODB.Recordset
    Dim literales As String
    Dim decimal2 As String
    Dim LiteralCry As String
    Monto = 0
'    db.Execute "UPDATE co_diario SET NOMCTADEBE = (SELECT CC_Plan_Cuentas.NombreCta From CC_Plan_Cuentas Where CC_Plan_Cuentas.Cuenta =  co_diario.d_Cuenta and CC_Plan_Cuentas.SubCta1 = co_diario.d_Subcta1 and CC_Plan_Cuentas.SubCta2 = '00')"
    
    Set recsetaux = New ADODB.Recordset
    If rs_datos.RecordCount <> 0 Then
          If recsetaux.State = 1 Then recsetaux.Close
          recsetaux.Open "SELECT DISTINCT Co_Comprobante_M.Cod_Comp, Co_Comprobante_M.Tipo_Comp,CO_Diario.D_MontoBs " & _
                       "FROM Co_Comprobante_M INNER JOIN CO_Diario ON Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp " & _
                       "WHERE Co_Comprobante_M.Cod_Comp = " & Ado_datos.Recordset!comp2, db, adOpenForwardOnly, adLockReadOnly

        If recsetaux.RecordCount <> 0 Then
            Set rs_aux1 = New ADODB.Recordset
            If rs_aux1.State = 1 Then rs_aux1.Close
            rs_aux1.Open "select sum(d_montoBs) as totbs, sum(D_MontoDl) as totdl from co_diario where Cod_Comp = " & Ado_datos.Recordset!comp2 & "  ", db, adOpenKeyset, adLockOptimistic
            If rs_aux1.RecordCount > 0 Then
                LiteralCry = Str(rs_aux1!totbs)
                literales = Literal(Str(rs_aux1!totbs)) + " Bolivianos"
                db.Execute "Update Co_Comprobante_M Set literal = '" & literales & "'  Where Cod_Comp = " & Ado_datos.Recordset!comp2 & "  "
            Else
                literales = "CERO 00/100 Bolivianos"
            End If

            Do While Not recsetaux.EOF
            LiteralCry = Str(Int(recsetaux!d_montoBs))
                Monto = CDbl(Monto) + recsetaux!d_montoBs
                recsetaux.MoveNext
            Loop
            LiteralCry = Str(Int(Monto))
            recsetaux.MoveFirst
            decimal2 = Str(Round((recsetaux!d_montoBs - Val(LiteralCry)), 2))
            If Monto <> 0 Then
                literales = Literal(Str(Monto)) + " Bolivianos"

            Else
                literales = "CERO 00/100 Bolivianos"
            End If
            Dim iResult As Integer
            CryComp_Manual.Destination = crptToWindow
            CryComp_Manual.WindowState = crptMaximized
            CryComp_Manual.WindowShowPrintSetupBtn = True
            CryComp_Manual.WindowShowRefreshBtn = True
            CryComp_Manual.ReportFileName = App.Path & "\reportes\Contabilidad\cr_registro_diario.rpt"
            CryComp_Manual.StoredProcParam(0) = recsetaux!Cod_Comp
            CryComp_Manual.StoredProcParam(1) = recsetaux!tipo_comp
            'CryComp_Manual.StoredProcParam(2) = "g--"
            CryComp_Manual.StoredProcParam(2) = literales
            VAR_TIT = "MODULO CONTABILIDAD"
            If Ado_datos.Recordset!tipo = "REC" Then
                CryComp_Manual.Formulas(0) = "titulo = 'COMPROBANTE DE INGRESO' "                    ' '" & dtc_desc14.Text & "' "
            Else
                CryComp_Manual.Formulas(0) = "titulo = 'COMPROBANTE DE TRASPASO' "                     ' '" & dtc_desc14.Text & "' "
            End If
            CryComp_Manual.Formulas(1) = "titulo1 = '" & VAR_TIT & "' "
            '
            iResult = CryComp_Manual.PrintReport
            If iResult <> 0 Then
                   MsgBox CryComp_Manual.LastErrorNumber & " : " & CryComp_Manual.LastErrorString, vbExclamation + vbOKOnly, "Error..."
            End If
       End If
    Else

       Exit Sub
    End If
End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
dtc_desc10.BoundText = dtc_codigo10.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
dtc_desc8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_codigo9_Click(Area As Integer)
dtc_desc9.BoundText = dtc_codigo9.BoundText
End Sub

Private Sub dtc_desc10_Click(Area As Integer)
dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
dtc_codigo8.BoundText = dtc_desc8.BoundText
End Sub

Private Sub dtc_desc9_Click(Area As Integer)
 dtc_codigo9.BoundText = dtc_desc9.BoundText
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLA
'    txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_ABM.Enabled = True
'    dg_datos.Enabled = True
End Sub

Private Sub ABRIR_TABLA()

'  'cc_Plan_Cuentas
'    Set rs_detalle1 = New ADODB.Recordset
'    If rs_detalle1.State = 1 Then rs_detalle1.Close
'    rs_detalle1.Open "select * from cc_plan_cuentas where nivel ='5'", db, adOpenStatic
'    Set Ado_detalle1.Recordset = rs_detalle1
'    txt_NombreCta.BoundText = txt_Cuenta.BoundText
    
    '******se carga de los COMBO CUENTAS cc_Plan_Cuentas -------------
    Set rs_detalle1 = New ADODB.Recordset
    If rs_detalle1.State = 1 Then rs_detalle1.Close
    rs_detalle1.Open "SELECT Cuenta +'-'+SubCta1+'-'+SubCta2 as CuentaAux, Cuenta +'-'+SubCta1+'-'+SubCta2+' '+ltrim(NombreCta) as NombreCtaAux,* FROM CC_Plan_Cuentas WHERE Nivel = '5' ", db, adOpenKeyset, adLockReadOnly
    Set Ado_detalle1.Recordset = rs_detalle1
    txt_Cuenta.BoundText = txt_correl.BoundText
    txt_NombreCta.BoundText = txt_correl.BoundText
    txt_SubCta1.BoundText = txt_correl.BoundText
    txt_SubCta2.BoundText = txt_correl.BoundText
    txt_Aux1.BoundText = txt_correl.BoundText
    txt_Aux2.BoundText = txt_correl.BoundText
    txt_Aux3.BoundText = txt_correl.BoundText
    txt_cuenta_des.BoundText = txt_correl.BoundText
    

End Sub

Private Sub ABRIR_AUX_TABLA()
    Set rs_detalle2 = New ADODB.Recordset
    If rs_detalle2.State = 1 Then rs_detalle2.Close
     rs_detalle2.Open "Select * from cc_tipo_auxiliar where aux = '" & VAR_AUX1 & "' order by aux ", db, adOpenStatic
    If rs_detalle2.RecordCount > 0 Then
        VAR_TABLA = rs_detalle2!NombreTabla
        VAR_CODIGO = rs_detalle2!nombre_codigo
        VAR_DES = rs_detalle2!nombre_descripcion
    Else
        VAR_TABLA = "NN"
        VAR_CODIGO = "NN"
        VAR_DES = "NN"
    End If
'    Set Ado_datos5.Recordset = rs_datos5
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
      Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
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

'Private Sub cmdRefresh_Click()
'  'Esto sólo es necesario en aplicaciones multiusuario
'  On Error GoTo RefreshErr
'  rs_datos.Requery
'  Exit Sub
'RefreshErr:
'  MsgBox Err.Description
'End Sub

Private Sub opt_aux1_Click()
  limpiar
  Fra_Aux1.Visible = True
  Fra_Aux2.Visible = False
  Fra_Aux3.Visible = False
End Sub

Private Sub opt_aux2_Click()
  limpiar
  Fra_Aux1.Visible = False
  Fra_Aux2.Visible = True
  Fra_Aux3.Visible = False
End Sub

Private Sub opt_aux3_Click()
  limpiar
  Fra_Aux1.Visible = False
  Fra_Aux2.Visible = False
  Fra_Aux3.Visible = True
End Sub

Private Sub stp_fecha_final_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Fra_ABM1.Visible = True
    Fra_op_aux.Visible = False
    BtnBuscar.Visible = False
End Sub

Private Sub stp_fecha_inicio_LostFocus()
    Fra_op_aux.Visible = False
    BtnBuscar.Visible = False
End Sub

Private Sub txt_CtaHelp_Change()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from CC_Plan_Cuentas WHERE Nivel = '5' and NombreCta like '" & txt_CtaHelp & "' ", db, adOpenStatic
    If rs_datos1.RecordCount > 0 Then
        VAR_TABLA = rs_detalle2!NombreTabla
        VAR_CODIGO = rs_detalle2!nombre_codigo
        VAR_DES = rs_detalle2!nombre_descripcion
    Else
        VAR_TABLA = "NN"
        VAR_CODIGO = "NN"
        VAR_DES = "NN"
    End If
End Sub

Private Sub txt_Aux1_Click(Area As Integer)
  txt_correl.BoundText = txt_Aux1.BoundText
  txt_Cuenta.BoundText = txt_Aux1.BoundText
  txt_NombreCta.BoundText = txt_Aux1.BoundText
  txt_SubCta1.BoundText = txt_Aux1.BoundText
  txt_SubCta2.BoundText = txt_Aux1.BoundText
  txt_Aux2.BoundText = txt_Aux1.BoundText
  txt_Aux3.BoundText = txt_Aux1.BoundText
  txt_cuenta_des.BoundText = txt_Aux1.BoundText
End Sub

Private Sub txt_Aux2_Click(Area As Integer)
  txt_correl.BoundText = txt_Aux2.BoundText
  txt_Cuenta.BoundText = txt_Aux2.BoundText
  txt_NombreCta.BoundText = txt_Aux2.BoundText
  txt_SubCta1.BoundText = txt_Aux2.BoundText
  txt_SubCta2.BoundText = txt_Aux2.BoundText
  txt_Aux1.BoundText = txt_Aux2.BoundText
  txt_Aux3.BoundText = txt_Aux2.BoundText
  txt_cuenta_des.BoundText = txt_Aux2.BoundText
End Sub

Private Sub txt_Aux3_Click(Area As Integer)
  txt_correl.BoundText = txt_Aux3.BoundText
  txt_Cuenta.BoundText = txt_Aux3.BoundText
  txt_NombreCta.BoundText = txt_Aux3.BoundText
  txt_SubCta1.BoundText = txt_Aux3.BoundText
  txt_SubCta2.BoundText = txt_Aux3.BoundText
  txt_Aux1.BoundText = txt_Aux3.BoundText
  txt_Aux2.BoundText = txt_Aux3.BoundText
  txt_cuenta_des.BoundText = txt_Aux3.BoundText
End Sub

Private Sub txt_correl_Click(Area As Integer)
  txt_Cuenta.BoundText = txt_correl.BoundText
  txt_NombreCta.BoundText = txt_correl.BoundText
  txt_SubCta1.BoundText = txt_correl.BoundText
  txt_SubCta2.BoundText = txt_correl.BoundText
  txt_Aux1.BoundText = txt_correl.BoundText
  txt_Aux2.BoundText = txt_correl.BoundText
  txt_Aux3.BoundText = txt_correl.BoundText
  txt_cuenta_des.BoundText = txt_correl.BoundText
End Sub

Private Sub txt_Cuenta_Click(Area As Integer)
 txt_correl.BoundText = txt_Cuenta.BoundText
  txt_NombreCta.BoundText = txt_Cuenta.BoundText
  txt_SubCta1.BoundText = txt_Cuenta.BoundText
  txt_SubCta2.BoundText = txt_Cuenta.BoundText
  txt_Aux1.BoundText = txt_Cuenta.BoundText
  txt_Aux2.BoundText = txt_Cuenta.BoundText
  txt_Aux3.BoundText = txt_Cuenta.BoundText
  txt_cuenta_des.BoundText = txt_Cuenta.BoundText
End Sub

Private Sub txt_cuenta_des_Click(Area As Integer)
  txt_correl.BoundText = txt_cuenta_des.BoundText
  txt_Cuenta.BoundText = txt_cuenta_des.BoundText
  txt_SubCta1.BoundText = txt_cuenta_des.BoundText
  txt_SubCta2.BoundText = txt_cuenta_des.BoundText
  txt_Aux1.BoundText = txt_cuenta_des.BoundText
  txt_Aux2.BoundText = txt_cuenta_des.BoundText
  txt_Aux3.BoundText = txt_cuenta_des.BoundText
  txt_NombreCta.BoundText = txt_cuenta_des.BoundText
  limpiar
  Fra_op_aux.Visible = False
  BtnBuscar.Visible = False
End Sub

Private Sub txt_NombreCta_Click(Area As Integer)
  txt_correl.BoundText = txt_NombreCta.BoundText
  txt_Cuenta.BoundText = txt_NombreCta.BoundText
  txt_SubCta1.BoundText = txt_NombreCta.BoundText
  txt_SubCta2.BoundText = txt_NombreCta.BoundText
  txt_Aux1.BoundText = txt_NombreCta.BoundText
  txt_Aux2.BoundText = txt_NombreCta.BoundText
  txt_Aux3.BoundText = txt_NombreCta.BoundText
  txt_cuenta_des.BoundText = txt_NombreCta.BoundText
  limpiar
  Fra_op_aux.Visible = False
  BtnBuscar.Visible = False
End Sub

Private Sub txt_SubCta1_Click(Area As Integer)
  txt_correl.BoundText = txt_SubCta1.BoundText
  txt_Cuenta.BoundText = txt_SubCta1.BoundText
  txt_NombreCta.BoundText = txt_SubCta1.BoundText
  txt_SubCta2.BoundText = txt_SubCta1.BoundText
  txt_Aux1.BoundText = txt_SubCta1.BoundText
  txt_Aux2.BoundText = txt_SubCta1.BoundText
  txt_Aux3.BoundText = txt_SubCta1.BoundText
  txt_cuenta_des.BoundText = txt_SubCta1.BoundText
End Sub

Private Sub txt_SubCta2_Click(Area As Integer)
  txt_correl.BoundText = txt_SubCta2.BoundText
  txt_Cuenta.BoundText = txt_SubCta2.BoundText
  txt_NombreCta.BoundText = txt_SubCta2.BoundText
  txt_SubCta1.BoundText = txt_SubCta2.BoundText
  txt_Aux1.BoundText = txt_SubCta2.BoundText
  txt_Aux2.BoundText = txt_SubCta2.BoundText
  txt_Aux3.BoundText = txt_SubCta2.BoundText
  txt_cuenta_des.BoundText = txt_SubCta2.BoundText
End Sub

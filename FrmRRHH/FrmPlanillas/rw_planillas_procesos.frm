VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form rw_planillas_procesos 
   BackColor       =   &H00C0C0C0&
   Caption         =   "RRHH - Proceso de Planillas"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
   Icon            =   "rw_planillas_procesos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   11280
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fra_generar 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFF00&
      Height          =   2415
      Left            =   7320
      TabIndex        =   38
      Top             =   1680
      Visible         =   0   'False
      Width           =   6135
      Begin VB.ComboBox cmb_nro_planillas 
         Height          =   315
         ItemData        =   "rw_planillas_procesos.frx":0A02
         Left            =   3840
         List            =   "rw_planillas_procesos.frx":0A33
         TabIndex        =   49
         Text            =   "0"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox cmb_gestion 
         Height          =   315
         ItemData        =   "rw_planillas_procesos.frx":0A73
         Left            =   1320
         List            =   "rw_planillas_procesos.frx":0A98
         TabIndex        =   48
         Text            =   "Combo1"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5880
         TabIndex        =   39
         Top             =   240
         Width           =   5880
         Begin VB.PictureBox Picture24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3195
            Picture         =   "rw_planillas_procesos.frx":0ADE
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   41
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox BtnGrabar18 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1680
            Picture         =   "rw_planillas_procesos.frx":13CA
            ScaleHeight     =   615
            ScaleWidth      =   1305
            TabIndex        =   40
            Top             =   0
            Width           =   1300
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
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
            Left            =   14175
            TabIndex        =   42
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero De Planillas Incluyendo Aguinaldos"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3600
         TabIndex        =   51
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Elije la Gestión para generar la Planilla"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1200
         TabIndex        =   50
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame fra_modificar2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modificar"
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
      Height          =   2535
      Left            =   120
      TabIndex        =   54
      Top             =   840
      Visible         =   0   'False
      Width           =   6375
      Begin MSComCtl2.DTPicker dtp_liq 
         DataField       =   "fecha_estimada_pla"
         DataSource      =   "Ado_datos1"
         Height          =   255
         Left            =   4920
         TabIndex        =   61
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CalendarTitleBackColor=   0
         CalendarTitleForeColor=   16776960
         Format          =   110100481
         CurrentDate     =   42570
      End
      Begin VB.TextBox txt_antecedente 
         DataField       =   "antecedente"
         DataSource      =   "Ado_datos1"
         Height          =   285
         Left            =   240
         TabIndex        =   60
         Top             =   1560
         Width           =   4590
      End
      Begin VB.PictureBox Picture28 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   6120
         TabIndex        =   55
         Top             =   240
         Width           =   6120
         Begin VB.PictureBox BtnGrabar30 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            Picture         =   "rw_planillas_procesos.frx":1BA0
            ScaleHeight     =   615
            ScaleWidth      =   1305
            TabIndex        =   57
            Top             =   0
            Width           =   1300
         End
         Begin VB.PictureBox Picture29 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3075
            Picture         =   "rw_planillas_procesos.frx":2376
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   56
            Top             =   0
            Width           =   1395
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
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
            Left            =   14175
            TabIndex        =   58
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Liquidación"
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
         Height          =   495
         Left            =   5040
         TabIndex        =   62
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Antecedente"
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
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.Frame Fra_personal_Ppla 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Planilla"
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
      Height          =   4935
      Left            =   2520
      TabIndex        =   109
      Top             =   5160
      Visible         =   0   'False
      Width           =   12015
      Begin VB.TextBox txt_impuesto_a_pagar 
         BackColor       =   &H00FFFFFF&
         DataField       =   "impuesto_a_pagar"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   191
         Text            =   "0"
         Top             =   3600
         Width           =   1470
      End
      Begin VB.TextBox txt_dependiente_a_favor1 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   171
         Text            =   "0"
         Top             =   3600
         Width           =   1470
      End
      Begin VB.TextBox txt_mes_anterior_mant1 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   170
         Text            =   "0"
         Top             =   3000
         Width           =   1470
      End
      Begin VB.TextBox txt_mes_anterior1 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   169
         Text            =   "0"
         Top             =   2280
         Width           =   1470
      End
      Begin VB.TextBox txt_fisco_a_favor1 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   168
         Text            =   "0"
         Top             =   3000
         Width           =   1470
      End
      Begin VB.TextBox txt_iva_1101 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6840
         TabIndex        =   167
         Text            =   "0"
         Top             =   2280
         Width           =   1470
      End
      Begin VB.TextBox txt_saldo_util1 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   166
         Text            =   "0"
         Top             =   3600
         Width           =   1470
      End
      Begin VB.TextBox txt_saldo_a_favor_depend1 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   165
         Text            =   "0"
         Top             =   2280
         Width           =   1470
      End
      Begin VB.TextBox txt_anticipo_refr1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   164
         Text            =   "0"
         Top             =   3000
         Width           =   1470
      End
      Begin VB.TextBox txt_anticipo_sb1 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         TabIndex        =   163
         Text            =   "0"
         Top             =   2280
         Width           =   1470
      End
      Begin VB.TextBox txt_prestamo1 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5160
         TabIndex        =   162
         Text            =   "0"
         Top             =   2280
         Width           =   1470
      End
      Begin VB.TextBox txt_otros_descuentos1 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   161
         Text            =   "0"
         Top             =   3000
         Width           =   1470
      End
      Begin VB.TextBox txt_saldo_a_favor_depend 
         BackColor       =   &H00000000&
         DataField       =   "saldo_a_favor_depend"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   159
         Text            =   "0"
         Top             =   2280
         Width           =   1470
      End
      Begin VB.TextBox txt_saldo_util 
         BackColor       =   &H00000000&
         DataField       =   "saldo_util"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   157
         Text            =   "0"
         Top             =   3600
         Width           =   1470
      End
      Begin VB.TextBox txt_refri 
         BackColor       =   &H00FFFFFF&
         DataField       =   "monto_refrigerio"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   133
         Text            =   "0"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txt_sueldo 
         BackColor       =   &H00FFFFFF&
         DataField       =   "sueldo_basico"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   132
         Text            =   "0"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txt_ci 
         BackColor       =   &H00FFFFFF&
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   131
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txt_nomb_ap 
         BackColor       =   &H00FFFFFF&
         DataField       =   "beneficiario_denominacion"
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   130
         Top             =   1440
         Visible         =   0   'False
         Width           =   5430
      End
      Begin VB.TextBox txt_afp2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "afp2"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         TabIndex        =   129
         Text            =   "0"
         Top             =   3600
         Width           =   1470
      End
      Begin VB.TextBox txt_bono_ant 
         BackColor       =   &H00FFFFFF&
         DataField       =   "bono_antiguedad"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "0"
         Top             =   3600
         Width           =   1470
      End
      Begin VB.TextBox txt_otros_descuentos 
         BackColor       =   &H00000000&
         DataField       =   "otros_dsctos"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   5160
         TabIndex        =   127
         Text            =   "0"
         Top             =   3000
         Width           =   1470
      End
      Begin VB.TextBox txt_prestamo 
         BackColor       =   &H00000000&
         DataField       =   "prestamo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   5160
         TabIndex        =   126
         Text            =   "0"
         Top             =   2280
         Width           =   1470
      End
      Begin VB.TextBox txt_afp1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "afp1"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5160
         TabIndex        =   125
         Text            =   "0"
         Top             =   3600
         Width           =   1470
      End
      Begin VB.TextBox txt_anticipo_sb 
         BackColor       =   &H00000000&
         DataField       =   "anticipo_sueldo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   2040
         TabIndex        =   124
         Text            =   "0"
         Top             =   2280
         Width           =   1470
      End
      Begin VB.TextBox txt_anticipo_refr 
         DataField       =   "anticipo_refrigerio"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         Height          =   285
         Left            =   2040
         TabIndex        =   123
         Text            =   "0"
         Top             =   3000
         Width           =   1470
      End
      Begin VB.TextBox txt_rc_iva 
         BackColor       =   &H00FFFFFF&
         DataField       =   "impuesto_a_pagar"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3600
         TabIndex        =   122
         Text            =   "0"
         Top             =   2280
         Width           =   1470
      End
      Begin VB.TextBox txt_liq_pagable 
         BackColor       =   &H00C0C0C0&
         DataField       =   "liquido_pagable_bs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
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
         Height          =   285
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   121
         Text            =   "0"
         Top             =   4320
         Width           =   1470
      End
      Begin VB.TextBox txt_total_descuento 
         BackColor       =   &H00C0C0C0&
         DataField       =   "total_dsctos"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   120
         Text            =   "0"
         Top             =   4320
         Width           =   1470
      End
      Begin VB.TextBox txt_total_ganado 
         BackColor       =   &H00C0C0C0&
         DataField       =   "total_ganado"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   119
         Text            =   "0"
         Top             =   4320
         Width           =   1470
      End
      Begin VB.PictureBox Picture31 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   11760
         TabIndex        =   115
         Top             =   240
         Width           =   11760
         Begin VB.PictureBox Picture33 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6195
            Picture         =   "rw_planillas_procesos.frx":2C62
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   117
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox BtnGrabar32 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4680
            Picture         =   "rw_planillas_procesos.frx":354E
            ScaleHeight     =   615
            ScaleWidth      =   1305
            TabIndex        =   116
            Top             =   0
            Width           =   1300
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
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
            Left            =   14175
            TabIndex        =   118
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.TextBox txt_iva_110 
         BackColor       =   &H00FFFFFF&
         DataField       =   "iva_110"
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6840
         TabIndex        =   114
         Text            =   "0"
         Top             =   2280
         Width           =   1470
      End
      Begin VB.TextBox txt_fisco_a_favor 
         BackColor       =   &H00000000&
         DataField       =   "fisco_a_favor"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "0"
         Top             =   3000
         Width           =   1470
      End
      Begin VB.TextBox txt_mes_anterior 
         BackColor       =   &H00000000&
         DataField       =   "mes_anterior"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   112
         Text            =   "0"
         Top             =   2280
         Width           =   1470
      End
      Begin VB.TextBox txt_mes_anterior_mant 
         BackColor       =   &H00000000&
         DataField       =   "saldo_para_mes_sig"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   111
         Text            =   "0"
         Top             =   3000
         Width           =   1470
      End
      Begin VB.TextBox txt_dependiente_a_favor 
         BackColor       =   &H00000000&
         DataField       =   "dependiente_a_favor"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   110
         Text            =   "0"
         Top             =   3600
         Width           =   1470
      End
      Begin MSDataListLib.DataCombo dtc_descripcion 
         Bindings        =   "rw_planillas_procesos.frx":3D24
         Height          =   315
         Left            =   360
         TabIndex        =   134
         Top             =   1440
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
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
      Begin MSDataListLib.DataCombo dtc_codigo 
         Bindings        =   "rw_planillas_procesos.frx":3D3D
         DataField       =   "beneficiario_codigo"
         Height          =   315
         Left            =   4080
         TabIndex        =   135
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
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
      Begin MSDataListLib.DataCombo dtc_sueldo 
         Bindings        =   "rw_planillas_procesos.frx":3D56
         Height          =   315
         Left            =   360
         TabIndex        =   136
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "beneficiario_haber_mensual"
         BoundColumn     =   "beneficiario_codigo"
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
      Begin MSDataListLib.DataCombo dtc_refrigerio 
         Bindings        =   "rw_planillas_procesos.frx":3D6F
         Height          =   315
         Left            =   360
         TabIndex        =   137
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "beneficiario_otro_mensual"
         BoundColumn     =   "beneficiario_codigo"
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
      Begin VB.Label Label53 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Impuesto a pagar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9960
         TabIndex        =   192
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFF00&
         X1              =   240
         X2              =   11760
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFF00&
         X1              =   240
         X2              =   6720
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFF00&
         X1              =   6720
         X2              =   6720
         Y1              =   1920
         Y2              =   4680
      End
      Begin VB.Label Label46 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo a Favor Depend"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9960
         TabIndex        =   160
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label45 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Util"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8400
         TabIndex        =   158
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFF00&
         X1              =   11760
         X2              =   11760
         Y1              =   1920
         Y2              =   3960
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFF00&
         X1              =   240
         X2              =   240
         Y1              =   1920
         Y2              =   4680
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFF00&
         X1              =   240
         X2              =   11760
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFF00&
         X1              =   -1440
         X2              =   10080
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFF00&
         X1              =   1920
         X2              =   1920
         Y1              =   1920
         Y2              =   4680
      End
      Begin VB.Label Label29 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "AFP 2"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   156
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label28 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Bono Antigüedad"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   155
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label27 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Otros Descuentos"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   154
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label26 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "RC-IVA"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   153
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label25 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "AFP 1"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   152
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label24 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Anticipo S.B."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   151
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label23 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Anticipo Refr"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   150
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label22 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Prestamo"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   149
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label21 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Liq. Pagable"
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
         Height          =   255
         Left            =   8880
         TabIndex        =   148
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Descuento"
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
         Height          =   255
         Left            =   5160
         TabIndex        =   147
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label18 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Ganado"
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
         Height          =   255
         Left            =   360
         TabIndex        =   146
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Refrigerio"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   145
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "C.I."
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
         Height          =   255
         Left            =   5760
         TabIndex        =   144
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Sueldo"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   143
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos y Nombre"
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
         Height          =   255
         Left            =   360
         TabIndex        =   142
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label40 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes Anterior"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8400
         TabIndex        =   141
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label41 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fisco a Favor"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6840
         TabIndex        =   140
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label42 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario 110"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6840
         TabIndex        =   139
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label43 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Para Mes Siguiente"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8400
         TabIndex        =   138
         Top             =   2760
         Width           =   1935
      End
   End
   Begin VB.Frame fra_imprimir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Imprimir"
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
      Height          =   5175
      Left            =   7200
      TabIndex        =   65
      Top             =   3000
      Visible         =   0   'False
      Width           =   6135
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Parametros"
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
         Height          =   1815
         Left            =   120
         TabIndex        =   90
         Top             =   960
         Width           =   5895
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "TODAS LAS PLANILLAS"
            ForeColor       =   &H00000000&
            Height          =   555
            Left            =   4560
            TabIndex        =   187
            Top             =   1080
            Width           =   1275
         End
         Begin VB.ComboBox cbo_mes_rep 
            Height          =   315
            ItemData        =   "rw_planillas_procesos.frx":3D88
            Left            =   1680
            List            =   "rw_planillas_procesos.frx":3DB0
            TabIndex        =   93
            Top             =   720
            Width           =   1575
         End
         Begin VB.ComboBox cmb_gestion_rep 
            Height          =   315
            ItemData        =   "rw_planillas_procesos.frx":3E19
            Left            =   1680
            List            =   "rw_planillas_procesos.frx":3E3E
            TabIndex        =   92
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txt_mes 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   91
            Text            =   "0"
            Top             =   720
            Visible         =   0   'False
            Width           =   630
         End
         Begin MSDataListLib.DataCombo dtc_rep_det 
            Bindings        =   "rw_planillas_procesos.frx":3E84
            Height          =   315
            Left            =   1920
            TabIndex        =   94
            Top             =   1200
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "planilla_descripcion"
            BoundColumn     =   "planilla_codigo"
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
         Begin MSDataListLib.DataCombo dtc_rep_cod 
            Bindings        =   "rw_planillas_procesos.frx":3EA0
            DataField       =   "planilla_codigo"
            Height          =   315
            Left            =   960
            TabIndex        =   95
            Top             =   1200
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "planilla_codigo"
            BoundColumn     =   "planilla_codigo"
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
         Begin VB.Label Label32 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "GESTIÓN"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   720
            TabIndex        =   98
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label33 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "MES"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   720
            TabIndex        =   97
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label34 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "PLANILLA"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   1320
            Width           =   855
         End
      End
      Begin VB.Frame fra_reportes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reportes"
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
         Height          =   2295
         Left            =   120
         TabIndex        =   81
         Top             =   2760
         Width           =   5895
         Begin VB.OptionButton rb_futuro 
            BackColor       =   &H00C0C0C0&
            Caption         =   "AFP FUTURO"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   84
            Top             =   840
            Width           =   1455
         End
         Begin VB.OptionButton rb_prevision 
            BackColor       =   &H00C0C0C0&
            Caption         =   "AFP PREVISIÓN"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   83
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton rb_pla_ministerio 
            BackColor       =   &H00C0C0C0&
            Caption         =   "PLANILLA MINISTERIO"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   82
            Top             =   1200
            Width           =   2175
         End
      End
      Begin VB.PictureBox Picture35 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5880
         TabIndex        =   66
         Top             =   240
         Width           =   5880
         Begin VB.PictureBox Picture37 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3360
            Picture         =   "rw_planillas_procesos.frx":3EBC
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   69
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox Picture36 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1680
            Picture         =   "rw_planillas_procesos.frx":47A8
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   68
            ToolTipText     =   "Imprimir el Listado de los Registros"
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
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
            Left            =   14175
            TabIndex        =   67
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.PictureBox Picture38 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5880
         TabIndex        =   85
         Top             =   240
         Width           =   5880
         Begin VB.PictureBox Picture39 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   1560
            Picture         =   "rw_planillas_procesos.frx":5075
            ScaleHeight     =   735
            ScaleWidth      =   600
            TabIndex        =   86
            ToolTipText     =   "Aprueba Cronograma"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label37 
            BackColor       =   &H80000006&
            Caption         =   "Cancelar"
            ForeColor       =   &H00FFFF00&
            Height          =   255
            Left            =   3840
            TabIndex        =   89
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label36 
            BackColor       =   &H80000006&
            Caption         =   "Aceptar"
            ForeColor       =   &H00FFFF00&
            Height          =   375
            Left            =   2160
            TabIndex        =   88
            Top             =   240
            Width           =   855
         End
         Begin VB.Image Image1 
            Height          =   375
            Left            =   3240
            OLEDropMode     =   1  'Manual
            Picture         =   "rw_planillas_procesos.frx":58A8
            Stretch         =   -1  'True
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
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
            Left            =   14175
            TabIndex        =   87
            Top             =   195
            Width           =   1005
         End
      End
   End
   Begin VB.Frame fra_ufv 
      BackColor       =   &H00C0C0C0&
      Caption         =   "UFV"
      ForeColor       =   &H00C00000&
      Height          =   3015
      Left            =   7680
      TabIndex        =   194
      Top             =   1800
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton Command4 
         Caption         =   "pruebas"
         Height          =   255
         Left            =   960
         TabIndex        =   206
         Top             =   2400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox Picture44 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5280
         TabIndex        =   197
         Top             =   240
         Width           =   5280
         Begin VB.PictureBox Picture46 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2880
            Picture         =   "rw_planillas_procesos.frx":C54D
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   199
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox BtnGrabar45 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1320
            Picture         =   "rw_planillas_procesos.frx":CE39
            ScaleHeight     =   615
            ScaleWidth      =   1305
            TabIndex        =   198
            Top             =   0
            Width           =   1300
         End
         Begin VB.Label Label54 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
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
            Left            =   14175
            TabIndex        =   200
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "rw_planillas_procesos.frx":D60F
         Left            =   1560
         List            =   "rw_planillas_procesos.frx":D637
         TabIndex        =   196
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   195
         Top             =   3000
         Visible         =   0   'False
         Width           =   630
      End
      Begin MSComCtl2.DTPicker DTP_ufv_ant 
         Height          =   285
         Left            =   720
         TabIndex        =   204
         Top             =   1560
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         _Version        =   393216
         CalendarTitleBackColor=   -2147483635
         CheckBox        =   -1  'True
         Format          =   110100481
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker DTC_ufv_actual 
         Height          =   285
         Left            =   3000
         TabIndex        =   205
         Top             =   1560
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         _Version        =   393216
         CalendarTitleBackColor=   -2147483635
         CheckBox        =   -1  'True
         Format          =   110100481
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin VB.Label Label58 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha UFV anterior"
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
         Height          =   255
         Left            =   720
         TabIndex        =   203
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label56 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha UFV actual"
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
         Height          =   255
         Left            =   3000
         TabIndex        =   202
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label55 
         BackColor       =   &H00000000&
         Caption         =   "Mes"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1080
         TabIndex        =   201
         Top             =   3000
         Width           =   735
      End
   End
   Begin VB.Frame fra_nueva_pla 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Crear Nueva Planilla"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   172
      Top             =   720
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox txt_mes_grupo 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   185
         Top             =   3000
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.ComboBox cbo_mes_pla 
         Height          =   315
         ItemData        =   "rw_planillas_procesos.frx":D6A0
         Left            =   1560
         List            =   "rw_planillas_procesos.frx":D6C8
         TabIndex        =   184
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox txt_concepto_pla 
         Height          =   285
         Left            =   1320
         TabIndex        =   181
         Top             =   2400
         Width           =   3870
      End
      Begin VB.PictureBox Picture41 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5280
         TabIndex        =   174
         Top             =   240
         Width           =   5280
         Begin VB.PictureBox BtnGrabar43 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1320
            Picture         =   "rw_planillas_procesos.frx":D731
            ScaleHeight     =   615
            ScaleWidth      =   1305
            TabIndex        =   176
            Top             =   0
            Width           =   1300
         End
         Begin VB.PictureBox Picture42 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2880
            Picture         =   "rw_planillas_procesos.frx":DF07
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   175
            Top             =   0
            Width           =   1400
         End
         Begin VB.Label Label47 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
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
            Left            =   14175
            TabIndex        =   177
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.ComboBox cbo_gestion_pla 
         Height          =   315
         ItemData        =   "rw_planillas_procesos.frx":E7F3
         Left            =   1320
         List            =   "rw_planillas_procesos.frx":E818
         TabIndex        =   173
         Text            =   "Combo1"
         Top             =   1200
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dtc_pla_det 
         Bindings        =   "rw_planillas_procesos.frx":E85E
         Height          =   315
         Left            =   2160
         TabIndex        =   179
         Top             =   1800
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "planilla_descripcion"
         BoundColumn     =   "planilla_codigo"
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
      Begin MSDataListLib.DataCombo dtc_pla_cod 
         Bindings        =   "rw_planillas_procesos.frx":E878
         DataField       =   "planilla_codigo"
         Height          =   315
         Left            =   1320
         TabIndex        =   180
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "planilla_codigo"
         BoundColumn     =   "planilla_codigo"
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
      Begin VB.Label Label51 
         BackColor       =   &H00000000&
         Caption         =   "Mes"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1080
         TabIndex        =   186
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label50 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Planilla"
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
         Height          =   255
         Left            =   360
         TabIndex        =   183
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label48 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
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
         Height          =   255
         Left            =   360
         TabIndex        =   182
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label49 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Gestión"
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
         Height          =   255
         Left            =   360
         TabIndex        =   178
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Frame fra_sub_grupo_unidad 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Crear Sub Grupo Por Unidad"
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
      Height          =   2295
      Left            =   120
      TabIndex        =   99
      Top             =   720
      Visible         =   0   'False
      Width           =   5535
      Begin MSDataListLib.DataCombo dt_unidad_det 
         Bindings        =   "rw_planillas_procesos.frx":E892
         DataField       =   "unidad_codigo_pla"
         Height          =   315
         Left            =   1800
         TabIndex        =   105
         Top             =   1200
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "unidad_descripcion_pla"
         BoundColumn     =   "unidad_codigo_pla"
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
      Begin VB.TextBox cbo_numero_pago 
         DataSource      =   "Ado_datos4"
         Height          =   285
         Left            =   2040
         TabIndex        =   108
         Text            =   "1"
         Top             =   1680
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5280
         TabIndex        =   100
         Top             =   240
         Width           =   5280
         Begin VB.PictureBox Picture13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2835
            Picture         =   "rw_planillas_procesos.frx":E8AB
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   102
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox BtnGrabar40 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1320
            Picture         =   "rw_planillas_procesos.frx":F197
            ScaleHeight     =   615
            ScaleWidth      =   1305
            TabIndex        =   101
            Top             =   0
            Width           =   1300
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
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
            Left            =   14175
            TabIndex        =   103
            Top             =   195
            Width           =   1005
         End
      End
      Begin MSDataListLib.DataCombo dt_unidad_cod 
         Bindings        =   "rw_planillas_procesos.frx":F96D
         DataField       =   "unidad_codigo_pla"
         Height          =   315
         Left            =   960
         TabIndex        =   104
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "unidad_codigo_pla"
         BoundColumn     =   "unidad_codigo_pla"
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
      Begin VB.Label Label39 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. de Pago"
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
         Height          =   255
         Left            =   720
         TabIndex        =   107
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label38 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad"
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
         Height          =   255
         Left            =   240
         TabIndex        =   106
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Frame Fra_modificar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modificar"
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
      Height          =   2175
      Left            =   7200
      TabIndex        =   43
      Top             =   3360
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox txt_descripcion_grupo 
         DataField       =   "descripcion_grupo"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   240
         TabIndex        =   53
         Top             =   1560
         Width           =   5670
      End
      Begin VB.PictureBox Picture25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H00FFFF00&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5880
         TabIndex        =   44
         Top             =   240
         Width           =   5880
         Begin VB.PictureBox Picture27 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1680
            Picture         =   "rw_planillas_procesos.frx":F986
            ScaleHeight     =   615
            ScaleWidth      =   1305
            TabIndex        =   46
            Top             =   0
            Width           =   1300
         End
         Begin VB.PictureBox Picture26 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3120
            Picture         =   "rw_planillas_procesos.frx":1015C
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   45
            Top             =   0
            Width           =   1400
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
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
            Left            =   14175
            TabIndex        =   47
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción Planilla"
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
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1200
         Width           =   2175
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Listado"
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
      Height          =   7815
      Left            =   0
      TabIndex        =   34
      Top             =   720
      Width           =   5535
      Begin ComctlLib.ProgressBar ProgressBar3 
         Height          =   270
         Left            =   120
         TabIndex        =   209
         Top             =   7080
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
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
         Left            =   1200
         TabIndex        =   70
         Top             =   7440
         Value           =   -1  'True
         Width           =   1455
      End
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
         TabIndex        =   64
         Top             =   7440
         Width           =   915
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   7320
         Width           =   5280
         _ExtentX        =   9313
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
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "rw_planillas_procesos.frx":10A48
         Height          =   7050
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   12435
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "ges_gestion"
            Caption         =   "Gestión"
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
            DataField       =   "planilla_codigo"
            Caption         =   "Planilla"
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
            DataField       =   "descripcion_grupo"
            Caption         =   "Descripcion.Planilla"
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
            DataField       =   "depto_codigo"
            Caption         =   "Depto"
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
            DataField       =   "mes_grupo"
            Caption         =   "Mes"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   3660.095
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Personal de la Planilla"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4335
      Left            =   5640
      TabIndex        =   20
      Top             =   4200
      Width           =   10485
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   208
         Top             =   3960
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.PictureBox fra_opciones_det_2 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   10200
         TabIndex        =   21
         Top             =   240
         Width           =   10200
         Begin VB.CommandButton Command5 
            BackColor       =   &H00808080&
            Caption         =   "RETROACTIVO"
            Height          =   720
            Left            =   7440
            MaskColor       =   &H00808080&
            Style           =   1  'Graphical
            TabIndex        =   207
            Top             =   0
            Width           =   1245
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00808080&
            Caption         =   "Generar Planilla Tributaria"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   8640
            MaskColor       =   &H00808080&
            Style           =   1  'Graphical
            TabIndex        =   193
            Top             =   0
            Width           =   1605
         End
         Begin VB.PictureBox Picture20 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6000
            Picture         =   "rw_planillas_procesos.frx":10A60
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   27
            ToolTipText     =   "Imprimir el Listado de los Registros"
            Top             =   0
            Width           =   1455
         End
         Begin VB.PictureBox Picture19 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   17760
            Picture         =   "rw_planillas_procesos.frx":1132D
            ScaleHeight     =   615
            ScaleWidth      =   1245
            TabIndex        =   26
            ToolTipText     =   "Cierra la Ventana Activa"
            Top             =   0
            Width           =   1245
         End
         Begin VB.PictureBox Picture17 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4440
            Picture         =   "rw_planillas_procesos.frx":11AEF
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   25
            ToolTipText     =   "Aprueba Cronograma"
            Top             =   0
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.PictureBox Picture16 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3120
            Picture         =   "rw_planillas_procesos.frx":12322
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   24
            ToolTipText     =   "Anular Cronograma"
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox Picture15 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            Picture         =   "rw_planillas_procesos.frx":12A6E
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   23
            ToolTipText     =   "Modifica Datos Del Detalle"
            Top             =   0
            Visible         =   0   'False
            Width           =   1430
         End
         Begin VB.PictureBox BtnNuevo14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            Picture         =   "rw_planillas_procesos.frx":13383
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   22
            Top             =   0
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00808080&
            Height          =   600
            Left            =   6360
            Picture         =   "rw_planillas_procesos.frx":13B42
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   0
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CRONOGRAMA"
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
            Left            =   12255
            TabIndex        =   28
            Top             =   200
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   10200
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   10200
         Begin VB.PictureBox Picture23 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2880
            Picture         =   "rw_planillas_procesos.frx":13D4C
            ScaleHeight     =   615
            ScaleWidth      =   1305
            TabIndex        =   31
            Top             =   0
            Width           =   1300
         End
         Begin VB.PictureBox Picture22 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4275
            Picture         =   "rw_planillas_procesos.frx":14522
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   30
            Top             =   0
            Width           =   1400
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
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
            Left            =   14175
            TabIndex        =   32
            Top             =   195
            Width           =   1005
         End
      End
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "rw_planillas_procesos.frx":14E0E
         Height          =   2970
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   5239
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
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
         ColumnCount     =   16
         BeginProperty Column00 
            DataField       =   "beneficiario_codigo"
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
         BeginProperty Column01 
            DataField       =   "sueldo_basico"
            Caption         =   "Sueldo.Basico"
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
         BeginProperty Column02 
            DataField       =   "monto_refrigerio"
            Caption         =   "Refrigerio"
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
         BeginProperty Column03 
            DataField       =   "total_ganado"
            Caption         =   "Total.Ganado"
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
         BeginProperty Column04 
            DataField       =   "total_dsctos"
            Caption         =   "Total.Dsctos."
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
         BeginProperty Column05 
            DataField       =   "liquido_pagable_bs"
            Caption         =   "Liq.Pagable"
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
         BeginProperty Column06 
            DataField       =   "beneficiario_denominacion"
            Caption         =   "Apellidos y Nombres"
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
         BeginProperty Column07 
            DataField       =   "bono_antiguedad"
            Caption         =   "Antiguedad"
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
         BeginProperty Column08 
            DataField       =   "anticipo_sueldo"
            Caption         =   "Anticipo.S.B."
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
         BeginProperty Column09 
            DataField       =   "anticipo_refrigerio"
            Caption         =   "Anticipo.Refr."
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
         BeginProperty Column10 
            DataField       =   "prestamo"
            Caption         =   "Prestamo"
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
         BeginProperty Column11 
            DataField       =   "afp1"
            Caption         =   "PREVISION"
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
         BeginProperty Column12 
            DataField       =   "afp1"
            Caption         =   "FUTURO"
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
         BeginProperty Column13 
            DataField       =   "impuesto_a_pagar"
            Caption         =   "RC-IVA"
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
         BeginProperty Column14 
            DataField       =   "otros_dsctos"
            Caption         =   "Otros.Dsctos."
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
         BeginProperty Column15 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2385.071
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column15 
               Alignment       =   2
               ColumnWidth     =   705.26
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
      ScaleWidth      =   20280
      TabIndex        =   1
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7080
         Picture         =   "rw_planillas_procesos.frx":14E2C
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   11
         ToolTipText     =   "Imprimir el Listado de los Registros"
         Top             =   0
         Width           =   1455
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17760
         Picture         =   "rw_planillas_procesos.frx":156F9
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   9
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5760
         Picture         =   "rw_planillas_procesos.frx":15EBB
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   8
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4320
         Picture         =   "rw_planillas_procesos.frx":16670
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   7
         ToolTipText     =   "Aprueba Cronograma"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3000
         Picture         =   "rw_planillas_procesos.frx":16EA3
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   6
         ToolTipText     =   "Anular Cronograma"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1545
         Picture         =   "rw_planillas_procesos.frx":175EF
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   5
         ToolTipText     =   "Modifica Datos Cabecera"
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "rw_planillas_procesos.frx":17F04
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   4
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   8640
         Picture         =   "rw_planillas_procesos.frx":186C3
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   9720
         Picture         =   "rw_planillas_procesos.frx":18B05
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
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
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   12720
         TabIndex        =   10
         Top             =   195
         Width           =   885
      End
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
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2880
         Picture         =   "rw_planillas_procesos.frx":18D0F
         ScaleHeight     =   615
         ScaleWidth      =   1305
         TabIndex        =   14
         Top             =   0
         Width           =   1300
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4275
         Picture         =   "rw_planillas_procesos.frx":194E5
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   13
         Top             =   0
         Width           =   1400
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENTAS"
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
         Left            =   14175
         TabIndex        =   15
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sub Grupo (Sub Planilla)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3375
      Left            =   5640
      TabIndex        =   0
      Top             =   720
      Width           =   10485
      Begin ComctlLib.ProgressBar ProgressBar2 
         Height          =   270
         Left            =   120
         TabIndex        =   210
         Top             =   3120
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.PictureBox fra_opciones_det_1 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   10200
         TabIndex        =   71
         Top             =   240
         Width           =   10200
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   360
            Picture         =   "rw_planillas_procesos.frx":19DD1
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   79
            ToolTipText     =   "Imprimir el Listado de los Registros"
            Top             =   600
            Width           =   1455
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   17760
            Picture         =   "rw_planillas_procesos.frx":1A69E
            ScaleHeight     =   615
            ScaleWidth      =   1245
            TabIndex        =   78
            ToolTipText     =   "Cierra la Ventana Activa"
            Top             =   0
            Width           =   1245
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3960
            Picture         =   "rw_planillas_procesos.frx":1AE60
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   77
            ToolTipText     =   "Aprueba Cronograma"
            Top             =   0
            Width           =   1320
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2760
            Picture         =   "rw_planillas_procesos.frx":1B693
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   76
            ToolTipText     =   "Anular Cronograma"
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1320
            Picture         =   "rw_planillas_procesos.frx":1BDDF
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   75
            ToolTipText     =   "Modifica Datos Modifica Datos Del Detalle"
            Top             =   0
            Width           =   1430
         End
         Begin VB.PictureBox BtnNuevo9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "rw_planillas_procesos.frx":1C6F4
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   72
            Top             =   0
            Width           =   1200
         End
         Begin MSDataListLib.DataCombo dtc_buscar_desc 
            Bindings        =   "rw_planillas_procesos.frx":1CEB3
            Height          =   315
            Left            =   6720
            TabIndex        =   188
            Top             =   240
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
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
         Begin MSDataListLib.DataCombo dtc_buscar_ci 
            Bindings        =   "rw_planillas_procesos.frx":1CED0
            DataField       =   "beneficiario_codigo"
            Height          =   315
            Left            =   8280
            TabIndex        =   190
            Top             =   0
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
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
         Begin VB.CommandButton Command2 
            BackColor       =   &H00808080&
            Height          =   600
            Left            =   7680
            Picture         =   "rw_planillas_procesos.frx":1CEED
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   0
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.PictureBox Picture34 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5280
            Picture         =   "rw_planillas_procesos.frx":1D0F7
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   73
            ToolTipText     =   "Imprimir el Listado de los Registros"
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label52 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Buscar..."
            ForeColor       =   &H00FFFF00&
            Height          =   255
            Left            =   6720
            TabIndex        =   189
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CRONOGRAMA"
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
            Left            =   12255
            TabIndex        =   80
            Top             =   200
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   10200
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   10200
         Begin VB.PictureBox Picture12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4275
            Picture         =   "rw_planillas_procesos.frx":1D9C4
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   18
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2880
            Picture         =   "rw_planillas_procesos.frx":1E2B0
            ScaleHeight     =   615
            ScaleWidth      =   1305
            TabIndex        =   17
            Top             =   0
            Width           =   1300
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
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
            Left            =   14175
            TabIndex        =   19
            Top             =   195
            Width           =   1005
         End
      End
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "rw_planillas_procesos.frx":1EA86
         Height          =   2010
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   3545
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "ges_gestion"
            Caption         =   "Gestion"
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
            DataField       =   "planilla_codigo"
            Caption         =   "Planilla"
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
            DataField       =   "mes_grupo"
            Caption         =   "Mes"
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
            DataField       =   "numero_pago"
            Caption         =   "Pago"
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
            DataField       =   "unidad_codigo_pla"
            Caption         =   "Unidad"
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
            DataField       =   "antecedente"
            Caption         =   "Antecedente"
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
            DataField       =   "fecha_estimada_pla"
            Caption         =   "Fecha.Liq."
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   6419.906
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
      Top             =   9840
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
   Begin MSAdodcLib.Adodc Ado_datos4 
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
      Left            =   0
      Top             =   9240
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
   Begin MSAdodcLib.Adodc Ado_datos_rep 
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
   Begin Crystal.CrystalReport CR02 
      Left            =   0
      Top             =   8760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   0   'False
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos6 
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc Ado_datos11 
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
   Begin MSAdodcLib.Adodc Ado_datos_busq 
      Height          =   330
      Left            =   8760
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
      Caption         =   "Ado_datos_busq"
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
   Begin VB.Label Label19 
      BackColor       =   &H00000000&
      Caption         =   "Total Ganado"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   10080
      TabIndex        =   63
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "rw_planillas_procesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
'Dim rs_datos5 As New ADODB.Recordset
 Dim rsNada As New ADODB.Recordset
 Dim permisos As Integer
 Dim TOTSALBN As Double
 Dim ufv_inicio, ufv_fin, FIN As Date
Dim rs_numeracion As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset
Dim rs_aux9 As New ADODB.Recordset
Dim rs_aux10 As New ADODB.Recordset
Dim rs_aux11 As New ADODB.Recordset
Dim rs_aux12 As New ADODB.Recordset
Dim rs_aux13 As New ADODB.Recordset
Dim rs_aux14 As New ADODB.Recordset
Dim rs_aux15 As New ADODB.Recordset
Dim rs_aux16 As New ADODB.Recordset
Dim rs_aux17 As New ADODB.Recordset
Dim rs_aux18 As New ADODB.Recordset
Dim rs_aux19 As New ADODB.Recordset
Dim rs_aux20 As New ADODB.Recordset
Dim rs_aux21 As New ADODB.Recordset
Dim rs_aux22 As New ADODB.Recordset
Dim rs_aux23 As New ADODB.Recordset
Dim rs_aux24 As New ADODB.Recordset
Dim rs_aux25 As New ADODB.Recordset
Dim rs_aux26 As New ADODB.Recordset
Dim rs_aux27 As New ADODB.Recordset
Dim rs_aux28 As New ADODB.Recordset
Dim rs_aux29 As New ADODB.Recordset
Dim rs_aux30 As New ADODB.Recordset
Dim rs_aux31 As New ADODB.Recordset
Dim rs_aux32 As New ADODB.Recordset
Dim rs_aux_retro As New ADODB.Recordset
Dim persona As New ADODB.Recordset
Dim retroactivos As New ADODB.Recordset
Dim mover As Integer
'Dim busq As Integer
'Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim totalminutos As Integer
Dim DIA_HOY, NUM_PERS As Integer
Dim MES_IN As Integer
Dim ANO_IN As Integer
Dim DIA_IN As Integer
Dim MES_FN As Integer
Dim ANO_FN As Integer
Dim DIA_FN, expira, inicia, num_promedio As Integer

Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim queryinicial As String
Dim mesnom, fecha_pla, f1, f2, f3, f4, f5, f6, f7, f8 As String
'OTROS
Dim imag2 As Long

Dim VAR_MOD, VAR_MOD1, VAR_MOD2 As String
Dim SQL_FOR As String
Dim sql As String

Dim continuar As String
Dim Numero As Integer
Dim sino As String
Dim NombreCarpeta, e As String
Dim parametro As String
Dim var_titulo, VAR_SubTitulo As String
Dim var_cod, VAR_GES As String
Dim VAR_VAL, VAR_ARCH, VAR_ARCH2 As String
Dim VAR_SW, VAR_BENEF As String
'Dim permisos As String
Dim VAR_AUX, VAR_CONT2, PRESTAMO_TOTAL As Double
Dim var_campoc31, var_campoc32, var_campoc33, var_campoc34 As Double
Dim var_campod11, var_campod12, var_campod13, var_campod14 As Double
Dim var_campoe11, var_campoe12, var_campoe13, var_campoe14 As Double
Dim var_campoe21, var_campoe22, var_campoe23, var_campoe24 As Double
Dim var_campoe31, var_campoe32, var_campoe33, var_campoe34 As Double
Dim var_campoe41, var_campoe42, var_campoe43, var_campoe44 As Double
Dim var_campog11, var_campog12, var_campog13, var_campog14 As Double
Dim var_campog21, var_campog22, var_campog23, var_campog24 As Double
Dim VAR_IVA, VAR_NETO, promedio_haber, promedio_bono, promedio_otro, promedio_totalg As Double

Dim mes2 As Integer

Dim mvBookMark, marca1 As Variant
Dim mbDataChanged As Boolean
Dim nuevos, expirados As String

'PARA CONTABILIZAR PLANILLAS
Dim VAR_BENEF1, VAR_BENEF2, VAR_BENEF3, VAR_BENEF4 As String
Dim VAR_CITE, VAR_GLOSA As String
Dim VAR_CTA, VAR_PROY2, VAR_COD4, VAR_TIPOV As String
Dim NroFactura, var_literal, VAR_MONEDA, VAR_CODTIPO, VAR_DOC, VAR_ETAPA, VAR_TCOMP As String
Dim VAR_ANIO, gestion0, VAR_MES As String
Dim VAR_PLA, VAR_SPLA, DESAUX As String
Dim fecha_aux, fecha_aux2  As Date

Dim VAR_DOL2, VAR_BS2 As Double
Dim VAR_AFP1_BS, VAR_AFP2_BS, VAR_AFP1_BS2, VAR_AFP2_BS2 As Double
Dim VAR_DSCTO_BS, VAR_SS_BS, VAR_FV_BS As Double
Dim VAR_INDE_BS, VAR_AGUI_BS As Double
Dim VAR_AFPJ As Double

Dim VAR_FFAC As Date


Function fun_dias360(fecha_ini, fecha_fin)

If Year(fecha_fin) > Year(Date) Then
fecha_fin = "31/12/" & Year(Date)
End If

If Year(fecha_ini) < Year(Date) Then
fecha_ini = "01/01/" & Year(Date)
End If
 
dia_ini = Day(fecha_ini)
dia_fin = Day(fecha_fin)
mes_ini = Month(fecha_ini)
mes_fin = Month(fecha_fin)
año_ini = Year(fecha_ini)
año_fin = Year(fecha_fin)
mes_dif = mes_fin - mes_ini
If dia_fin > 30 And mes_dif <> 0 Then
dia_fin = 30
End If
'If año_ini Mod 4 = 0 Then
'If (año_ini Mod 100 = 0) And Not (año_ini Mod 400 = 0) Then
'If dia_ini < 29 Then
'dia_fin_aux = 30 - (dia_ini + 2)
'dia_ini = dia_fin_aux
'End If
'Else
'dia_fin_aux = 30 - (dia_ini + 1)
'dia_ini = 30 - dia_fin_aux
'End If
'Else
'dia_fin_aux = 30 - (dia_ini + 2)
'dia_ini = 30 - dia_fin_aux
'End If
'
'If Month(fecha_ini) + 1 <> mes_ini And mes_dif <> 0 Then
'dia_ini = 30
'End If
dia_dif = dia_fin - (dia_ini - 1)
año_dif = año_fin - año_ini
dif = año_dif * 360 + mes_dif * 30 + dia_dif
fun_dias360 = dif / 30

If fun_dias360 < 3 Then
fun_dias360 = 0
End If

End Function


Private Sub generar_aguinaldo()
'VARIABLE PARA SACAR PROMEDIO DESDE EL MES 9
num_promedio = 9
NUM_PERS = 0
Dim rs_aux16 As New ADODB.Recordset
'ABRE LAS SUB PLANILLAS
If rs_aux16.State = 1 Then rs_aux16.Close
rs_aux16.Open "select * from ro_pagos_cronograma where ges_gestion = '" & rs_datos!ges_gestion & "' AND planilla_codigo = '" & rs_datos!planilla_codigo & "' AND mes_grupo = " & rs_datos!mes_grupo & "", db, adOpenKeyset, adLockReadOnly  ', adOpenKeyset, adOpenStatic, adCmdText
rs_aux16.MoveFirst
While Not rs_aux16.EOF
    ProgressBar1.Visible = True
    Dim rs_aux6 As New ADODB.Recordset
    'ABRE LAS PERSONAS CONTRATADAS
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "SELECT * FROM ro_personal_contratado WHERE unidad_codigo_pla = '" & rs_aux16!unidad_codigo_pla & "' AND estado_jubilado = 'REG' AND year(fecha_expiracion) >= '" & rs_datos!ges_gestion & "'", db, adOpenKeyset, adLockOptimistic 'adOpenStatic 'order by beneficiario_denominacion
    'rs_aux6.Open "SELECT * FROM av_ro_peronal_vs_gc_beneficiario  WHERE unidad_codigo = '" & rs_datos1!unidad_codigo_pla & "' AND estado_codigo = 'APR' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic
    If rs_aux6.RecordCount > 0 Then 'verifica si existe personal en esa sub_planilla
       rs_aux6.MoveFirst
       With ProgressBar1
        .Max = rs_aux6.RecordCount
        .Min = 0
        .Value = 0
       End With
      'ProgressBar1.Max =
   '
       While Not rs_aux6.EOF
       
        ProgressBar1.Value = ProgressBar1.Value + 1
            DIA_FN = Day(rs_aux6!fecha_expiracion) 'FECHA FIN
            MES_FN = Month(rs_aux6!fecha_expiracion)
            ANO_FN = Year(rs_aux6!fecha_expiracion)
            If rs_aux6!beneficiario_codigo = "6753027" Then
            sino = ""
            End If
            
'         If rs_aux6!beneficiario_codigo = "159256" Then
'          MsgBox ""
'         End If
           
     expira = Day(DateSerial(rs_datos!ges_gestion, rs_datos!mes_grupo + 1, 0))
     'COMPONE FECHA PARA COMPARAR CON EL FIN DEL CONTRATO
     FIN = "01/12/" & rs_datos!ges_gestion
    
     'COMPARA FECHA DE FIN
    If rs_aux6!fecha_expiracion >= 0 Then 'PARA LAS BAJAS
        
        Dim rs_aux5 As New ADODB.Recordset
        If rs_aux5.State = 1 Then rs_aux5.Close
'        rs_aux5.Open "select * from ro_pagos_cronograma_detalle where ges_gestion = '" & Ado_datos1.Recordset!ges_gestion & "' AND planilla_codigo = '" & Ado_datos1.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos1.Recordset!mes_grupo & " AND numero_pago = " & Ado_datos1.Recordset!NUMERO_PAGO & " AND beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "' AND unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo_pla & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
       'VERIFICA SI LA PERSONA YA ESTA EN LA PLANILLA
        rs_aux5.Open "select * from ro_pagos_cronograma_detalle where ges_gestion = '" & Ado_datos1.Recordset!ges_gestion & "' AND planilla_codigo = '" & Ado_datos1.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos1.Recordset!mes_grupo & " AND numero_pago = " & Ado_datos1.Recordset!NUMERO_PAGO & " AND beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        If rs_aux5.RecordCount = 0 Then
         'REUNE DATOS PARA SACAR EL PROMEDIO
            While num_promedio <> 12
                Dim rs_aux27 As New ADODB.Recordset
                If rs_aux27.State = 1 Then rs_aux27.Close
                rs_aux27.Open "select * from ro_pagos_cronograma_detalle where ges_gestion = '" & rs_datos!ges_gestion & "' AND beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "' AND mes_grupo = " & num_promedio & " AND unidad_codigo = '" & rs_aux16!unidad_codigo_pla & "'", db, adOpenKeyset, adLockReadOnly  ', adOpenKeyset, adOpenStatic, adCmdText
                If rs_aux27.RecordCount > 0 Then
                    If num_promedio = 11 Then
                        promedio_haber = promedio_haber + Math.Round(rs_aux27!sueldo_basico, 2)
                        promedio_bono = promedio_bono + IIf(IsNull(Math.Round(rs_aux27!bono_antiguedad, 2)), 0, Math.Round(rs_aux27!bono_antiguedad, 2))
                        promedio_otro = promedio_otro + IIf(IsNull(Math.Round(rs_aux27!bono_transporte, 2)), 0, Math.Round(rs_aux27!bono_transporte, 2))
                        promedio_totalg = promedio_totalg + IIf(IsNull(Math.Round(rs_aux27!total_ganado, 2)), 0, Math.Round(rs_aux27!total_ganado, 2))
            '            promedio_haber = promedio_haber + Math.Round(rs_aux27!sueldo_basico * 2, 2)
            '            promedio_bono = promedio_bono + IIf(IsNull(Math.Round(rs_aux27!bono_antiguedad, 2)), 0, Math.Round(rs_aux27!bono_antiguedad * 2, 2))
            '            promedio_otro = promedio_otro + IIf(IsNull(Math.Round(rs_aux27!bono_transporte, 2)), 0, Math.Round(rs_aux27!bono_transporte * 2, 2))
            '            promedio_totalg = promedio_totalg + IIf(IsNull(Math.Round(rs_aux27!total_ganado, 2)), 0, Math.Round(rs_aux27!total_ganado * 2, 2))
                
                    Else
                        promedio_haber = promedio_haber + Math.Round(rs_aux27!sueldo_basico, 2)
                        promedio_bono = promedio_bono + IIf(IsNull(Math.Round(rs_aux27!bono_antiguedad, 2)), 0, Math.Round(rs_aux27!bono_antiguedad, 2))
                        promedio_otro = promedio_otro + IIf(IsNull(Math.Round(rs_aux27!bono_transporte, 2)), 0, Math.Round(rs_aux27!bono_transporte, 2))
                        promedio_totalg = promedio_totalg + IIf(IsNull(Math.Round(rs_aux27!total_ganado, 2)), 0, Math.Round(rs_aux27!total_ganado, 2))
                    End If
                End If
                num_promedio = num_promedio + 1
                fecha_aux = "01/10/" & rs_datos!ges_gestion
                If Str(rs_aux6!fecha_ingreso) = fecha_aux And num_promedio = 12 Then
                ' promedio_haber = promedio_haber + Math.Round(rs_aux6!beneficiario_haber_mensual, 2)
                End If
            Wend
            NUM_PERS = NUM_PERS + 1
        'EMPIEZA A GUARDAR LOS DATOS GENERADOS
            rs_datos2.AddNew
            rs_datos2!cargo_codigo = rs_aux6!cargo_codigo
            rs_datos2!puesto_codigo = rs_aux6!puesto_codigo
            rs_datos2!Unidad = rs_aux6!unidad_codigo
           rs_datos2!fecha_ingreso = rs_aux6!fecha_ingreso
           rs_datos2!fecha_expiracion = rs_aux6!fecha_expiracion
          rs_datos2!horas_extras = Math.Round(fun_dias360(rs_aux6!fecha_ingreso, rs_aux6!fecha_expiracion), 2)
  
          fecha_aux2 = "31/12/" & rs_datos!ges_gestion
          If rs_datos2!horas_extras = 12 Then
          rs_datos2!sueldo_basico = Math.Round(promedio_haber / 3, 2)
          rs_datos2!bono_antiguedad = Math.Round(promedio_bono / 3, 2)
          rs_datos2!bono_transporte = promedio_otro / 3
          rs_datos2!total_ganado = Math.Round(rs_datos2!sueldo_basico + rs_datos2!bono_antiguedad + rs_datos2!bono_transporte, 2)
          rs_datos2!beneficiario_codigo = rs_aux6!beneficiario_codigo
          Else
'          If rs_aux6!beneficiario_codigo = "4333735" Then
'          MsgBox ""
'          End If
            If Str(rs_aux6!fecha_expiracion) < fecha_aux2 Then
            Dim dia, mes, contador  As Integer
            contador = 0
            promedio_haber = 0
            promedio_bono = 0
            promedio_otro = 0
            promedio_totalg = 0
            dia = Day(rs_aux6!fecha_expiracion)
            mes = Month(rs_aux6!fecha_expiracion)
            If dia > 1 Then
            mes = mes - 4
            Else
            mes = mes - 3
            End If
             While contador <> 3
            'Dim rs_aux27 As New ADODB.Recordset
            If rs_aux27.State = 1 Then rs_aux27.Close
            rs_aux27.Open "select * from ro_pagos_cronograma_detalle where ges_gestion = '" & rs_datos!ges_gestion & "' AND beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "' AND mes_grupo = " & mes & "", db, adOpenKeyset, adLockReadOnly  ', adOpenKeyset, adOpenStatic, adCmdText
            If rs_aux27.RecordCount > 0 Then
            
            promedio_haber = promedio_haber + Math.Round(rs_aux27!sueldo_basico, 2)
            promedio_bono = promedio_bono + IIf(IsNull(Math.Round(rs_aux27!bono_antiguedad, 2)), 0, Math.Round(rs_aux27!bono_antiguedad, 2))
            promedio_otro = promedio_otro + IIf(IsNull(Math.Round(rs_aux27!bono_transporte, 2)), 0, Math.Round(rs_aux27!bono_transporte, 2))
            promedio_totalg = promedio_totalg + IIf(IsNull(Math.Round(rs_aux27!total_ganado, 2)), 0, Math.Round(rs_aux27!total_ganado, 2))
            End If
            contador = contador + 1
            mes = mes + 1
         Wend
         
        NUM_PERS = NUM_PERS + 1
        End If
          rs_datos2!sueldo_basico = Math.Round(promedio_haber / 3, 2)
          rs_datos2!bono_antiguedad = Math.Round(promedio_bono / 3, 2)
          rs_datos2!bono_transporte = promedio_otro / 3
          rs_datos2!total_ganado = Math.Round(rs_datos2!sueldo_basico + rs_datos2!bono_antiguedad + rs_datos2!bono_transporte, 2)
          rs_datos2!beneficiario_codigo = rs_aux6!beneficiario_codigo
                    
          End If
          
          rs_datos2!liquido_pagable_bs = (rs_datos2!total_ganado * rs_datos2!horas_extras) / 12
          
          rs_datos2!ges_gestion = rs_datos!ges_gestion
          rs_datos2!planilla_codigo = rs_datos!planilla_codigo
          rs_datos2!mes_grupo = rs_datos1!mes_grupo
          rs_datos2!NUMERO_PAGO = rs_datos1!NUMERO_PAGO
          rs_datos2!unidad_codigo = rs_aux6!unidad_codigo_pla
          rs_datos2!tipo_moneda = "BOB"
          rs_datos2!tipo_cambio = GlTipoCambioOficial
            'Adicionar  beneficiario_haber_mensual_ant
'
            DIA_IN = Day(rs_aux6!fecha_ingreso)
            MES_IN = Month(rs_aux6!fecha_ingreso)
            
            ANO_IN = Year(rs_aux6!fecha_ingreso)
              DIA_HOY = Day(Now)
            
            
            'rs_datos2!sueldo_basico = rs_aux6!beneficiario_haber_mensual
            
            rs_datos2!Numero_consultoriaHist = rs_aux6!beneficiario_item
            'rc_antiguedad

            rs_datos2!emite_factura = "N"
             
            rs_datos2!cite_conformidad = "-"
            'rs_datos2!Numero_consultoriaHist = " "
            rs_datos2!fte_financiamientoHist = "-"
            rs_datos2!estado_devengado = "REG"
            '70522788
            rs_datos2!estado_codigo = "REG"
            rs_datos2!fecha_registro = Date
            rs_datos2!usr_codigo = glusuario
            
            rs_datos2!iva_110 = "0"
            rs_datos2!fisco_a_favor = "0"
            rs_datos2!dependiente_a_favor = "0"
            rs_datos2!mes_anterior = "0"
            rs_datos2!mes_anterior_mant = "0"
            rs_datos2!saldo_util = "0"
            rs_datos2!saldo_a_favor_depend = "0"
            rs_datos2!rciva = "0"
             rs_datos2!codigo_empresa = rs_aux6!codigo_empresa
            rs_datos2!hora_registro = Format(Time, "HH:mm:ss")
            'ABRIR_TABLA_DET (2)
            If rs_aux6!estado_jubilado = "APR" Then
                VAR_AFPJ = 0.0271
            Else
                VAR_AFPJ = 0.1271
            End If
            Select Case rs_aux6!beneficiario_codigo_afp
            Case "1006803"      'AFP1 FUTURO
                rs_datos2!afp1 = rs_datos2!total_ganado * VAR_AFPJ
                rs_datos2!afp2 = "0"       'falta 987654
                VAR_NETO = rs_datos2!total_ganado - rs_datos2!afp1
            Case "987654"       'AFP2 PREVISION
                rs_datos2!afp1 = "0"       'falta 1006803
                rs_datos2!afp2 = rs_datos2!total_ganado * VAR_AFPJ
                VAR_NETO = rs_datos2!total_ganado - rs_datos2!afp2
            Case Else
                rs_datos2!afp1 = "0"
                rs_datos2!afp2 = "0"
                VAR_NETO = rs_datos2!total_ganado
        End Select
        
            rs_datos2.Update
            'Call OptFilGral1_Click
            'rs_datos.MoveLast
            mbDataChanged = False
    '
        End If
    Else 'PARA LAS BAJAS
    rs_aux6!estado_codigo = "ANL"
    End If 'PARA LAS BAJAS
        promedio_haber = 0
        promedio_bono = 0
        promedio_otro = 0
        promedio_totalg = 0
         num_promedio = 9
          
        rs_aux6.MoveNext
'           If rs_aux6!beneficiario_codigo = "3518716" Then '"4333735"
'            sino = ""
'            End If
       Wend
  End If 'verifica si existe personal en esa sub_planilla
       rs_aux16.MoveNext
     Wend
       
       Call ABRIR_TABLA_DET(2)
       Call ABRIR_TABLAS_AUX(5)
       Call numeracion_planilla
       'rs_datos2.RecordCount
       
   'sino = MsgBox("Se genero correctamente la planilla", vbInformation, "Atención")
    continuar = "SI"
    ProgressBar1.Visible = False
    dtc_buscar_desc.Visible = True
    Label52.Visible = True

    
    Fra_personal_Ppla.Visible = False
    FraNavega.Enabled = True
    fraOpciones.Enabled = True
    ' FraGrabarCancelar.Visible = True
    dg_datos.Enabled = True
    dg_det1.Enabled = True
    fra_opciones_det_1.Enabled = True
    fra_opciones_det_2.Enabled = True
   
End Sub
Public Sub numeracion_planilla()
Dim cont As Integer
cont = 0

Dim rs_aux32 As New ADODB.Recordset
If rs_aux32.State = 1 Then rs_aux32.Close
      'SE NUMERA A TODA LA PLANILLA MENOS EL PERSONAL A PRUEBA
      rs_aux32.Open "select * from gc_empresas where estado_codigo = 'APR'", db, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText 'adOpenKeyset ', adLockReadOnly   ', adOpenKeyset, adOpenStatic, adCmdText" 'adOpenKeyset, adLockReadOnly   ', adOpenKeyset, adOpenStatic, adCmdText
      If rs_aux32.RecordCount > 0 Then
      rs_aux32.MoveFirst
      While Not rs_aux32.EOF
      cont = 0
      
Dim rs_numeracion As New ADODB.Recordset
If rs_numeracion.State = 1 Then rs_numeracion.Close
      'SE NUMERA A TODA LA PLANILLA MENOS EL PERSONAL A PRUEBA
      rs_numeracion.Open "select * from ro_pagos_cronograma_detalle where ges_gestion = '" & rs_datos!ges_gestion & "' AND mes_grupo = " & rs_datos!mes_grupo & " AND unidad_codigo <> 'P010' AND unidad_codigo <> 'P020' AND unidad_codigo <> 'P030' AND unidad_codigo <> 'P040' AND unidad_codigo <> 'P050' AND unidad_codigo <> 'P060' AND unidad_codigo <> 'P070' AND unidad_codigo <> 'P080' AND unidad_codigo <> 'P090' AND codigo_empresa = " & rs_aux32!codigo_empresa & " ORDER BY planilla_codigo, unidad_codigo, total_ganado DESC, beneficiario_codigo", db, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText 'adOpenKeyset ', adLockReadOnly   ', adOpenKeyset, adOpenStatic, adCmdText" 'adOpenKeyset, adLockReadOnly   ', adOpenKeyset, adOpenStatic, adCmdText
      If rs_numeracion.RecordCount > 0 Then
      rs_numeracion.MoveFirst
      While Not rs_numeracion.EOF
'
'      If rs_numeracion!beneficiario_codigo = "5809459" Then
'      MsgBox ("")
'      Else
'
'      End If
      cont = cont + 1
      rs_numeracion!Numero_consultoriaHist = cont

      rs_numeracion.Update
      rs_numeracion.MoveNext
      Wend
      'rs_numeracion.Update
      If Ado_datos.Recordset!mes_grupo < 13 Then
      db.Execute "update ro_personal_contratado set ro_personal_contratado.beneficiario_item = '0'"
      db.Execute "update ro_personal_contratado set ro_personal_contratado.beneficiario_item = ro_pagos_cronograma_detalle.Numero_consultoriaHist FROM ro_pagos_cronograma_detalle where ro_pagos_cronograma_detalle.beneficiario_codigo = ro_personal_contratado.beneficiario_codigo AND ro_pagos_cronograma_Detalle.mes_grupo =" & rs_datos!mes_grupo & "AND ro_pagos_cronograma_Detalle.ges_gestion ='" & rs_datos!ges_gestion & "' and ro_pagos_cronograma_Detalle.codigo_empresa = " & rs_aux32!codigo_empresa & ""
      db.Execute "update ro_personal_contratado set ro_personal_contratado.bono_antiguedad = ro_pagos_cronograma_detalle.bono_antiguedad FROM ro_pagos_cronograma_detalle where ro_pagos_cronograma_detalle.beneficiario_codigo = ro_personal_contratado.beneficiario_codigo AND ro_pagos_cronograma_Detalle.mes_grupo =" & rs_datos!mes_grupo & "AND ro_pagos_cronograma_Detalle.ges_gestion ='" & rs_datos!ges_gestion & "' and ro_pagos_cronograma_Detalle.codigo_empresa = " & rs_aux32!codigo_empresa & ""
      End If
      End If
      
rs_aux32.MoveNext
Wend
End If
End Sub
Private Sub RETROACTIVO()
Dim meses As Integer
Dim promedio As Double
meses = 0
While meses <> 4
meses = meses + 1
Set retroactivos = New Recordset
If retroactivos.State = 1 Then retroactivos.Close
      
      retroactivos.Open "select * from ro_pagos_cronograma_detalle where ges_gestion = '" & "2018" & "' AND mes_grupo = " & meses & " AND unidad_codigo <> 'P010' AND unidad_codigo <> 'P020' AND unidad_codigo <> 'P030' AND unidad_codigo <> 'P040' AND unidad_codigo <> 'P050' AND unidad_codigo <> 'P060' AND unidad_codigo <> 'P070' AND unidad_codigo <> 'P080' AND unidad_codigo <> 'P090' ORDER BY planilla_codigo, unidad_codigo, total_ganado DESC, beneficiario_codigo", db, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText 'adOpenKeyset ', adLockReadOnly   ', adOpenKeyset, adOpenStatic, adCmdText" 'adOpenKeyset, adLockReadOnly   ', adOpenKeyset, adOpenStatic, adCmdText
      If retroactivos.RecordCount > 0 Then
      retroactivos.MoveFirst
      
      While Not retroactivos.EOF
      
             
      
      Set rs_aux_retro = New Recordset
      If rs_aux_retro.State = 1 Then rs_aux_retro.Close
      rs_aux_retro.Open "SELECT * FROM ro_retroactivo_aux WHERE beneficiario_codigo ='" & retroactivos!beneficiario_codigo & "' and ges_gestion = '" & "2018" & "'", db, adOpenKeyset, adLockOptimistic
      

      
      If rs_aux_retro.RecordCount = 0 Then
      rs_aux_retro.AddNew
      rs_aux_retro!ges_gestion = retroactivos!ges_gestion
      rs_aux_retro!beneficiario_codigo = retroactivos!beneficiario_codigo
      
      
      End If
  
      
      Select Case meses
      Case "1"
      
      If rs_aux_retro!incremento <> "NO" Then
        If retroactivos!beneficiario_haber_mensual <= 2000 Then
        
          If retroactivos!dias_trabajados < 30 Then
            rs_aux_retro!ENERO = ((2060 - retroactivos!beneficiario_haber_mensual) / 30) * retroactivos!dias_trabajados
          Else
           rs_aux_retro!ENERO = 2060 - retroactivos!beneficiario_haber_mensual
          End If
          
            rs_aux_retro!haber_basico_incre = "2060"
            rs_aux_retro!haber_basico_incre_enero = "2060"
            rs_aux_retro!haber_basico_ant = retroactivos!beneficiario_haber_mensual
        Else
             
          If retroactivos!dias_trabajados < 30 Then
            rs_aux_retro!ENERO = ((retroactivos!beneficiario_haber_mensual * 0.055) / 30) * retroactivos!dias_trabajados
          Else
           rs_aux_retro!ENERO = retroactivos!beneficiario_haber_mensual * 0.055
          End If
          
            rs_aux_retro!haber_basico_ant = retroactivos!beneficiario_haber_mensual
            rs_aux_retro!haber_basico_incre = rs_aux_retro!haber_basico_ant + (rs_aux_retro!haber_basico_ant * 0.055)
            rs_aux_retro!haber_basico_incre_enero = rs_aux_retro!haber_basico_ant + (rs_aux_retro!haber_basico_ant * 0.055)
            
        End If
      Else
            rs_aux_retro!haber_basico_ant = retroactivos!beneficiario_haber_mensual
            rs_aux_retro!ENERO = 0
            rs_aux_retro!haber_basico_incre = rs_aux_retro!haber_basico_ant
            rs_aux_retro!haber_basico_incre_enero = rs_aux_retro!haber_basico_ant
      End If
        
     
      
       
       Set rs_aux6 = New Recordset
       If rs_aux6.State = 1 Then rs_aux6.Close
        'ABRE LAS PERSONAL CONTRATADAS VIGENTES
       rs_aux6.Open "SELECT * FROM ro_personal_contratado WHERE beneficiario_codigo ='" & retroactivos!beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic


            fecha_pla = DateSerial(rs_datos!ges_gestion, meses + 1, 1 - 1)
            
             'ABRE TABLA DONDE ESTAN LOS PARAMETROS DE ANTIGUEDAD
             If rs_aux8.State = 1 Then rs_aux8.Close
             rs_aux8.Open "select * from rc_antiguedad", db, adOpenKeyset, adLockOptimistic, adCmdText
             rs_aux8.MoveFirst
             While Not rs_aux8.EOF
             'GUARDA LAS FECHA MINIMA Y LA MAXIMA EN LA QUE DEBERIA ENTRAR LA PERSONA PARA RECIBIR EL BONO ANTIGUEDAD
             f1 = DateAdd("yyyy", -rs_aux8!parametro_inicial, CDate(fecha_pla))
             f2 = DateAdd("yyyy", -rs_aux8!parametro_final, CDate(fecha_pla))
             'COMPARA LA FECHA INGRESO CON EL PARAMETRO
             If rs_aux6!fecha_ingreso <= CDate(f1) And rs_aux6!fecha_ingreso > CDate(f2) Then
             'GUARDA EL MONTO QUE CORRESPONDE
             rs_aux_retro!bono_antiguedad = rs_aux8!antig_valor - retroactivos!bono_antiguedad
             rs_aux8.MoveLast
             End If
             rs_aux8.MoveNext
             Wend
      
        rs_aux_retro!total_ganado = IIf(IsNull(rs_aux_retro!ENERO), 0, rs_aux_retro!ENERO) + IIf(IsNull(rs_aux_retro!bono_antiguedad), 0, rs_aux_retro!bono_antiguedad)
        rs_aux_retro!total_ganado_enero = IIf(IsNull(rs_aux_retro!ENERO), 0, rs_aux_retro!ENERO) + IIf(IsNull(rs_aux_retro!bono_antiguedad), 0, rs_aux_retro!bono_antiguedad)
            If rs_aux6!estado_jubilado = "APR" Then
                VAR_AFPJ = 0.0271
            Else
                VAR_AFPJ = 0.1271
            End If
        Select Case rs_aux6!beneficiario_codigo_afp
                Case "1006803"      'AFP1
                    rs_aux_retro!afp1 = rs_aux_retro!total_ganado * VAR_AFPJ
                    rs_aux_retro!afp2 = "0"       'falta 987654
                    
                     rs_aux_retro!afp1_enero = rs_aux_retro!total_ganado_enero * VAR_AFPJ
                     rs_aux_retro!afp2_enero = "0"
                    
                    'VAR_NETO = rs_aux_retro!total_ganado - rs_datos2!afp1
                     'rs_aux_retro!total_dsctos = rs_aux_retro!afp1
                Case "987654"       'AFP2
                    rs_aux_retro!afp1 = "0"       'falta 1006803
                    rs_aux_retro!afp2 = rs_aux_retro!total_ganado * VAR_AFPJ
                    'rs_aux_retro!total_dsctos = rs_aux_retro!afp2
                     
                    rs_aux_retro!afp1_enero = "0"       'falta 1006803
                    rs_aux_retro!afp2_enero = rs_aux_retro!total_ganado_enero * VAR_AFPJ
                     
                    'VAR_NETO = rs_aux_retro!total_ganado - rs_aux_retro!afp2
                Case Else
                    rs_aux_retro!afp1 = "0"
                    rs_aux_retro!afp2 = "0"
                    rs_aux_retro!afp1_enero = "0"
                    rs_aux_retro!afp2_enero = "0"
                    'VAR_NETO = rs_aux_retro!total_ganado
      End Select
        
        rs_aux_retro!total_dsctos = IIf(IsNull(rs_aux_retro!afp1_enero), 0, rs_aux_retro!afp1_enero) + IIf(IsNull(rs_aux_retro!afp2_enero), 0, rs_aux_retro!afp2_enero)
        
        
      rs_aux_retro!liquido_pagable_bs = rs_aux_retro!total_ganado - rs_aux_retro!total_dsctos
      'rs_aux_retro!haber_basico_incre = rs_aux_retro!haber_basico_ant + (rs_aux_retro!haber_basico_ant * 0.7)
      
      
      Case "2"
    If rs_aux_retro!incremento <> "NO" Then

        If retroactivos!beneficiario_haber_mensual <= 2000 Then
        
          If retroactivos!dias_trabajados < 30 Then
            rs_aux_retro!febrero = ((2060 - retroactivos!beneficiario_haber_mensual) / 30) * retroactivos!dias_trabajados
          Else
           rs_aux_retro!febrero = 2060 - retroactivos!beneficiario_haber_mensual
          End If
          
         
         rs_aux_retro!haber_basico_incre = "2060"
         rs_aux_retro!haber_basico_incre_febrero = "2060"
         rs_aux_retro!haber_basico_ant = retroactivos!beneficiario_haber_mensual
        Else
        
           If retroactivos!dias_trabajados < 30 Then
            rs_aux_retro!febrero = ((retroactivos!beneficiario_haber_mensual * 0.055) / 30) * retroactivos!dias_trabajados
          Else
            rs_aux_retro!febrero = retroactivos!beneficiario_haber_mensual * 0.055
          End If
          
        
        rs_aux_retro!haber_basico_ant = retroactivos!beneficiario_haber_mensual
        rs_aux_retro!haber_basico_incre = rs_aux_retro!haber_basico_ant + (rs_aux_retro!haber_basico_ant * 0.055)
        rs_aux_retro!haber_basico_incre_febrero = rs_aux_retro!haber_basico_ant + (rs_aux_retro!haber_basico_ant * 0.055)
        End If
        
    Else
            rs_aux_retro!haber_basico_ant = retroactivos!beneficiario_haber_mensual
            rs_aux_retro!febrero = 0
            rs_aux_retro!haber_basico_incre = rs_aux_retro!haber_basico_ant
            rs_aux_retro!haber_basico_incre_febrero = rs_aux_retro!haber_basico_ant
    End If
     
         
                Set rs_aux6 = New Recordset
       If rs_aux6.State = 1 Then rs_aux6.Close
        'ABRE LAS PERSONAL CONTRATADAS VIGENTES
       rs_aux6.Open "SELECT * FROM ro_personal_contratado WHERE beneficiario_codigo ='" & retroactivos!beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic


            fecha_pla = DateSerial(rs_datos!ges_gestion, meses + 1, 1 - 1)
            
             'ABRE TABLA DONDE ESTAN LOS PARAMETROS DE ANTIGUEDAD
             If rs_aux8.State = 1 Then rs_aux8.Close
             rs_aux8.Open "select * from rc_antiguedad", db, adOpenKeyset, adLockOptimistic, adCmdText
             rs_aux8.MoveFirst
             While Not rs_aux8.EOF
             'GUARDA LAS FECHA MINIMA Y LA MAXIMA EN LA QUE DEBERIA ENTRAR LA PERSONA PARA RECIBIR EL BONO ANTIGUEDAD
             f1 = DateAdd("yyyy", -rs_aux8!parametro_inicial, CDate(fecha_pla))
             f2 = DateAdd("yyyy", -rs_aux8!parametro_final, CDate(fecha_pla))
             'COMPARA LA FECHA INGRESO CON EL PARAMETRO
             If rs_aux6!fecha_ingreso <= CDate(f1) And rs_aux6!fecha_ingreso > CDate(f2) Then
             'GUARDA EL MONTO QUE CORRESPONDE
             rs_aux_retro!bono_antiguedad2 = rs_aux8!antig_valor - retroactivos!bono_antiguedad
             rs_aux8.MoveLast
             End If
             rs_aux8.MoveNext
             Wend
      
      
            rs_aux_retro!total_ganado = IIf(IsNull(rs_aux_retro!ENERO), 0, rs_aux_retro!ENERO) + IIf(IsNull(rs_aux_retro!febrero), 0, rs_aux_retro!febrero) + IIf(IsNull(rs_aux_retro!bono_antiguedad), 0, rs_aux_retro!bono_antiguedad) + IIf(IsNull(rs_aux_retro!bono_antiguedad2), 0, rs_aux_retro!bono_antiguedad2)
            rs_aux_retro!total_ganado_febrero = IIf(IsNull(rs_aux_retro!febrero), 0, rs_aux_retro!febrero) + IIf(IsNull(rs_aux_retro!bono_antiguedad2), 0, rs_aux_retro!bono_antiguedad2)
            If rs_aux6!estado_jubilado = "APR" Then
                VAR_AFPJ = 0.0271
            Else
                VAR_AFPJ = 0.1271
            End If
            Select Case rs_aux6!beneficiario_codigo_afp
                Case "1006803"      'AFP1
                    rs_aux_retro!afp1 = rs_aux_retro!total_ganado * VAR_AFPJ
                    rs_aux_retro!afp2 = "0"       'falta 987654
                    
                    rs_aux_retro!afp1_febrero = rs_aux_retro!total_ganado_febrero * VAR_AFPJ
                    
                    'VAR_NETO = rs_aux_retro!total_ganado - rs_datos2!afp1
                     'rs_aux_retro!total_dsctos = rs_aux_retro!afp1
                Case "987654"       'AFP2
                    rs_aux_retro!afp1 = "0"       'falta 1006803
                    rs_aux_retro!afp2 = rs_aux_retro!total_ganado * VAR_AFPJ
                    
                   ' rs_aux_retro!total_dsctos = rs_aux_retro!afp2
                     
                    rs_aux_retro!afp1_febrero = "0"       'falta 1006803
                    rs_aux_retro!afp2_febrero = rs_aux_retro!total_ganado_febrero * VAR_AFPJ
                    'VAR_NETO = rs_aux_retro!total_ganado - rs_aux_retro!afp2
                Case Else
                    rs_aux_retro!afp1 = "0"
                    rs_aux_retro!afp2 = "0"
                    
                    rs_aux_retro!afp1_febrero = "0"
                    rs_aux_retro!afp2_febrero = "0"
                    'VAR_NETO = rs_aux_retro!total_ganado
      End Select
      rs_aux_retro!total_dsctos = IIf(IsNull(rs_aux_retro!afp1_enero), 0, rs_aux_retro!afp1_enero) + IIf(IsNull(afp2_enero), 0, rs_aux_retro!afp2_enero) + IIf(IsNull(rs_aux_retro!afp1_febrero), 0, rs_aux_retro!afp1_febrero) + IIf(IsNull(rs_aux_retro!afp2_febrero), 0, rs_aux_retro!afp2_febrero)
      rs_aux_retro!liquido_pagable_bs = rs_aux_retro!total_ganado - rs_aux_retro!total_dsctos
      'rs_aux_retro!haber_basico_incre = rs_aux_retro!haber_basico_ant + (rs_aux_retro!haber_basico_ant * 0.7)
      
      Case "3"
      
If rs_aux_retro!incremento <> "NO" Then
     
 If retroactivos!beneficiario_codigo = "8706813" Then
 retroactivos!beneficiario_codigo = "8706813"
 
 End If
      If retroactivos!beneficiario_haber_mensual <= 2000 Then
      
          If retroactivos!dias_trabajados < 30 Then
            rs_aux_retro!marzo = ((2060 - retroactivos!beneficiario_haber_mensual) / 30) * retroactivos!dias_trabajados
          Else
            rs_aux_retro!marzo = 2060 - retroactivos!beneficiario_haber_mensual
          End If
      
      
      rs_aux_retro!haber_basico_incre = "2060"
      rs_aux_retro!haber_basico_incre_marzo = "2060"
      rs_aux_retro!haber_basico_ant = retroactivos!beneficiario_haber_mensual
      Else
      
           If retroactivos!dias_trabajados < 30 Then
            rs_aux_retro!marzo = ((retroactivos!beneficiario_haber_mensual * 0.055) / 30) * retroactivos!dias_trabajados
          Else
            rs_aux_retro!marzo = retroactivos!beneficiario_haber_mensual * 0.055
          End If
      
      
      rs_aux_retro!haber_basico_ant = retroactivos!beneficiario_haber_mensual
      rs_aux_retro!haber_basico_incre = rs_aux_retro!haber_basico_ant + (rs_aux_retro!haber_basico_ant * 0.055)
      rs_aux_retro!haber_basico_incre_marzo = rs_aux_retro!haber_basico_ant + (rs_aux_retro!haber_basico_ant * 0.055)
      End If
Else
            rs_aux_retro!haber_basico_ant = retroactivos!beneficiario_haber_mensual
            rs_aux_retro!marzo = 0
            rs_aux_retro!haber_basico_incre = rs_aux_retro!haber_basico_ant
            rs_aux_retro!haber_basico_incre_marzo = rs_aux_retro!haber_basico_ant
End If
              
   

      
                Set rs_aux6 = New Recordset
       If rs_aux6.State = 1 Then rs_aux6.Close
        'ABRE LAS PERSONAL CONTRATADAS VIGENTES
       rs_aux6.Open "SELECT * FROM ro_personal_contratado WHERE beneficiario_codigo ='" & retroactivos!beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic


            fecha_pla = DateSerial(rs_datos!ges_gestion, meses + 1, 1 - 1)
            
             'ABRE TABLA DONDE ESTAN LOS PARAMETROS DE ANTIGUEDAD
             If rs_aux8.State = 1 Then rs_aux8.Close
             rs_aux8.Open "select * from rc_antiguedad", db, adOpenKeyset, adLockOptimistic, adCmdText
             rs_aux8.MoveFirst
             While Not rs_aux8.EOF
             'GUARDA LAS FECHA MINIMA Y LA MAXIMA EN LA QUE DEBERIA ENTRAR LA PERSONA PARA RECIBIR EL BONO ANTIGUEDAD
             f1 = DateAdd("yyyy", -rs_aux8!parametro_inicial, CDate(fecha_pla))
             f2 = DateAdd("yyyy", -rs_aux8!parametro_final, CDate(fecha_pla))
             'COMPARA LA FECHA INGRESO CON EL PARAMETRO
             If rs_aux6!fecha_ingreso <= CDate(f1) And rs_aux6!fecha_ingreso > CDate(f2) Then
             'GUARDA EL MONTO QUE CORRESPONDE
             rs_aux_retro!bono_antiguedad3 = rs_aux8!antig_valor - retroactivos!bono_antiguedad
             rs_aux8.MoveLast
             End If
             rs_aux8.MoveNext
             Wend
      
      
        rs_aux_retro!total_ganado = IIf(IsNull(rs_aux_retro!ENERO), 0, rs_aux_retro!ENERO) + IIf(IsNull(rs_aux_retro!febrero), 0, rs_aux_retro!febrero) + IIf(IsNull(rs_aux_retro!marzo), 0, rs_aux_retro!marzo) + IIf(IsNull(rs_aux_retro!bono_antiguedad), 0, rs_aux_retro!bono_antiguedad) + IIf(IsNull(rs_aux_retro!bono_antiguedad2), 0, rs_aux_retro!bono_antiguedad2) + IIf(IsNull(rs_aux_retro!bono_antiguedad3), 0, rs_aux_retro!bono_antiguedad3)
        rs_aux_retro!total_ganado_marzo = IIf(IsNull(rs_aux_retro!marzo), 0, rs_aux_retro!marzo) + IIf(IsNull(rs_aux_retro!bono_antiguedad3), 0, rs_aux_retro!bono_antiguedad3)
            If rs_aux6!estado_jubilado = "APR" Then
                VAR_AFPJ = 0.0271
            Else
                VAR_AFPJ = 0.1271
            End If
        Select Case rs_aux6!beneficiario_codigo_afp
                Case "1006803"      'AFP1
                    rs_aux_retro!afp1 = rs_aux_retro!total_ganado * VAR_AFPJ
                    rs_aux_retro!afp2 = "0"       'falta 987654
                    
                    rs_aux_retro!afp1_marzo = rs_aux_retro!total_ganado_marzo * VAR_AFPJ
                    
                    'VAR_NETO = rs_aux_retro!total_ganado - rs_datos2!afp1
                     'rs_aux_retro!total_dsctos = rs_aux_retro!afp1
                Case "987654"       'AFP2
                    rs_aux_retro!afp1 = "0"       'falta 1006803
                    rs_aux_retro!afp2 = rs_aux_retro!total_ganado * VAR_AFPJ
                    'rs_aux_retro!total_dsctos = rs_aux_retro!afp2
                     
                    rs_aux_retro!afp1_marzo = "0"
                    rs_aux_retro!afp2_marzo = rs_aux_retro!total_ganado_marzo * VAR_AFPJ
                     
                    'VAR_NETO = rs_aux_retro!total_ganado - rs_aux_retro!afp2
                Case Else
                    rs_aux_retro!afp1 = "0"
                    rs_aux_retro!afp2 = "0"
                    
                    rs_aux_retro!afp1_marzo = "0"
                    rs_aux_retro!afp2_marzo = "0"
                    
                    'VAR_NETO = rs_aux_retro!total_ganado
      End Select
      rs_aux_retro!total_dsctos = IIf(IsNull(rs_aux_retro!afp1_enero), 0, rs_aux_retro!afp1_enero) + IIf(IsNull(rs_aux_retro!afp2_enero), 0, rs_aux_retro!afp2_enero) + IIf(IsNull(rs_aux_retro!afp1_febrero), 0, rs_aux_retro!afp1_febrero) + IIf(IsNull(rs_aux_retro!afp2_febrero), 0, rs_aux_retro!afp2_febrero) + IIf(IsNull(rs_aux_retro!afp1_marzo), 0, rs_aux_retro!afp1_marzo) + IIf(IsNull(rs_aux_retro!afp2_marzo), 0, rs_aux_retro!afp2_marzo)
      rs_aux_retro!liquido_pagable_bs = rs_aux_retro!total_ganado - rs_aux_retro!total_dsctos
      
      
     ' rs_aux_retro!haber_basico_incre = rs_aux_retro!haber_basico_ant + (rs_aux_retro!haber_basico_ant * 0.7)
      
      Case "4"
      


If rs_aux_retro!incremento <> "NO" Then

      If retroactivos!beneficiario_haber_mensual <= 2060 Then
      
          If retroactivos!dias_trabajados < 30 Then
            rs_aux_retro!abril = ((2060 - retroactivos!beneficiario_haber_mensual) / 30) * retroactivos!dias_trabajados
          Else
            rs_aux_retro!abril = 2060 - retroactivos!beneficiario_haber_mensual
          End If
      
      
      rs_aux_retro!haber_basico_incre = "2060"
      rs_aux_retro!haber_basico_incre_abril = "2060"
      rs_aux_retro!haber_basico_ant = retroactivos!beneficiario_haber_mensual
      Else
      
          If retroactivos!dias_trabajados < 30 Then
            rs_aux_retro!abril = ((retroactivos!beneficiario_haber_mensual * 0.055) / 30) * retroactivos!dias_trabajados
          Else
            rs_aux_retro!abril = retroactivos!beneficiario_haber_mensual * 0.055
          End If
      
      
      rs_aux_retro!haber_basico_ant = retroactivos!beneficiario_haber_mensual
      rs_aux_retro!haber_basico_incre = rs_aux_retro!haber_basico_ant + (rs_aux_retro!haber_basico_ant * 0.055)
      rs_aux_retro!haber_basico_incre_abril = rs_aux_retro!haber_basico_ant + (rs_aux_retro!haber_basico_ant * 0.055)
      End If
      
      
Else
            rs_aux_retro!haber_basico_ant = retroactivos!beneficiario_haber_mensual
            rs_aux_retro!abril = 0
            rs_aux_retro!haber_basico_incre = rs_aux_retro!haber_basico_ant
            rs_aux_retro!haber_basico_incre_abril = rs_aux_retro!haber_basico_ant
End If
             

                Set rs_aux6 = New Recordset
       If rs_aux6.State = 1 Then rs_aux6.Close
        'ABRE LAS PERSONAL CONTRATADAS VIGENTES
       rs_aux6.Open "SELECT * FROM ro_personal_contratado WHERE beneficiario_codigo ='" & retroactivos!beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic


            fecha_pla = DateSerial(rs_datos!ges_gestion, meses + 1, 1 - 1)
            
             'ABRE TABLA DONDE ESTAN LOS PARAMETROS DE ANTIGUEDAD
             If rs_aux8.State = 1 Then rs_aux8.Close
             rs_aux8.Open "select * from rc_antiguedad", db, adOpenKeyset, adLockOptimistic, adCmdText
             rs_aux8.MoveFirst
             While Not rs_aux8.EOF
             'GUARDA LAS FECHA MINIMA Y LA MAXIMA EN LA QUE DEBERIA ENTRAR LA PERSONA PARA RECIBIR EL BONO ANTIGUEDAD
             f1 = DateAdd("yyyy", -rs_aux8!parametro_inicial, CDate(fecha_pla))
             f2 = DateAdd("yyyy", -rs_aux8!parametro_final, CDate(fecha_pla))
             'COMPARA LA FECHA INGRESO CON EL PARAMETRO
             If rs_aux6!fecha_ingreso <= CDate(f1) And rs_aux6!fecha_ingreso > CDate(f2) Then
             'GUARDA EL MONTO QUE CORRESPONDE
             rs_aux_retro!bono_antiguedad4 = rs_aux8!antig_valor - retroactivos!bono_antiguedad
             rs_aux8.MoveLast
             End If
             rs_aux8.MoveNext
             Wend
      
      
      rs_aux_retro!total_ganado = IIf(IsNull(rs_aux_retro!ENERO), 0, rs_aux_retro!ENERO) + IIf(IsNull(rs_aux_retro!febrero), 0, rs_aux_retro!febrero) + IIf(IsNull(rs_aux_retro!marzo), 0, rs_aux_retro!marzo) + IIf(IsNull(rs_aux_retro!abril), 0, rs_aux_retro!abril) + IIf(IsNull(rs_aux_retro!bono_antiguedad), 0, rs_aux_retro!bono_antiguedad) + IIf(IsNull(rs_aux_retro!bono_antiguedad2), 0, rs_aux_retro!bono_antiguedad2) + IIf(IsNull(rs_aux_retro!bono_antiguedad3), 0, rs_aux_retro!bono_antiguedad3) + IIf(IsNull(rs_aux_retro!bono_antiguedad4), 0, rs_aux_retro!bono_antiguedad4)
      rs_aux_retro!total_ganado_abril = IIf(IsNull(rs_aux_retro!abril), 0, rs_aux_retro!abril) + IIf(IsNull(rs_aux_retro!bono_antiguedad4), 0, rs_aux_retro!bono_antiguedad4)
          Select Case rs_aux6!beneficiario_codigo_afp
                Case "1006803"      'AFP1
                    rs_aux_retro!afp1 = rs_aux_retro!total_ganado * 0.1271
                    rs_aux_retro!afp2 = "0"       'falta 987654
                    
                    rs_aux_retro!afp1_abril = rs_aux_retro!total_ganado_abril * 0.1271
                    rs_aux_retro!afp2_abril = "0"
                    
                    'VAR_NETO = rs_aux_retro!total_ganado - rs_datos2!afp1
                     'rs_aux_retro!total_dsctos = rs_aux_retro!afp1
                Case "987654"       'AFP2
                    rs_aux_retro!afp1 = "0"       'falta 1006803
                    rs_aux_retro!afp2 = rs_aux_retro!total_ganado * 0.1271
                  '  rs_aux_retro!total_dsctos = rs_aux_retro!afp2
                     
                    rs_aux_retro!afp1_abril = "0"
                    rs_aux_retro!afp2_abril = rs_aux_retro!total_ganado_abril * 0.1271
                     
                    'VAR_NETO = rs_aux_retro!total_ganado - rs_aux_retro!afp2
                Case Else
                    rs_aux_retro!afp1 = "0"
                    rs_aux_retro!afp2 = "0"
                    
                    rs_aux_retro!afp1_abril = "0"
                    rs_aux_retro!afp2_abril = "0"
                    
                    'VAR_NETO = rs_aux_retro!total_ganado
      End Select
      rs_aux_retro!total_dsctos = IIf(IsNull(rs_aux_retro!afp1_enero), 0, rs_aux_retro!afp1_enero) + IIf(IsNull(rs_aux_retro!afp2_enero), 0, rs_aux_retro!afp2_enero) + IIf(IsNull(rs_aux_retro!afp1_febrero), 0, rs_aux_retro!afp1_febrero) + IIf(IsNull(rs_aux_retro!afp2_febrero), 0, rs_aux_retro!afp2_febrero) + IIf(IsNull(rs_aux_retro!afp1_marzo), 0, rs_aux_retro!afp1_marzo) + IIf(IsNull(rs_aux_retro!afp2_marzo), 0, rs_aux_retro!afp2_marzo) + IIf(IsNull(rs_aux_retro!afp1_abril), 0, rs_aux_retro!afp1_abril) + IIf(IsNull(rs_aux_retro!afp2_abril), 0, rs_aux_retro!afp2_abril)
      rs_aux_retro!liquido_pagable_bs = rs_aux_retro!total_ganado - rs_aux_retro!total_dsctos
      'rs_aux_retro!haber_basico_incre = rs_aux_retro!haber_basico_ant + (rs_aux_retro!haber_basico_ant * 0.7)
      End Select
      
      

      rs_aux_retro.Update
      retroactivos.MoveNext

      Wend
    End If
   Wend
End Sub
Private Sub generar_rc_iva()

If rs_aux23.State = 1 Then rs_aux23.Close
      rs_aux23.Open "select * from rc_basico_nacional where estado_codigo = 'APR'", db, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText 'adOpenKeyset ', adLockReadOnly   ', adOpenKeyset, adOpenStatic, adCmdText" 'adOpenKeyset, adLockReadOnly   ', adOpenKeyset, adOpenStatic, adCmdText
        

Dim neto, dif_imp, impuest, smn As Double
 If rs_aux18.State = 1 Then rs_aux18.Close
      rs_aux18.Open "select * from ro_pagos_cronograma_detalle where ges_gestion = '" & rs_datos!ges_gestion & "' AND mes_grupo = " & rs_datos!mes_grupo, db, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText  'adOpenKeyset ', adLockReadOnly   ', adOpenKeyset, adOpenStatic, adCmdText" 'adOpenKeyset, adLockReadOnly   ', adOpenKeyset, adOpenStatic, adCmdText
      rs_aux18.MoveFirst
      
      With ProgressBar1
        .Max = rs_aux18.RecordCount
        .Min = 0
        .Value = 0
       End With
       ProgressBar1.Visible = True
While Not rs_aux18.EOF
 'If rs_aux22.State = 1 Then rs_aux22.Close
 ' rs_aux22.Open "select * from ro_pagos_cronograma_detalle where ges_gestion = '" & rs_datos!ges_gestion & "' AND mes_grupo = " & rs_datos!mes_grupo & " AND beneficiario_codigo = '" & rs_aux18!beneficiario_codigo & "'", db, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText 'adOpenKeyset ', adLockReadOnly   ', adOpenKeyset, adOpenStatic, adCmdText"
      
'If rs_aux18!beneficiario_codigo = "159256" Then
'sino = ""
'End If

ProgressBar1.Value = ProgressBar1.Value + 1
rs_aux18!sueldo_neto = Round((rs_aux18!total_ganado - rs_aux18!afp1 - rs_aux18!afp2), 2)
'------------------------------------------------------------------------->  esto solo es para preuebas
'If rs_aux18!beneficiario_codigo = "3361040" Then
'rs_aux18!sueldo_neto = "1805" ' esto solo es para preuebas
'rs_aux18!iva_110 = "1000" ' esto solo es para preuebas
'sino = ""
'End If
'-------------------------------------------------------------------------> esto solo es para preuebas
If rs_aux18!sueldo_neto > rs_aux23!bn_monto * 2 Then
rs_aux18!minimo_imponible = Round((rs_aux23!bn_monto * 2), 2)
rs_aux18!smn_iva = rs_aux23!bn_monto

Else
rs_aux18!minimo_imponible = rs_aux18!sueldo_neto
End If

If rs_aux18!sueldo_neto - rs_aux18!minimo_imponible > 0 Then

rs_aux18!dif_impuesto = Round((rs_aux18!sueldo_neto - rs_aux18!minimo_imponible), 2)
rs_aux18!impuesto_13 = Round((rs_aux18!dif_impuesto * 0.13), 2) '15/07/2017 de redondeo 0 a 2
'If rs_aux18!sueldo_neto >= rs_aux23!bn_monto * rs_aux23!bn_numero_sueldos Then ' minimo para entrar al rc iva
'iva110

rs_aux18!smn_2_13 = Round(((rs_aux23!bn_monto * 2) * 0.13), 0)
If rs_aux18!smn_2_13 > rs_aux18!impuesto_13 Then
rs_aux18!smn_2_13 = Round((rs_aux18!impuesto_13), 2) '31/07/2017 de redondeo 0 a 2
End If
 
If rs_aux18!impuesto_13 - rs_aux18!iva_110 - rs_aux18!smn_2_13 >= 0 Then
rs_aux18!fisco_a_favor = Round(rs_aux18!impuesto_13 - rs_aux18!iva_110 - rs_aux18!smn_2_13, 0)
Else
rs_aux18!fisco_a_favor = 0
End If
If rs_aux18!iva_110 + rs_aux18!smn_2_13 - rs_aux18!impuesto_13 >= 0 Then
rs_aux18!dependiente_a_favor = Round(((rs_aux18!iva_110 + rs_aux18!smn_2_13) - rs_aux18!impuesto_13), 0)
Else
rs_aux18!dependiente_a_favor = 0
End If
Dim mes_ante As Integer
Dim ges_ant As String
mes_ante = rs_datos!mes_grupo - 1
ges_ant = rs_datos!ges_gestion
If mes_ante = 0 Then
mes_ante = 12
ges_ant = rs_datos!ges_gestion - 1
End If
 If rs_aux19.State = 1 Then rs_aux19.Close
rs_aux19.Open "select * from ro_pagos_cronograma_detalle where ges_gestion = '" & ges_ant & "' AND mes_grupo = " & mes_ante & " AND beneficiario_codigo = '" & rs_aux18!beneficiario_codigo & "'", db, adOpenKeyset, adLockReadOnly  ', adOpenKeyset, adOpenStatic, adCmdText
If rs_aux19.RecordCount > 0 Then
rs_aux18!mes_anterior = IIf(IsNull(rs_aux19!saldo_para_mes_sig), 0, rs_aux19!saldo_para_mes_sig)
Else
rs_aux18!mes_anterior = 0
End If
'ufv
If rs_aux20.State = 1 Then rs_aux20.Close
rs_aux20.Open "select * from gc_tipo_cambio where fecha_cambio = '" & DTP_ufv_ant.Value & "'", db, adOpenKeyset, adLockReadOnly  ', adOpenKeyset, adOpenStatic, adCmdText
If rs_aux21.State = 1 Then rs_aux21.Close
rs_aux21.Open "select * from gc_tipo_cambio where fecha_cambio = '" & DTC_ufv_actual.Value & "'", db, adOpenKeyset, adLockReadOnly  ', adOpenKeyset, adOpenStatic, adCmdText
'actualizacion
rs_aux18!actualizacion = Round((((rs_aux21!cambio_ufv / rs_aux20!cambio_ufv) - 1) * rs_aux18!mes_anterior), 0)
rs_aux18!total = Round(rs_aux18!mes_anterior + rs_aux18!actualizacion, 0)
rs_aux18!saldo_a_favor_depend = Round(rs_aux18!dependiente_a_favor + rs_aux18!total, 0)
rs_aux18!ufv_1 = rs_aux21!cambio_ufv
rs_aux18!ufv_2 = rs_aux20!cambio_ufv
rs_aux18!fecha_ufv_1 = DTC_ufv_actual.Value
rs_aux18!fecha_ufv_2 = DTP_ufv_ant.Value

If rs_aux18!fisco_a_favor >= rs_aux18!saldo_a_favor_depend Then
rs_aux18!saldo_util = rs_aux18!mes_anterior
Else
rs_aux18!saldo_util = rs_aux18!fisco_a_favor
End If
'impuesto_a_pagar
If rs_aux18!fisco_a_favor > 0 Then
rs_aux18!impuesto_a_pagar = Round(rs_aux18!fisco_a_favor - rs_aux18!saldo_util, 0)
rs_aux18!rciva = rs_aux18!impuesto_a_pagar
Else
rs_aux18!impuesto_a_pagar = 0
rs_aux18!rciva = rs_aux18!impuesto_a_pagar
End If
If rs_aux18!saldo_a_favor_depend > 0 Then
rs_aux18!saldo_para_mes_sig = Round(rs_aux18!saldo_a_favor_depend - rs_aux18!saldo_util, 0)
Else
rs_aux18!saldo_para_mes_sig = 0
End If
'para que se descuente

'If rs_aux18!beneficiario_codigo = "159256" Then
'sino = ""
'End If
 rs_aux18!total_dsctos = rs_datos2!anticipo_sueldo + rs_datos2!anticipo_refrigerio + rs_datos2!prestamo + rs_datos2!afp1 + rs_datos2!afp2 + rs_datos2!otros_dsctos + rs_aux18!impuesto_a_pagar
                        
 rs_aux18!liquido_pagable_bs = rs_datos2!total_ganado - rs_aux18!total_dsctos
 rs_aux18!liquido_pagable_us = rs_datos2!liquido_pagable_bs / GlTipoCambioOficial

'Else
'rs_aux18!fisco_a_favor = 0
'rs_aux18!dependiente_a_favor = 0
'rs_aux18!mes_anterior = 0
'rs_aux18!actualizacion = 0
'rs_aux18!total = 0
'rs_aux18!saldo_a_favor_depend = 0
'rs_aux18!saldo_util = 0
'rs_aux18!impuesto_a_pagar = 0
'rs_aux18!saldo_para_mes_sig = 0
'
'End If ' minimo para entrar al rc iva
''Else ' si no es mayor a 2 salarios minimos

End If

rs_aux18.MoveNext
Wend
ProgressBar1.Visible = False
 Call ABRIR_TABLA_DET(2)
End Sub

Private Sub generar_personas()
NUM_PERS = 0
nuevos = ""
expirados = ""

Dim rs_aux16 As New ADODB.Recordset

 If rs_aux16.State = 1 Then rs_aux16.Close
      'SE ABRE LAS SUB PLANILLAS
      rs_aux16.Open "select * from ro_pagos_cronograma where ges_gestion = '" & rs_datos!ges_gestion & "' AND planilla_codigo = '" & rs_datos!planilla_codigo & "' AND mes_grupo = " & rs_datos!mes_grupo & " AND numero_pago = 1 ", db, adOpenKeyset, adLockReadOnly  ', adOpenKeyset, adOpenStatic, adCmdText
      rs_aux16.MoveFirst
While Not rs_aux16.EOF
ProgressBar1.Visible = True

If rs_aux6.State = 1 Then rs_aux6.Close
        'ABRE LAS PERSONAL CONTRATADAS VIGENTES
       rs_aux6.Open "SELECT * FROM ro_personal_contratado WHERE unidad_codigo_pla = '" & rs_aux16!unidad_codigo_pla & "' and estado_codigo <> 'ANL' AND estado_jubilado = 'REG'", db, adOpenKeyset, adLockOptimistic 'adOpenStatic 'order by beneficiario_denominacion
      'rs_aux6.Open "SELECT * FROM av_ro_peronal_vs_gc_beneficiario  WHERE unidad_codigo = '" & rs_datos1!unidad_codigo_pla & "' AND estado_codigo = 'APR' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic
   If rs_aux6.RecordCount > 0 Then 'verifica si existe personal en esa sub_planilla
       rs_aux6.MoveFirst
       
       
       
       With ProgressBar1
        .Max = rs_aux6.RecordCount
        .Min = 0
        .Value = 0
       End With
      'ProgressBar1.Max =
   
       While Not rs_aux6.EOF
       
        ProgressBar1.Value = ProgressBar1.Value + 1
            
            'VARIABLES PARA CALCULAR SUELDO SI SU CONTRATO TERMINA EN ESTE MES
            DIA_FN = Day(rs_aux6!fecha_expiracion) 'FECHA FIN
            MES_FN = Month(rs_aux6!fecha_expiracion)
            ANO_FN = Year(rs_aux6!fecha_expiracion)
'            If rs_aux6!beneficiario_codigo = "4773922" Then
'            sino = ""
'            End If
'
           
     expira = Day(DateSerial(rs_datos!ges_gestion, rs_datos!mes_grupo + 1, 0))
     'SE CONMPONE LA FECHA PARA COMPARAR CON EL FIN DEL CONTRATO
     FIN = "1" & "/" & rs_datos!mes_grupo & "/" & rs_datos!ges_gestion
     inicio = DateSerial(Year(FIN), Month(FIN) + 1, 0)
    'COMPARACION DE FECHA DE FIN DE CONTRATO
    If rs_aux6!fecha_expiracion >= FIN Then 'PARA LAS BAJAS
      If rs_aux6!fecha_ingreso <= inicio Then 'INICIO CONTRATO
        If rs_aux5.State = 1 Then rs_aux5.Close
'        rs_aux5.Open "select * from ro_pagos_cronograma_detalle where ges_gestion = '" & Ado_datos1.Recordset!ges_gestion & "' AND planilla_codigo = '" & Ado_datos1.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos1.Recordset!mes_grupo & " AND numero_pago = " & Ado_datos1.Recordset!NUMERO_PAGO & " AND beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "' AND unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo_pla & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        'VERIFICA SI YA EXISTE EN LA PLANILLA LA PERSONA
        rs_aux5.Open "select * from ro_pagos_cronograma_detalle where ges_gestion = '" & Ado_datos1.Recordset!ges_gestion & "' AND unidad_codigo = '" & rs_aux16!unidad_codigo_pla & "' AND planilla_codigo = '" & Ado_datos1.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos1.Recordset!mes_grupo & " AND numero_pago = " & Ado_datos1.Recordset!NUMERO_PAGO & " AND beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        If rs_aux5.RecordCount = 0 Then 'LA PERSONA NO ESTA GENERADA
            
            'GUARDADO DE DATOS
        NUM_PERS = NUM_PERS + 1
            
'        If rs_aux6!beneficiario_codigo = "4889246" Then 'para pruebas
'        sino = ""
'        End If
        
        Dim persona As New ADODB.Recordset
        If persona.State = 1 Then persona.Close
        persona.Open "select * from gc_beneficiario where beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
            
            
           
            
            rs_datos2.AddNew
            
            
            rs_datos2!fecha_ingreso = rs_aux6!fecha_ingreso
            rs_datos2!fecha_expiracion = rs_aux6!fecha_expiracion
            rs_datos2!cargo_codigo = rs_aux6!cargo_codigo
            rs_datos2!puesto_codigo = rs_aux6!puesto_codigo
            rs_datos2!beneficiario_haber_mensual = rs_aux6!beneficiario_haber_mensual
            rs_datos2!Unidad = rs_aux6!unidad_codigo
            
            rs_datos2!ges_gestion = rs_datos!ges_gestion
            rs_datos2!planilla_codigo = rs_datos!planilla_codigo
            rs_datos2!mes_grupo = rs_datos1!mes_grupo
            rs_datos2!NUMERO_PAGO = rs_datos1!NUMERO_PAGO
            
            rs_datos2!beneficiario_codigo = rs_aux6!beneficiario_codigo
            VAR_BENEF = LTrim(RTrim(rs_aux6!beneficiario_codigo))
            rs_datos2!unidad_codigo = rs_aux16!unidad_codigo_pla
            rs_datos2!tipo_moneda = "BOB"
            rs_datos2!tipo_cambio = GlTipoCambioOficial
            'Adicionar  beneficiario_haber_mensual_ant WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'
            DIA_IN = Day(rs_aux6!fecha_ingreso)
            MES_IN = Month(rs_aux6!fecha_ingreso)
            ANO_IN = Year(rs_aux6!fecha_ingreso)
            
            DIA_HOY = Day(Now)
              
            DIA_FN = Day(rs_aux6!fecha_expiracion) 'FECHA FIN
            MES_FN = Month(rs_aux6!fecha_expiracion)
            ANO_FN = Year(rs_aux6!fecha_expiracion)
             If DIA_FN > 30 Then
             DIA_FN = 30
             End If
          If rs_aux6!beneficiario_codigo = "2689587" Then
            sino = ""
            End If
          
            'VERIFICA SI ENTRO EN EL MES Y EL AÑO AL QUE CORRESPONDE LA PLANILLA
            If MES_IN = rs_datos2!mes_grupo And ANO_IN = rs_datos2!ges_gestion Then
             'CALCULO PARA PAGAR DE DIAS TRABAJADOS EN EL MES EN CASO DE QUE NO ENTRO EN EL DIA 1 DEL MES
             rs_datos2!sueldo_basico = (rs_aux6!beneficiario_haber_mensual / 30) * (30 - (DIA_IN - 1))
             rs_datos2!dias_trabajados = (30 - (DIA_IN - 1))
             nuevos = nuevos & "    " & persona!beneficiario_codigo & " " & persona!beneficiario_denominacion & vbCrLf & "    Fecha Ingreso: " & rs_aux6!fecha_ingreso & vbCrLf & "    Haber Basico: " & rs_datos2!beneficiario_haber_mensual & vbCrLf & "    Dias Trabajados: " & rs_datos2!dias_trabajados & vbCrLf & "    Haber Basico del Mes: " & rs_datos2!sueldo_basico & vbCrLf & vbCrLf
             
            Else
              'SI ESTRO EN ANTES DE EL MES DE LA PLANILLA
              rs_datos2!sueldo_basico = rs_aux6!beneficiario_haber_mensual
              rs_datos2!dias_trabajados = "30"
            End If
             'VERIFICA SI SU CONTRATO EXPIRA EN EL MES DE LA PLANILLA
              If MES_FN = rs_datos2!mes_grupo And ANO_FN = rs_datos2!ges_gestion Then 'FECHA FIN
              
             'CALCULO DE PAGO POR DIAS TRABAJADOS
              rs_datos2!sueldo_basico = (rs_aux6!beneficiario_haber_mensual / 30) * (DIA_FN)
              rs_datos2!dias_trabajados = DIA_FN
              expirados = expirados & "    " & persona!beneficiario_codigo & " " & persona!beneficiario_denominacion & vbCrLf & "    Fecha Fin: " & rs_aux6!fecha_expiracion & vbCrLf & "    Haber Basico: " & rs_datos2!beneficiario_haber_mensual & vbCrLf & "    Dias Trabajados: " & rs_datos2!dias_trabajados & vbCrLf & "    Haber Basico del Mes: " & rs_datos2!sueldo_basico & vbCrLf & vbCrLf
           
             End If
            
           sino = persona.RecordCount
            
            
            
            rs_datos2!monto_refrigerio = IIf(IsNull(rs_aux6!beneficiario_otro_mensual), "0", rs_aux6!beneficiario_otro_mensual)

             'PONE EN ULTIMO DIA DEL MES PARA COMPARAR ANTIGUEDAD
             fecha_pla = DateSerial(rs_datos!ges_gestion, rs_datos!mes_grupo + 1, 1 - 1)
             'fecha_pla = "29/02/2016" SOLO PARA PRUEBAS
             'ABRE TABLA DONDE ESTAN LOS PARAMETROS DE ANTIGUEDAD
             If rs_aux8.State = 1 Then rs_aux8.Close
             rs_aux8.Open "select * from rc_antiguedad", db, adOpenKeyset, adLockOptimistic, adCmdText
             rs_aux8.MoveFirst
             While Not rs_aux8.EOF
             'GUARDA LAS FECHA MINIMA Y LA MAXIMA EN LA QUE DEBERIA ENTRAR LA PERSONA PARA RECIBIR EL BONO ANTIGUEDAD
             f1 = DateAdd("yyyy", -rs_aux8!parametro_inicial, CDate(fecha_pla))
             f2 = DateAdd("yyyy", -rs_aux8!parametro_final, CDate(fecha_pla))
             'COMPARA LA FECHA INGRESO CON EL PARAMETRO
             If rs_aux6!fecha_ingreso <= CDate(f1) And rs_aux6!fecha_ingreso > CDate(f2) Then
             'GUARDA EL MONTO QUE CORRESPONDE
             rs_datos2!bono_antiguedad = rs_aux8!antig_valor
             rs_aux8.MoveLast
             End If
             rs_aux8.MoveNext
             Wend
             
 rs_datos2!bono_transporte = 0
    
   ' End If
            'rs_datos2!horas_extras = dtc_refrigerio.Text
            'rs_datos2!bono_transporte = dtc_refrigerio.Text
             'rs_datos2!total_ganado = rs_datos2!sueldo_basico + rs_datos2!monto_refrigerio + rs_datos2!bono_antiguedad
             
            rs_datos2!total_ganado = rs_datos2!sueldo_basico + rs_datos2!bono_antiguedad + rs_datos2!bono_transporte
            rs_datos2!provision_aguinaldo = rs_datos2!total_ganado * 0.0833
            rs_datos2!prevision_indemnizacion = rs_datos2!total_ganado * 0.0833
            rs_datos2!anticipo_sueldo = "0"
            rs_datos2!anticipo_refrigerio = "0"
            
            If VAR_BENEF = "3395947" Then 'SOLO PARA PRUEBAS
            sino = ""
            End If
            
            'VARIABLE PARA SUMA DE PAGOS DE PRESTAMO
            PRESTAMO_TOTAL = 0
            Set rs_aux24 = New Recordset
            If rs_aux24.State = 1 Then rs_aux24.Close
            'VERIFICA SI TIENE ALGUN PRESTAMO LA PERSONA
            rs_aux24.Open "select * from ro_prestamos where beneficiario_codigo = '" & VAR_BENEF & "'  and estado_codigo = 'APR' AND codigo_empresa = " & rs_aux6!codigo_empresa & " ", db, adOpenKeyset, adLockOptimistic, adCmdText
            If rs_aux24.RecordCount > 0 Then
            rs_aux24.MoveFirst
            While Not rs_aux24.EOF
                If rs_aux24!estado_codigo = "APR" Then
                    Set rs_aux25 = New Recordset
                    If rs_aux25.State = 1 Then rs_aux25.Close
                    'VERIFICA SI EXISTE PAGO PARA ESTE MES SEGUN EL CRONOGRANA GENERADO
                    rs_aux25.Open "select * from ro_prestamo_prog where beneficiario_codigo = '" & VAR_BENEF & "' and prestamo_codigo = " & rs_aux24!prestamo_codigo & " AND mes_planilla = " & rs_datos!mes_grupo & " and year(cobranza_fecha_prog) = '" & rs_datos!ges_gestion & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
                    If rs_aux25.RecordCount > 0 Then
                    'SUMA LOS PAGOS
                    PRESTAMO_TOTAL = PRESTAMO_TOTAL + rs_aux25!cobranza_programada_bs
                    'APRUEBA EL PAGO
                    rs_aux25!estado_codigo = "APR"
                    rs_aux25!cobranza_fecha_cobro = Date
                    rs_aux25.Update
                    
                    rs_aux24!correl_prog = rs_aux25!prestamo_prog_codigo
                    Set rs_aux26 = New Recordset
                    If rs_aux26.State = 1 Then rs_aux26.Close
                    'SUMA TODOS LOS PAGOS REALIZADOS PARA GUARDAR EN LA CABECERA
                    rs_aux26.Open "select SUM(cobranza_programada_bs)AS TOTAL_COB from ro_prestamo_prog where beneficiario_codigo = '" & VAR_BENEF & "' and estado_codigo = 'APR' AND prestamo_codigo = " & rs_aux24!prestamo_codigo, db, adOpenKeyset, adLockOptimistic, adCmdText
                    rs_aux24!total_cobrado = rs_aux26!TOTAL_COB
                     
                    End If
               End If
                rs_aux24.MoveNext
            Wend
            End If
            rs_datos2!prestamo = PRESTAMO_TOTAL
            'CALCULO DE AFP
            Select Case rs_aux6!beneficiario_codigo_afp
                Case "1006803"      'AFP1
                    rs_datos2!afp1 = rs_datos2!total_ganado * 0.1271
                    rs_datos2!afp2 = "0"       'falta 987654
                    VAR_NETO = rs_datos2!total_ganado - rs_datos2!afp1
                Case "987654"       'AFP2
                    rs_datos2!afp1 = "0"       'falta 1006803
                    rs_datos2!afp2 = rs_datos2!total_ganado * 0.1271
                    VAR_NETO = rs_datos2!total_ganado - rs_datos2!afp2
                Case Else
                    rs_datos2!afp1 = "0"
                    rs_datos2!afp2 = "0"
                    VAR_NETO = rs_datos2!total_ganado
            End Select
             '
'            VAR_IVA = 1805 * 4
'            If VAR_NETO > VAR_IVA Then
'                rs_datos2!rciva = rs_datos2!total_ganado * 0.13
'            Else
'                rs_datos2!rciva = "0"        'mayor a 4 SUELDOS BASICOA
'            End If
            '
            db.Execute "UPDATE ro_controlasistencia SET ges_gestion = year(Fecha_control), Mes_control = month(Fecha_control), Dia_control= day(Fecha_control)"
            'sqlAux = "SELECT '     TOTAL MINUTOS DE RETRASO: ' + CONVERT(VARCHAR, ISNULL(SUM(DATEDIFF(MINUTE, '0:00:00', Tardanza)),0)) AS totHrs FROM ro_controlasistencia WHERE beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' "
            'rs_AsisTT.Open sqlAux, db, adOpenKeyset, adLockOptimistic, adCmdText
            'rs_AsisTT.MoveFirst
            'AdoAsistencia.Caption = CStr(rs_AsisTT!totHrs)
            '

            'db.Execute "UPDATE ro_controlasistencia SET TotalMin1 = convert(int,TardanzaCadena) "
            'rs_aux9.Open "select sum(convert(int,TardanzaCadena)) as TardanzaMes from ro_controlasistencia where beneficiario_codigo = " & VAR_BENEF & " and Mes_control = '" & Str(Ado_datos1.Recordset!mes_grupo) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
             'Dim rs_aux9 As New ADODB.Recordset
            If rs_aux9.State = 1 Then rs_aux9.Close
            rs_aux9.Open "select sum(AtrasoMin1) as TardanzaMes from ro_controlasistencia where beneficiario_codigo = '" & RTrim(LTrim(VAR_BENEF)) & "' AND ges_gestion = '" & RTrim(LTrim(Ado_datos1.Recordset!ges_gestion)) & "' and Mes_control = '" & RTrim(LTrim(Str(Ado_datos1.Recordset!mes_grupo))) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
             'select sum(convert(int,TardanzaCadena)) as TardanzaMes from ro_controlasistencia where beneficiario_codigo = '6960987' and Mes_control = 7
            If rs_aux14.State = 1 Then rs_aux14.Close
            mesnom = UCase(MonthName(Ado_datos1.Recordset!mes_grupo))
            rs_aux14.Open "select sum(total_minuto) as PermisoMes from ro_permisos where beneficiario_codigo = '" & RTrim(LTrim(VAR_BENEF)) & "' AND ges_gestion = '" & RTrim(LTrim(Ado_datos1.Recordset!ges_gestion)) & "' AND Mes_control = '" & mesnom & "' AND estado_codigo = 'APR' and dias_permiso = 0 AND codigo_empresa = " & rs_aux6!codigo_empresa, db, adOpenKeyset, adLockOptimistic, adCmdText
            If rs_aux14!PermisoMes <> "NULL" Then
                permisos = rs_aux14!PermisoMes
            Else
                permisos = "0"
            End If
            If rs_aux9!TardanzaMes <> "NULL" Then
                totalminutos = rs_aux9!TardanzaMes - permisos
                If totalminutos >= 45 And totalminutos <= 60 Then
                    rs_datos2!otros_dsctos = ((rs_datos2!sueldo_basico / 30) / 2)
                Else
                    If totalminutos >= 61 And totalminutos <= 75 Then
                        rs_datos2!otros_dsctos = (rs_datos2!sueldo_basico / 30)
                    Else
                        If totalminutos >= 76 And totalminutos <= 105 Then
                            rs_datos2!otros_dsctos = ((rs_datos2!sueldo_basico / 30) * 2)
                        Else
                            If totalminutos >= 106 Then
                                rs_datos2!otros_dsctos = ((rs_datos2!sueldo_basico / 30) * 3)
                            Else
                                rs_datos2!otros_dsctos = 0
                            End If
                        End If
                    End If
                End If
            Else
              If continuar = "SI" Then
                sino = MsgBox("No se Cargo la asistencia del mes de " & UCase(MonthName(rs_datos1!mes_grupo)) & " de algunas personas " & vbCrLf & "¿Desea generar de todas maneras?" & vbCrLf & "NOTA: En el campo de OTROS DESCUENTOS se asignará cero (0) por defecto", vbCritical + vbYesNo, "Atención")
                If sino = vbYes Then
                    rs_datos2!otros_dsctos = 0
                    continuar = "NO"
                    Numero = Numero + 1
                Else
                    ProgressBar1.Visible = False
                    Fra_personal_Ppla.Visible = False
                    FraNavega.Enabled = True
                    fraOpciones.Enabled = True
                    ' FraGrabarCancelar.Visible = True
                    dg_datos.Enabled = True
                    dg_det1.Enabled = True
                    fra_opciones_det_1.Enabled = True
                    fra_opciones_det_2.Enabled = True
        
                    dg_det2.Enabled = True
                    Call ABRIR_TABLA_DET(2)
                    Exit Sub
                End If
              Else
                rs_datos2!otros_dsctos = 0
                Numero = Numero + 1
              End If
            End If
            'rs_datos2!otros_dsctos = "0"   'FIN Atrasos y Faltas
            rs_datos2!r_provision_aguinaldo = "0"
            rs_datos2!r_prevision_indemnizacion = "0"
            
            If rs_aux15.State = 1 Then rs_aux15.Close
                rs_aux15.Open "select SUM(monto) AS totalmonto, SUM(dias) AS Totaldias from ro_memorandas where ges_gestion = '" & Ado_datos1.Recordset!ges_gestion & "' AND mes_descuento = " & Ado_datos1.Recordset!mes_grupo & " AND beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "' AND descuento_pla = 'SI' AND estado_codigo = 'APR' AND tipo_memo <> 'ANT' AND codigo_empresa = " & rs_aux6!codigo_empresa & "", db, adOpenKeyset, adLockOptimistic, adCmdText
             
         If rs_aux15.RecordCount <> 0 Then
            If rs_aux15!totalmonto > 0 Then
                total = rs_datos2!otros_dsctos + IIf(IsNull(rs_aux15!totalmonto), 0, rs_aux15!totalmonto)
               rs_datos2!otros_dsctos = total
              End If
              
              If rs_aux15!Totaldias > 0 Then
                total = rs_datos2!otros_dsctos + ((rs_aux6!beneficiario_haber_mensual / 30) * rs_aux15!Totaldias)
                'total = total + rs_datos2!otros_dsctos
                rs_datos2!otros_dsctos = total
            End If
              
         End If
         If rs_aux30.State = 1 Then rs_aux30.Close
              rs_aux30.Open "select SUM(monto) AS totalmonto, SUM(dias) AS Totaldias from ro_memorandas where ges_gestion = '" & Ado_datos1.Recordset!ges_gestion & "' AND mes_descuento = " & Ado_datos1.Recordset!mes_grupo & " AND beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "' AND descuento_pla = 'SI' AND estado_codigo = 'APR' AND tipo_memo = 'ANT' AND codigo_empresa = " & rs_aux6!codigo_empresa & "", db, adOpenKeyset, adLockOptimistic, adCmdText
             
         If rs_aux30.RecordCount <> 0 Then
         
              
              If rs_aux30!totalmonto > 0 Then
                total = rs_datos2!anticipo_sueldo + IIf(IsNull(rs_aux30!totalmonto), 0, rs_aux30!totalmonto)
               rs_datos2!anticipo_sueldo = total
              End If
              
'              If rs_aux30!Totaldias > 0 Then
'                total = rs_aux30!anticipo_sueldo + ((rs_aux6!beneficiario_haber_mensual / 30) * rs_aux30!Totaldias)
'                'total = total + rs_datos2!otros_dsctos
'                rs_datos2!anticipo_sueldo = total
'              End If
              
          End If
            'rs_datos2.Update
            'rs_datos2!total_dsctos = "0"
            rs_datos2!total_dsctos = rs_datos2!anticipo_sueldo + rs_datos2!anticipo_refrigerio + rs_datos2!prestamo + rs_datos2!afp1 + rs_datos2!afp2 + rs_datos2!otros_dsctos
                        
            rs_datos2!liquido_pagable_bs = rs_datos2!total_ganado - rs_datos2!total_dsctos
            rs_datos2!liquido_pagable_us = rs_datos2!liquido_pagable_bs / GlTipoCambioOficial
             'rs_datos2!total_dsctos = "0"
            rs_datos2!emite_factura = "N"
             
            rs_datos2!cite_conformidad = "-"
            'rs_datos2!Numero_consultoriaHist = " "
            rs_datos2!fte_financiamientoHist = "-"
            rs_datos2!estado_devengado = "REG"
             '70522788
            rs_datos2!estado_codigo = "REG"
            rs_datos2!fecha_registro = Date
            rs_datos2!usr_codigo = glusuario
            
            rs_datos2!iva_110 = "0"
            rs_datos2!fisco_a_favor = "0"
            rs_datos2!dependiente_a_favor = "0"
            rs_datos2!mes_anterior = "0"
            rs_datos2!mes_anterior_mant = "0"
            rs_datos2!saldo_util = "0"
            rs_datos2!saldo_a_favor_depend = "0"
            rs_datos2!rciva = "0"
            rs_datos2!codigo_empresa = rs_aux6!codigo_empresa
            'ABRIR_TABLA_DET (2)
            rs_datos2!hora_registro = Format(Time, "HH:mm:ss")
            
            rs_datos2.Update
            'Call OptFilGral1_Click
            'rs_datos.MoveLast
            mbDataChanged = False
    '
        End If
       End If 'INICIO CONTRATO
    Else 'PARA LAS BAJAS
    rs_aux6!estado_codigo = "ANL"
    End If 'PARA LAS BAJAS
        rs_aux6.MoveNext
       Wend
  End If 'verifica si existe personal en esa sub_planilla
       rs_aux16.MoveNext
     Wend
       
       Call ABRIR_TABLA_DET(2)
       Call ABRIR_TABLAS_AUX(5)
       Call numeracion_planilla
       'rs_datos2.RecordCount
       
    If nuevos = "" Then
    nuevos = "    Ninguna persona tiene fecha de INICIO en este mes." & vbCrLf & vbCrLf
    End If
    If expirados = "" Then
    expirados = "    Ninguna persona tiene fecha de FIN en este mes." & vbCrLf & vbCrLf
    End If
    If NUM_PERS > 0 Then
    sino = MsgBox("Se agregó " & NUM_PERS & " persona(s) a " & rs_datos!descripcion_grupo & " " & rs_datos!ges_gestion & vbCrLf & vbCrLf & "• Personas que tienen fecha de INICIO en " & UCase(MonthName(rs_datos!mes_grupo)) & vbCrLf & vbCrLf & nuevos & "---------------------------------------------------------" & vbCrLf & "• Personas que tienen fecha de FIN en " & UCase(MonthName(rs_datos!mes_grupo)) & vbCrLf & vbCrLf & expirados & "---------------------------------------------------------", vbInformation, "Atención")
    End If
    continuar = "SI"
    ProgressBar1.Visible = False
    dtc_buscar_desc.Visible = True
    Label52.Visible = True
    
End Sub
  
  
  
Public Function Dias_Del_Mes(Optional ByVal Fecha As Variant) As Integer
  
    Dim NUMDIA As Integer
    
    Dim mes As Integer, Y  As Integer
  
    If IsMissing(Fecha) Then Fecha = Date
  
    If IsDate(Fecha) Then
        Y = Year(Fecha)
        mes = Month(Fecha)
    ElseIf IsNumeric(Fecha) Then
        Y = Year(Date)
        mes = IIf(Fecha > 0 And Fecha < 13, CInt(Fecha), 0)
    ElseIf VarType(Fecha) = vbString Then
        Y = Year(Date)
        Select Case UCase(Left$(Fecha, 3))
            Case "FEB":                                             mes = 2
            Case "JAN", "MAR", "MAY", "JUL", "AUG", "OCT", "DEC":   mes = 1
            Case "APR", "JUN", "SEP", "NOV":                        mes = 4
        End Select
    End If
  
    Select Case mes
        Case 2:                     NUMDIA = IIf(saltarYear(Fecha), 29, 28)
        Case 1, 3, 5, 7, 8, 10, 12: NUMDIA = 31
        Case 4, 6, 9, 11:           NUMDIA = 30
    End Select
  
End Function

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
     '<-- Inicio                Identificación del Cliente                Fin -->
   If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
     'If VAR_SW = "NO" Then
'
        If Ado_datos.Recordset.RecordCount > 0 Then
            Call ABRIR_TABLA_DET(1)
            Call ABRIR_TABLAS_AUX(5)
            busq = 0
           
        End If
        VAR_SW = ""
    Else
        VAR_SW = ""
        'Set rs_det1 = New ADODB.Recordset
        'Set dg_det1.DataSource = rsNada
       ' Set dg_det2.DataSource = rsNada
    'End If
  End If
End Sub


Private Sub Ado_datos1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If mover = 1 Then
    mover = 0
    Exit Sub
End If
    dtc_buscar_desc.Text = ""
     '<-- Inicio                Identificación del Cliente                Fin -->
   If (Not Ado_datos1.Recordset.BOF) And (Not Ado_datos1.Recordset.EOF) Then
     'If VAR_SW = "NO" Then
'
        If Ado_datos.Recordset.RecordCount > 0 Then
            Call ABRIR_TABLA_DET(2)
           'Call ABRIR_TABLAS_AUX (0)
        End If
        VAR_SW = ""
    Else
        VAR_SW = ""
        'Set rs_det1 = New ADODB.Recordset
        'Set dg_det1.DataSource = rsNada
        'Set dg_det2.DataSource = rsNada
    'End If
  End If
End Sub

Private Sub BtnAnlDetalle_Click()
   If Ado_detalle1.Recordset("estado_activo") = "REG" Then
      sino = MsgBox("Está Seguro de cambiar a HORARIO NO LABORABLE ? (Este ya no será considerado en el Cronograma) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        Ado_detalle1.Recordset!estado_activo = "ANL"
        Ado_detalle1.Recordset!observaciones = "HORARIO NO LABORABLE"
        Ado_detalle1.Recordset.Update
        'Call ABRIR_TABLA_DET
      End If
   Else
        MsgBox "No se puede ANULAR, el registro ya fue Aprobado (Estado=APR) o ya fue Anulado anteriormente (Estado=ANL)...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAñadir_Click()
On Error GoTo UpdateErr
'If rs_aux12.State = 1 Then rs_aux12.Close
'      rs_aux12.Open "select * from ro_pagos_grupos where ges_gestion = '" & Year(Date) & "'", db, adOpenKeyset, adLockOptimistic
'     If rs_aux12.RecordCount = 0 Then
  sino = MsgBox("¿Desea que el sistema genere automáticamente la(s) planilla(s)?", vbYesNo + vbQuestion, "Atención")
  If sino = vbYes Then
    cmb_gestion.Text = Year(Date)
    fra_generar.Visible = True
    FraNavega.Enabled = False
    fraOpciones.Enabled = False
    ' FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    dg_det1.Enabled = False
    fra_opciones_det_1.Enabled = False
    fra_opciones_det_2.Enabled = False
    dg_det2.Enabled = False
    swnuevo = 1
  Else
    Call ABRIR_TABLAS_AUX(1)
    cbo_gestion_pla.Text = Year(Date)
    fra_nueva_pla.Visible = True
    FraNavega.Enabled = False
    fraOpciones.Enabled = False
    ' FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    dg_det1.Enabled = False
    fra_opciones_det_1.Enabled = False
    fra_opciones_det_2.Enabled = False
    dg_det2.Enabled = False
  End If
 Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAprobar_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
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
       MsgBox "No se puede APROBAR un registro Anulado (ERR) o Aprobado (APR) anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
Else
      MsgBox "No existen registros", vbExclamation, "Error"
End If
End Sub

Private Sub BtnBuscar_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        OptFilGral1.Visible = True
        OptFilGral2.Visible = True
'        If Ado_datos.Recordset!estado_codigo = "REG" Then
'            Call OptFilGral1_Click
'        Else
'            Call OptFilGral2_Click
'        End If
        Set ClBuscaGrid = New ClBuscaEnGridExterno
        Set ClBuscaGrid.Conexión = db
        ClBuscaGrid.EsTdbGrid = False
        Set ClBuscaGrid.GridTrabajo = dg_datos
        ClBuscaGrid.QueryUtilizado = queryinicial
        Set ClBuscaGrid.RecordsetTrabajo = rs_datos
        'ClBuscaGrid.CamposVisibles = "11010011"
        ClBuscaGrid.Ejecutar
    Else
      MsgBox "NO se puede Procesar !!. Verifique si existen registros. ", vbExclamation, "Atención!"
      OptFilGral1.Visible = True
      OptFilGral2.Visible = True
    End If
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelUpdate
'        Call ABRIR_TABLA
        rs_datos.MoveFirst
        'mbDataChanged = False
        Fra_datos.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
'        FrmABMDet.Visible = True
        dg_det1.Enabled = True
        swnuevo = 0
    End If

End Sub

Private Function ExisteReg(where As String, tabla As String) As Boolean
        Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM " & tabla & " WHERE " & where & ""
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
    
End Function


Private Sub BtnEliminar_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
On Error GoTo UpdateErr
   If ExisteReg(" ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' AND planilla_codigo = '" & Ado_datos.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos.Recordset!mes_grupo, "ro_pagos_cronograma") Then
      sino = MsgBox("No se puede ELIMINAR porque el Registro ya fue utilizado. Desea marcar como ERRADO ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "ERR"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
      sino = MsgBox("Está Seguro de ELIMINAR fisicamente el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         db.Execute "DELETE ro_pagos_grupos where ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' AND planilla_codigo = '" & Ado_datos.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos.Recordset!mes_grupo
      End If
   End If
   Call OptFilGral1_Click
   Exit Sub

UpdateErr:
  MsgBox Err.Description
  Else
      MsgBox "No existen registros", vbExclamation, "Error"
End If
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If swnuevo = 1 Then
        db.Execute "Update to_cronograma_mensual Set estado_codigo = 'REG' Where ges_gestion = '" & cmb_gestion & "' AND fmes_correl = " & mes2 & " AND zpiloto_codigo =" & dtc_codigo3.Text & "  "
'    Else
'     rs_datos!fmes_fecha_registro = DTPfecha1.Value
'     rs_datos!beneficiario_codigo_resp = dtc_codigo4.Text
'     rs_datos!observaciones = Txt_campo2.Text
'
'     rs_datos!fecha_registro = Date     'no cambia
'     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
'     rs_datos.Update    'Batch 'adAffectAll
    End If
     db.Execute "Update to_cronograma_diario Set beneficiario_codigo_resp = " & dtc_codigo4.Text & " Where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'   "
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        Call OptFilGral1_Click
    Else
        Call OptFilGral2_Click
    End If
     'rs_datos.MoveFirst
'     mbDataChanged = False
     Fra_datos.Enabled = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
'     FrmABMDet.Visible = True
     dg_det1.Enabled = True
     'dtc_desc1.BackColor = &HFFFFC0
     swnuevo = 0
  End If
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
'  If (dtc_codigo4 = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (Txt_campo2.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  
End Sub


Private Sub BtnImprimir_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim iResult As Integer
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
      CR01.Formulas(2) = "periodo = '" & Cmb_Mes & "' "
      'CR01.Formulas(2) = "periodo = '" & Cmb_Mes & "' "
      
'    cr01.StoredProcParam(0) = "2015"    'Me.Ado_datos.Recordset!ges_gestion
'    cr01.StoredProcParam(1) = "DNMAN"   'Me.Ado_datos.Recordset!unidad_codigo_tec
'    cr01.StoredProcParam(2) = 0     'Me.Ado_datos.Recordset!zpiloto_codigo
'    cr01.StoredProcParam(3) = 1     'Me.Ado_datos.Recordset!fmes_correl
    
    CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo_tec
    CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!zpiloto_codigo
    CR01.StoredProcParam(3) = Me.Ado_datos.Recordset!fmes_correl
    
    iResult = CR01.PrintReport
    If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnImprimir2_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim iResult As Integer
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
    iResult = CR01.PrintReport
    If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnModDetalle_Click()
    If Ado_detalle1.Recordset("estado_activo") = "ANL" Then             'And Ado_detalle1.Recordset("estado_codigo") = "REG"
      sino = MsgBox("Está Seguro de cambiar a HORARIO LABORABLE ? (Este volverá a ser considerado en el Cronograma) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        Ado_detalle1.Recordset!estado_activo = "REG"
        Ado_detalle1.Recordset!observaciones = "HORARIO LABORABLE"
        Ado_detalle1.Recordset.Update
        'Call ABRIR_TABLA_DET
      End If
   Else
        MsgBox "No se puede Habilitar, el registro ya fue Procesado (Estado=APR) o ya está Habilitado (Estado=REG) ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnModificar_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
  On Error GoTo EditErr
  
 'lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
    Fra_modificar.Visible = True
          FraNavega.Enabled = False
       fraOpciones.Enabled = False
       'FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
         dg_det1.Enabled = False
         fra_opciones_det_1.Enabled = False
          fra_opciones_det_2.Enabled = False
        dg_det2.Enabled = False
        swnuevo = 2
    '    BtnVer.Visible = True
    Else
      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
    End If
  Exit Sub

EditErr:
  MsgBox Err.Description
  Else
      MsgBox "No existen registros", vbExclamation, "Error"
End If
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub BtnVer_Click()
    'ARREGLO 1
    
End Sub

Private Sub Cmb_Mes_Change()
    Select Case Cmb_Mes
        Case "ENERO"
            mes2 = "1"
        Case "FEBRERO"
            mes2 = "2"
        Case "MARZO"
            mes2 = "3"
        Case "ABRIL"
            mes2 = "4"
        Case "MAYO"
            mes2 = "5"
        Case "JUNIO"
            mes2 = "6"
        Case "JULIO"
            mes2 = "7"
        Case "AGOSTO"
            mes2 = "8"
        Case "SEPTIEMBRE"
            mes2 = "9"
        Case "OCTUBRE"
            mes2 = "10"
        Case "NOVIEMBRE"
            mes2 = "11"
        Case "DICIEMBRE"
            mes2 = "12"

    End Select
End Sub




Private Sub cbo_mes_pla_Click()
txt_mes_grupo.Text = cbo_mes_pla.ListIndex
txt_mes_grupo.Text = Val(txt_mes_grupo.Text) + 1
End Sub

Private Sub cbo_mes_rep_Click()
txt_mes.Text = cbo_mes_rep.ListIndex
txt_mes.Text = Val(txt_mes.Text) + 1
End Sub


Private Sub Command3_Click()
If Ado_datos2.Recordset.RecordCount > 0 Then
 On Error GoTo UpdateErr
DTP_ufv_ant.Value = Date
DTC_ufv_actual.Value = Date
fra_ufv.Visible = True
FraNavega.Enabled = False
       fraOpciones.Enabled = False
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
         dg_det1.Enabled = False
         fra_opciones_det_1.Enabled = False
          fra_opciones_det_2.Enabled = False
        dg_det2.Enabled = False
Exit Sub
UpdateErr:
  MsgBox Err.Description
  Else
  sino = MsgBox("Primero tiene que crear la planilla ", vbCritical, "Atención")
  End If
End Sub


Private Sub Command4_Click()
' SOLO PARA POBRAR FUNCIONES
Dim a, B As Date
a = InputBox("Introduzca la fecha ini", "Dato saca")
B = InputBox("Introduzca la fecha fin", "Dato saca")
Call fun_dias360(a, B)
End Sub

Private Sub Command5_Click()
Call RETROACTIVO
End Sub

Private Sub dg_datos_Click()
    VAR_SW = "NO"
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub



Private Sub dtc_codigo_1_Click(Index As Integer, Area As Integer)
 dtc_descripcion.BoundText = dtc_codigo.BoundText
End Sub




Private Sub dt_unidad_cod_Click(Area As Integer)
 dt_unidad_det.BoundText = dt_unidad_cod.BoundText
End Sub

Private Sub dt_unidad_det_Click(Area As Integer)
 dt_unidad_cod.BoundText = dt_unidad_det.BoundText
End Sub

Private Sub dtc_buscar_ci_Click(Area As Integer)
dtc_buscar_desc.BoundText = dtc_buscar_ci.BoundText
End Sub



Private Sub dtc_buscar_desc_Change()
dtc_buscar_ci.BoundText = dtc_buscar_desc.BoundText
 If dtc_buscar_ci.SelectedItem <> "" Then
 'busq = busq + 1
 'If busq = 2 Then
 Call ABRIR_TABLA_DET(3)
 'busq = 0
 'End If
 End If
End Sub

Private Sub dtc_codigo_Click(Area As Integer)
 dtc_descripcion.BoundText = dtc_codigo.BoundText
 dtc_sueldo.BoundText = dtc_codigo.BoundText
  dtc_refrigerio.BoundText = dtc_codigo.BoundText

End Sub

Private Sub dtc_descripcion_Click(Area As Integer)
    dtc_codigo.BoundText = dtc_descripcion.BoundText
    dtc_sueldo.BoundText = dtc_descripcion.BoundText
    dtc_refrigerio.BoundText = dtc_descripcion.BoundText
   
End Sub

Private Sub dtc_descripcion_LostFocus()
'txt_haber_mensual.Text = Ado_datos4.Recordset!beneficiario_haber_mensual

End Sub

Private Sub dtc_pla_cod_Click(Area As Integer)
 dtc_pla_det.BoundText = dtc_pla_cod.BoundText
End Sub

Private Sub dtc_pla_det_Click(Area As Integer)
 dtc_pla_cod.BoundText = dtc_pla_det.BoundText
End Sub

Private Sub dtc_refrigerio_Click(Area As Integer)
 dtc_descripcion.BoundText = dtc_refrigerio.BoundText
 dtc_sueldo.BoundText = dtc_refrigerio.BoundText
  dtc_refrigerio.BoundText = dtc_refrigerio.BoundText
txt_total_ganado.Text = (dtc_sueldo.Text + dtc_refrigerio.Text)
End Sub

Private Sub dtc_rep_cod_Click(Area As Integer)
 dtc_rep_det.BoundText = dtc_rep_cod.BoundText
  Option1.Value = False
End Sub

Private Sub dtc_rep_det_Click(Area As Integer)
  dtc_rep_cod.BoundText = dtc_rep_det.BoundText
  Option1.Value = False
End Sub

Private Sub dtc_sueldo_Click(Area As Integer)
 dtc_descripcion.BoundText = dtc_sueldo.BoundText
 dtc_sueldo.BoundText = dtc_sueldo.BoundText
  dtc_refrigerio.BoundText = dtc_sueldo.BoundText
txt_total_ganado.Text = (dtc_sueldo.Text + dtc_refrigerio.Text)
End Sub

Private Sub Form_Load()
 'frm_ro_pagos_grupo_principal.Visible = True
  Call ABRIR_TABLAS_AUX(1)
    swnuevo = 0
    VAR_SW = ""
    continuar = "SI"
    'Fra_Gestion.Visible = True
    'VAR_GES = Cmb_gestion.Text
    'parametro = Aux
    Call OptFilGral1_Click
  
   
    
'    Fra_datos.Enabled = False
  '  dg_datos.Enabled = True
    'lbl_aux1.Visible = False
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
   'If Not Ado_datos.Recordset.EOF Then
            'SSTab1.Tab = 0
            'SSTab1.TabEnabled(0) = True
            ''SSTab1.TabEnabled(1) = False
            'SSTab1.TabVisible(1) = False
   'End If
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX(Pos As Integer)
Select Case Pos
 Case 5
'    busqueda
    Set rs_aux17 = New ADODB.Recordset
    If rs_aux17.State = 1 Then rs_aux17.Close
    rs_aux17.Open "select * from av_gc_beneficiario_vs_ro_pagos_cronograma_detalle where ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' AND planilla_codigo = '" & Ado_datos.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos.Recordset!mes_grupo & "order by beneficiario_denominacion asc", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_datos_busq.Recordset = rs_aux17
    dtc_buscar_ci.BoundText = dtc_buscar_desc.BoundText
    If rs_aux17.RecordCount > 0 Then
    dtc_buscar_desc.Visible = True
    Label52.Visible = True
    Else
    dtc_buscar_desc.Visible = False
    Label52.Visible = False
    End If
    
Case 3
'    gc_unidad_ejecutora
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "SELECT * FROM av_ro_peronal_vs_gc_beneficiario  WHERE unidad_codigo_pla = '" & rs_datos1!unidad_codigo_pla & "' AND estado_codigo <> 'ANL' order by beneficiario_denominacion", db, adOpenStatic
    sql = "rp_agregar_nuevo_a_planilla " & rs_datos1!unidad_codigo_pla & "," & rs_datos1!mes_grupo & "," & rs_datos1!ges_gestion
    rs_datos4.Open sql, db, adOpenKeyset, adLockOptimistic, adCmdText
   Set Ado_datos4.Recordset = rs_datos4
   dtc_descripcion.BoundText = dtc_codigo.BoundText
Case 2
       ' gc_unidad_ejecutora
    Set rs_aux7 = New ADODB.Recordset
    If rs_aux7.State = 1 Then rs_aux7.Close
    rs_aux7.Open "SELECT * FROM rc_planilla_grupo", db, adOpenStatic
    Set Ado_datos_rep.Recordset = rs_aux7
  dtc_rep_det.BoundText = dtc_rep_cod.BoundText
      
Case 4
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos5.Close
    rs_datos8.Open "rc_planilla_sub_grupo where estado_codigo = 'APR' AND planilla_codigo = '" & rs_datos!planilla_codigo & " '", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos8
    dt_unidad_det.BoundText = dt_unidad_cod.BoundText
Case 1
    'Beneficiario Funcionario CGI (Vendedor, Cobrador, Adm, etc.)
    Set rs_aux11 = New ADODB.Recordset
    If rs_aux11.State = 1 Then rs_datos11.Close
    'rs_aux11.Open "rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_aux11.Open "SELECT * FROM rc_planilla_grupo", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_aux11
    dtc_pla_det.BoundText = dtc_pla_cod.BoundText
    
End Select
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
'    Call pnivel1(dtc_codigo1.BoundText)
'    dtc_desc10.Enabled = True
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub Image1_Click()

fra_reportes.Visible = False

 FraNavega.Enabled = True
       fraOpciones.Enabled = True
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
         dg_det1.Enabled = True
         fra_opciones_det_1.Enabled = True
fra_opciones_det_2.Enabled = True

        dg_det2.Enabled = True
End Sub

Private Sub Label36_Click()
fra_imprimir.Visible = True
fra_reportes.Visible = False
End Sub

Private Sub Label37_Click()
fra_reportes.Visible = False
 FraNavega.Enabled = True
       fraOpciones.Enabled = True
       ' FraGrabarCancelar.Visible = Tru
        dg_datos.Enabled = True
         dg_det1.Enabled = True
         fra_opciones_det_1.Enabled = True
fra_opciones_det_2.Enabled = True

        dg_det2.Enabled = True

End Sub


Private Sub OptFilGral1_Click()
    '===== Proceso para filtrado general de datos (registros NO aprobados)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close   '
    'queryinicial = "select * From tv_cronograma_mensual_zona WHERE estado_codigo = 'REG' "      'AND unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '2015'
    queryinicial = "select * From ro_pagos_grupos WHERE estado_codigo = 'REG'"
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
      
    
End Sub

Private Sub OptFilGral2_Click()
    '===== Proceso para filtrado general de datos (todos los registros)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "Select * from tv_cronograma_mensual_zona where  unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '" & glGestion & "' "
    queryinicial = "Select * from ro_pagos_grupos" 'WHERE estado_codigo <> 'ERR'         'where  unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '" & glGestion & "'
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
     
End Sub

Private Sub ABRIR_TABLA()
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    queryinicial = "Select * from ao_solicitud_cotiza_venta where " + parametro
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    Set dg_datos.DataSource = Ado_datos.Recordset
        
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

Private Sub ABRIR_TABLA_DET(posicion As Integer)
Select Case posicion

 Case 1
   ' Dim rs_datos1 As New ADODB.Recordset
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "select * from ro_pagos_cronograma where ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' AND planilla_codigo = '" & Ado_datos.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos.Recordset!mes_grupo & " order by unidad_codigo_pla ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_datos1.Recordset = rs_datos1
    Set dg_det1.DataSource = Ado_datos1.Recordset
     If Ado_datos1.Recordset.RecordCount > 0 Then
     
    Else
     
    Set rs_datos2 = New ADODB.Recordset
'    If rs_datos2.State = 1 Then rs_datos2.Close
'    rs_datos2.Open "select * from av_gc_beneficiario_vs_ro_pagos_cronograma_detalle where ges_gestion = '1000' ", db, adOpenKeyset, adLockOptimistic, adCmdText
       If rs_datos2.State = 1 Then rs_datos2.Close
       rs_datos2.Open "select * from av_gc_beneficiario_vs_ro_pagos_cronograma_detalle where ges_gestion = '1000' ", db, adOpenKeyset, adLockOptimistic, adCmdText
      'Set rs_datos2 = Nothing
     
     Set Ado_datos2.Recordset = rs_datos2
     
    Set dg_det2.DataSource = Ado_datos2.Recordset
End If

  
  Case 2
     If Ado_datos1.Recordset.RecordCount > 0 Then
        FraDet2.Visible = True
        'Dim rs_datos2 As New ADODB.Recordset
        Set rs_datos2 = New ADODB.Recordset
        If rs_datos2.State = 1 Then rs_datos2.Close
        rs_datos2.Open "select * from av_gc_beneficiario_vs_ro_pagos_cronograma_detalle where ges_gestion = '" & Ado_datos1.Recordset!ges_gestion & "' AND planilla_codigo = '" & Ado_datos1.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos1.Recordset!mes_grupo & " AND numero_pago = " & Ado_datos1.Recordset!NUMERO_PAGO & " AND unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo_pla & "' order by Numero_consultoriaHist asc", db, adOpenKeyset, adLockOptimistic, adCmdText
        Set Ado_datos2.Recordset = rs_datos2
        Set dg_det2.DataSource = Ado_datos2.Recordset
        
        If Ado_datos2.Recordset.RecordCount > 0 Then

       
        Else
        Set rs_datos2 = New ADODB.Recordset
        If rs_datos2.State = 1 Then rs_datos2.Close
           rs_datos2.Open "select * from av_gc_beneficiario_vs_ro_pagos_cronograma_detalle where ges_gestion = 1000", db, adOpenKeyset, adLockOptimistic, adCmdText
        Set Ado_datos2.Recordset = rs_datos2
        End If
           Call ABRIR_TABLAS_AUX(0)
     Else
        Set dg_det2.DataSource = rsNada
        FraDet2.Visible = False
    End If
  
Case 3
 Set rs_datos2 = New ADODB.Recordset
 If rs_datos2.State = 1 Then rs_datos2.Close
 rs_datos2.Open "select * from av_gc_beneficiario_vs_ro_pagos_cronograma_detalle where ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' AND planilla_codigo = '" & Ado_datos.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos.Recordset!mes_grupo & "AND beneficiario_codigo = '" & dtc_buscar_ci.Text & "' order by Numero_consultoriaHist asc", db, adOpenKeyset, adLockOptimistic, adCmdText
 Set Ado_datos2.Recordset = rs_datos2
 Set dg_det2.DataSource = Ado_datos2.Recordset
 mover = 1
' Call ABRIR_TABLA_DET(1)
''dg_det1.SelBookmarks.Remove (0)
''dg_det1.ClearFields
' mover = 1
'Me.dgv.Currentcell = Nothing

 If (dg_det1.SelBookmarks.Count <> 0) Then
            dg_det1.SelBookmarks.Remove 0
 End If
If rs_datos2.RecordCount > 0 Then

rs_datos1.Find "unidad_codigo_pla = '" & rs_datos2!unidad_codigo & "'", , , 1

dg_det1.SelBookmarks.Add (rs_datos1.Bookmark)
 
 Else
 sino = MsgBox("No se encontro a nadie con ese nombre", vbInformation, "Aviso")
 Call ABRIR_TABLA_DET(2)
 dtc_buscar_desc.Text = ""
 End If
End Select
   
   
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
dtc_rep_cod.Text = "%"
dtc_rep_det.Text = "TODAS LAS PLANILLAS"
Else
dtc_rep_cod.Text = ""
dtc_rep_det.Text = ""
End If
End Sub

Private Sub Picture13_Click()
fra_sub_grupo_unidad.Visible = False
FraNavega.Enabled = True
       fraOpciones.Enabled = True
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
         dg_det1.Enabled = True
         fra_opciones_det_1.Enabled = True
fra_opciones_det_2.Enabled = True

        dg_det2.Enabled = True
End Sub

Private Sub BtnNuevo14_Click()
 If rs_datos!estado_codigo = "REG" Then
    If Ado_datos.Recordset.RecordCount > 0 Then
        If Ado_datos1.Recordset.RecordCount > 0 Then
            Call ABRIR_TABLAS_AUX(3)
            Numero = 0
            On Error GoTo UpdateErr
            ' If rs_datos1!estado_codigo = "APR" Then
            sino = MsgBox("¿Desea que el sistema genere autamáticamente la planilla ?", vbYesNo + vbQuestion, "Atención")
            If sino = vbYes Then
                ProgressBar1.Visible = True
                If rs_aux6.State = 1 Then rs_aux6.Close
                rs_aux6.Open "SELECT * FROM ro_personal_contratado WHERE unidad_codigo_pla = '" & Ado_datos1.Recordset!unidad_codigo_pla & "' and estado_codigo <> 'ANL' AND estado_jubilado = 'REG'", db, adOpenKeyset, adLockOptimistic 'adOpenStatic 'order by beneficiario_denominacion
                'rs_aux6.Open "SELECT * FROM av_ro_peronal_vs_gc_beneficiario  WHERE unidad_codigo = '" & rs_datos1!unidad_codigo_pla & "' AND estado_codigo = 'APR' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic
                If rs_aux6.RecordCount > 0 Then 'verifica si existe personal en esa sub_planilla
                    rs_aux6.MoveFirst
                    With ProgressBar1
                      .Max = rs_aux6.RecordCount
                      .Min = 0
                      .Value = 0
                    End With
                    'ProgressBar1.Max =
                    While Not rs_aux6.EOF
'        If rs_aux6!beneficiario_codigo = "159256" Then
'        sino = ""
'        End If
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    DIA_FN = Day(rs_aux6!fecha_expiracion) 'FECHA FIN
                    MES_FN = Month(rs_aux6!fecha_expiracion)
                    ANO_FN = Year(rs_aux6!fecha_expiracion)
                    If rs_aux6!beneficiario_codigo = "9895734" Then
                        sino = ""
                    End If
   expira = Day(DateSerial(rs_datos!ges_gestion, rs_datos!mes_grupo + 1, 0))
     'SE CONMPONE LA FECHA PARA COMPARAR CON EL FIN DEL CONTRATO
     FIN = "1" & "/" & rs_datos!mes_grupo & "/" & rs_datos!ges_gestion
     inicio = DateSerial(Year(FIN), Month(FIN) + 1, 0)
    'COMPARACION DE FECHA DE FIN DE CONTRATO
    If rs_aux6!fecha_expiracion >= FIN Then 'PARA LAS BAJAS
      If rs_aux6!fecha_ingreso <= inicio Then 'INICIO CONTRATO
        If rs_aux5.State = 1 Then rs_aux5.Close
'        rs_aux5.Open "select * from ro_pagos_cronograma_detalle where ges_gestion = '" & Ado_datos1.Recordset!ges_gestion & "' AND planilla_codigo = '" & Ado_datos1.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos1.Recordset!mes_grupo & " AND numero_pago = " & Ado_datos1.Recordset!NUMERO_PAGO & " AND beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "' AND unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo_pla & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        'VERIFICA SI YA EXISTE EN LA PLANILLA LA PERSONA
        rs_aux5.Open "select * from ro_pagos_cronograma_detalle where ges_gestion = '" & Ado_datos1.Recordset!ges_gestion & "' AND unidad_codigo = '" & rs_datos1!unidad_codigo_pla & "' AND planilla_codigo = '" & Ado_datos1.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos1.Recordset!mes_grupo & " AND numero_pago = " & Ado_datos1.Recordset!NUMERO_PAGO & " AND beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
        If rs_aux5.RecordCount = 0 Then 'LA PERSONA NO ESTA GENERADA
            
            'GUARDADO DE DATOS
        NUM_PERS = NUM_PERS + 1
            
'        If rs_aux6!beneficiario_codigo = "4889246" Then 'para pruebas
'        sino = ""
'        End If
        
        Dim persona As New ADODB.Recordset
        If persona.State = 1 Then persona.Close
        persona.Open "select * from gc_beneficiario where beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
            
            
           
            
            rs_datos2.AddNew
            
            
            rs_datos2!fecha_ingreso = rs_aux6!fecha_ingreso
            rs_datos2!fecha_expiracion = rs_aux6!fecha_expiracion
            rs_datos2!cargo_codigo = rs_aux6!cargo_codigo
            rs_datos2!puesto_codigo = rs_aux6!puesto_codigo
            rs_datos2!beneficiario_haber_mensual = rs_aux6!beneficiario_haber_mensual
            rs_datos2!Unidad = rs_aux6!unidad_codigo
            
            rs_datos2!ges_gestion = rs_datos!ges_gestion
            rs_datos2!planilla_codigo = rs_datos!planilla_codigo
            rs_datos2!mes_grupo = rs_datos1!mes_grupo
            rs_datos2!NUMERO_PAGO = rs_datos1!NUMERO_PAGO
            
            rs_datos2!beneficiario_codigo = rs_aux6!beneficiario_codigo
            VAR_BENEF = LTrim(RTrim(rs_aux6!beneficiario_codigo))
            rs_datos2!unidad_codigo = rs_datos1!unidad_codigo_pla
            rs_datos2!tipo_moneda = "BOB"
            rs_datos2!tipo_cambio = GlTipoCambioOficial
            'Adicionar  beneficiario_haber_mensual_ant WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'
            DIA_IN = Day(rs_aux6!fecha_ingreso)
            MES_IN = Month(rs_aux6!fecha_ingreso)
            ANO_IN = Year(rs_aux6!fecha_ingreso)
            
            DIA_HOY = Day(Now)
              
            DIA_FN = Day(rs_aux6!fecha_expiracion) 'FECHA FIN
            MES_FN = Month(rs_aux6!fecha_expiracion)
            ANO_FN = Year(rs_aux6!fecha_expiracion)
             If DIA_FN > 30 Then
             DIA_FN = 30
             End If
          If rs_aux6!beneficiario_codigo = "2689587" Then
            sino = ""
            End If
          
            'VERIFICA SI ENTRO EN EL MES Y EL AÑO AL QUE CORRESPONDE LA PLANILLA
            If MES_IN = rs_datos2!mes_grupo And ANO_IN = rs_datos2!ges_gestion Then
             'CALCULO PARA PAGAR DE DIAS TRABAJADOS EN EL MES EN CASO DE QUE NO ENTRO EN EL DIA 1 DEL MES
             rs_datos2!sueldo_basico = (rs_aux6!beneficiario_haber_mensual / 30) * (30 - (DIA_IN - 1))
             rs_datos2!dias_trabajados = (30 - (DIA_IN - 1))
             nuevos = nuevos & "    " & persona!beneficiario_codigo & " " & persona!beneficiario_denominacion & vbCrLf & "    Fecha Ingreso: " & rs_aux6!fecha_ingreso & vbCrLf & "    Haber Basico: " & rs_datos2!beneficiario_haber_mensual & vbCrLf & "    Dias Trabajados: " & rs_datos2!dias_trabajados & vbCrLf & "    Haber Basico del Mes: " & rs_datos2!sueldo_basico & vbCrLf & vbCrLf
             
            Else
              'SI ESTRO EN ANTES DE EL MES DE LA PLANILLA
              rs_datos2!sueldo_basico = rs_aux6!beneficiario_haber_mensual
              rs_datos2!dias_trabajados = "30"
            End If
             'VERIFICA SI SU CONTRATO EXPIRA EN EL MES DE LA PLANILLA
              If MES_FN = rs_datos2!mes_grupo And ANO_FN = rs_datos2!ges_gestion Then 'FECHA FIN
              
             'CALCULO DE PAGO POR DIAS TRABAJADOS
              rs_datos2!sueldo_basico = (rs_aux6!beneficiario_haber_mensual / 30) * (DIA_FN)
              rs_datos2!dias_trabajados = DIA_FN
              expirados = expirados & "    " & persona!beneficiario_codigo & " " & persona!beneficiario_denominacion & vbCrLf & "    Fecha Fin: " & rs_aux6!fecha_expiracion & vbCrLf & "    Haber Basico: " & rs_datos2!beneficiario_haber_mensual & vbCrLf & "    Dias Trabajados: " & rs_datos2!dias_trabajados & vbCrLf & "    Haber Basico del Mes: " & rs_datos2!sueldo_basico & vbCrLf & vbCrLf
           
             End If
            
           sino = persona.RecordCount
            
            
            
            rs_datos2!monto_refrigerio = IIf(IsNull(rs_aux6!beneficiario_otro_mensual), "0", rs_aux6!beneficiario_otro_mensual)

             'PONE EN ULTIMO DIA DEL MES PARA COMPARAR ANTIGUEDAD
             fecha_pla = DateSerial(rs_datos!ges_gestion, rs_datos!mes_grupo + 1, 1 - 1)
             'fecha_pla = "29/02/2016" SOLO PARA PRUEBAS
             'ABRE TABLA DONDE ESTAN LOS PARAMETROS DE ANTIGUEDAD
             If rs_aux8.State = 1 Then rs_aux8.Close
             rs_aux8.Open "select * from rc_antiguedad", db, adOpenKeyset, adLockOptimistic, adCmdText
             rs_aux8.MoveFirst
             While Not rs_aux8.EOF
             'GUARDA LAS FECHA MINIMA Y LA MAXIMA EN LA QUE DEBERIA ENTRAR LA PERSONA PARA RECIBIR EL BONO ANTIGUEDAD
             f1 = DateAdd("yyyy", -rs_aux8!parametro_inicial, CDate(fecha_pla))
             f2 = DateAdd("yyyy", -rs_aux8!parametro_final, CDate(fecha_pla))
             'COMPARA LA FECHA INGRESO CON EL PARAMETRO
             If rs_aux6!fecha_ingreso <= CDate(f1) And rs_aux6!fecha_ingreso > CDate(f2) Then
             'GUARDA EL MONTO QUE CORRESPONDE
             rs_datos2!bono_antiguedad = rs_aux8!antig_valor
             rs_aux8.MoveLast
             End If
             rs_aux8.MoveNext
             Wend
             
             
                rs_datos2!bono_transporte = 0
    
   ' End If
            
            'rs_datos2!horas_extras = dtc_refrigerio.Text
            'rs_datos2!bono_transporte = dtc_refrigerio.Text
             'rs_datos2!total_ganado = rs_datos2!sueldo_basico + rs_datos2!monto_refrigerio + rs_datos2!bono_antiguedad
             
            rs_datos2!total_ganado = rs_datos2!sueldo_basico + rs_datos2!bono_antiguedad + rs_datos2!bono_transporte
            rs_datos2!provision_aguinaldo = rs_datos2!total_ganado * 0.0833
            rs_datos2!prevision_indemnizacion = rs_datos2!total_ganado * 0.0833
            rs_datos2!anticipo_sueldo = "0"
            rs_datos2!anticipo_refrigerio = "0"
            
            If VAR_BENEF = "3395947" Then 'SOLO PARA PRUEBAS
            sino = ""
            End If
            
            'VARIABLE PARA SUMA DE PAGOS DE PRESTAMO
            PRESTAMO_TOTAL = 0
            Set rs_aux24 = New Recordset
            If rs_aux24.State = 1 Then rs_aux24.Close
            'VERIFICA SI TIENE ALGUN PRESTAMO LA PERSONA
            rs_aux24.Open "select * from ro_prestamos where beneficiario_codigo = '" & VAR_BENEF & "'  and estado_codigo = 'APR'  and codigo_empresa = " & rs_aux6!codigo_empresa & "", db, adOpenKeyset, adLockOptimistic, adCmdText
            If rs_aux24.RecordCount > 0 Then
            rs_aux24.MoveFirst
            While Not rs_aux24.EOF
                If rs_aux24!estado_codigo = "APR" Then
                    Set rs_aux25 = New Recordset
                    If rs_aux25.State = 1 Then rs_aux25.Close
                    'VERIFICA SI EXISTE PAGO PARA ESTE MES SEGUN EL CRONOGRANA GENERADO
                    rs_aux25.Open "select * from ro_prestamo_prog where beneficiario_codigo = '" & VAR_BENEF & "' and prestamo_codigo = " & rs_aux24!prestamo_codigo & " AND mes_planilla = " & rs_datos!mes_grupo & " and year(cobranza_fecha_prog) = '" & rs_datos!ges_gestion & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
                    If rs_aux25.RecordCount > 0 Then
                    'SUMA LOS PAGOS
                    PRESTAMO_TOTAL = PRESTAMO_TOTAL + rs_aux25!cobranza_programada_bs
                    'APRUEBA EL PAGO
                    rs_aux25!estado_codigo = "APR"
                    rs_aux25!cobranza_fecha_cobro = Date
                    rs_aux25.Update
                    
                    rs_aux24!correl_prog = rs_aux25!prestamo_prog_codigo
                    Set rs_aux26 = New Recordset
                    If rs_aux26.State = 1 Then rs_aux26.Close
                    'SUMA TODOS LOS PAGOS REALIZADOS PARA GUARDAR EN LA CABECERA
                    rs_aux26.Open "select SUM(cobranza_programada_bs)AS TOTAL_COB from ro_prestamo_prog where beneficiario_codigo = '" & VAR_BENEF & "' and estado_codigo = 'APR' AND prestamo_codigo = " & rs_aux24!prestamo_codigo, db, adOpenKeyset, adLockOptimistic, adCmdText
                    rs_aux24!total_cobrado = rs_aux26!TOTAL_COB
                     
                    End If
               End If
                rs_aux24.MoveNext
            Wend
            End If
            rs_datos2!prestamo = PRESTAMO_TOTAL
            'CALCULO DE AFP
            Select Case rs_aux6!beneficiario_codigo_afp
                Case "1006803"      'AFP1
                    rs_datos2!afp1 = rs_datos2!total_ganado * 0.1271
                    rs_datos2!afp2 = "0"       'falta 987654
                    VAR_NETO = rs_datos2!total_ganado - rs_datos2!afp1
                Case "987654"       'AFP2
                    rs_datos2!afp1 = "0"       'falta 1006803
                    rs_datos2!afp2 = rs_datos2!total_ganado * 0.1271
                    VAR_NETO = rs_datos2!total_ganado - rs_datos2!afp2
                Case Else
                    rs_datos2!afp1 = "0"
                    rs_datos2!afp2 = "0"
                    VAR_NETO = rs_datos2!total_ganado
            End Select
             '
'            VAR_IVA = 1805 * 4
'            If VAR_NETO > VAR_IVA Then
'                rs_datos2!rciva = rs_datos2!total_ganado * 0.13
'            Else
'                rs_datos2!rciva = "0"        'mayor a 4 SUELDOS BASICOA
'            End If
            '
            db.Execute "UPDATE ro_controlasistencia SET ges_gestion = year(Fecha_control), Mes_control = month(Fecha_control), Dia_control= day(Fecha_control)"
            'sqlAux = "SELECT '     TOTAL MINUTOS DE RETRASO: ' + CONVERT(VARCHAR, ISNULL(SUM(DATEDIFF(MINUTE, '0:00:00', Tardanza)),0)) AS totHrs FROM ro_controlasistencia WHERE beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' "
            'rs_AsisTT.Open sqlAux, db, adOpenKeyset, adLockOptimistic, adCmdText
            'rs_AsisTT.MoveFirst
            'AdoAsistencia.Caption = CStr(rs_AsisTT!totHrs)
            '
           
           
            'db.Execute "UPDATE ro_controlasistencia SET TotalMin1 = convert(int,TardanzaCadena) "
            'rs_aux9.Open "select sum(convert(int,TardanzaCadena)) as TardanzaMes from ro_controlasistencia where beneficiario_codigo = " & VAR_BENEF & " and Mes_control = '" & Str(Ado_datos1.Recordset!mes_grupo) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
             'Dim rs_aux9 As New ADODB.Recordset
            If rs_aux9.State = 1 Then rs_aux9.Close
            rs_aux9.Open "select sum(AtrasoMin1) as TardanzaMes from ro_controlasistencia where beneficiario_codigo = '" & RTrim(LTrim(VAR_BENEF)) & "' AND ges_gestion = '" & RTrim(LTrim(Ado_datos1.Recordset!ges_gestion)) & "' and Mes_control = '" & RTrim(LTrim(Str(Ado_datos1.Recordset!mes_grupo))) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
             'select sum(convert(int,TardanzaCadena)) as TardanzaMes from ro_controlasistencia where beneficiario_codigo = '6960987' and Mes_control = 7
            If rs_aux14.State = 1 Then rs_aux14.Close
            mesnom = UCase(MonthName(Ado_datos1.Recordset!mes_grupo))
            rs_aux14.Open "select sum(total_minuto) as PermisoMes from ro_permisos where beneficiario_codigo = '" & RTrim(LTrim(VAR_BENEF)) & "' AND ges_gestion = '" & RTrim(LTrim(Ado_datos1.Recordset!ges_gestion)) & "' AND Mes_control = '" & mesnom & "' AND estado_codigo = 'APR' and dias_permiso = 0", db, adOpenKeyset, adLockOptimistic, adCmdText
            If rs_aux14!PermisoMes <> "NULL" Then
                permisos = rs_aux14!PermisoMes
            Else
                permisos = "0"
            End If
            If rs_aux9!TardanzaMes <> "NULL" Then
                totalminutos = rs_aux9!TardanzaMes - permisos
                If totalminutos >= 45 And totalminutos <= 60 Then
                    rs_datos2!otros_dsctos = ((rs_datos2!sueldo_basico / 30) / 2)
                Else
                    If totalminutos >= 61 And totalminutos <= 75 Then
                        rs_datos2!otros_dsctos = (rs_datos2!sueldo_basico / 30)
                    Else
                        If totalminutos >= 76 And totalminutos <= 105 Then
                            rs_datos2!otros_dsctos = ((rs_datos2!sueldo_basico / 30) * 2)
                        Else
                            If totalminutos >= 106 Then
                                rs_datos2!otros_dsctos = ((rs_datos2!sueldo_basico / 30) * 3)
                            Else
                                rs_datos2!otros_dsctos = 0
                            End If
                        End If
                    End If
                End If
            Else
              If continuar = "SI" Then
                sino = MsgBox("No se Cargo la asistencia del mes de " & UCase(MonthName(rs_datos1!mes_grupo)) & " de algunas personas " & vbCrLf & "¿Desea generar de todas maneras?" & vbCrLf & "NOTA: En el campo de OTROS DESCUENTOS se asignará cero (0) por defecto", vbCritical + vbYesNo, "Atención")
                If sino = vbYes Then
                    rs_datos2!otros_dsctos = 0
                    continuar = "NO"
                    Numero = Numero + 1
                Else
                    ProgressBar1.Visible = False
                    Fra_personal_Ppla.Visible = False
                    FraNavega.Enabled = True
                    fraOpciones.Enabled = True
                    ' FraGrabarCancelar.Visible = True
                    dg_datos.Enabled = True
                    dg_det1.Enabled = True
                    fra_opciones_det_1.Enabled = True
                    fra_opciones_det_2.Enabled = True
        
                    dg_det2.Enabled = True
                    Call ABRIR_TABLA_DET(2)
                    Exit Sub
                End If
              Else
                rs_datos2!otros_dsctos = 0
                Numero = Numero + 1
              End If
            End If
            'rs_datos2!otros_dsctos = "0"   'FIN Atrasos y Faltas
            rs_datos2!r_provision_aguinaldo = "0"
            rs_datos2!r_prevision_indemnizacion = "0"
            
              If rs_aux15.State = 1 Then rs_aux15.Close
              rs_aux15.Open "select SUM(monto) AS totalmonto, SUM(dias) AS Totaldias from ro_memorandas where ges_gestion = '" & Ado_datos1.Recordset!ges_gestion & "' AND mes_descuento = " & Ado_datos1.Recordset!mes_grupo & " AND beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "' AND descuento_pla = 'SI' AND estado_codigo = 'APR' AND tipo_memo <> 'ANT' AND codigo_empresa = " & rs_aux6!codigo_empresa & "", db, adOpenKeyset, adLockOptimistic, adCmdText
             
         If rs_aux15.RecordCount <> 0 Then
         
           
              If rs_aux15!totalmonto > 0 Then
                total = rs_datos2!otros_dsctos + IIf(IsNull(rs_aux15!totalmonto), 0, rs_aux15!totalmonto)
               rs_datos2!otros_dsctos = total
              End If
              
              If rs_aux15!Totaldias > 0 Then
                total = rs_datos2!otros_dsctos + ((rs_aux6!beneficiario_haber_mensual / 30) * rs_aux15!Totaldias)
                'total = total + rs_datos2!otros_dsctos
                rs_datos2!otros_dsctos = total
              End If
              
         End If
         If rs_aux30.State = 1 Then rs_aux30.Close
             
             rs_aux30.Open "select SUM(monto) AS totalmonto, SUM(dias) AS Totaldias from ro_memorandas where ges_gestion = '" & Ado_datos1.Recordset!ges_gestion & "' AND mes_descuento = " & Ado_datos1.Recordset!mes_grupo & " AND beneficiario_codigo = '" & rs_aux6!beneficiario_codigo & "' AND descuento_pla = 'SI' AND estado_codigo = 'APR' AND tipo_memo = 'ANT' AND codigo_empresa = " & rs_aux6!codigo_empresa & "", db, adOpenKeyset, adLockOptimistic, adCmdText
             
         If rs_aux30.RecordCount <> 0 Then
         
              
              If rs_aux30!totalmonto > 0 Then
                total = rs_datos2!anticipo_sueldo + IIf(IsNull(rs_aux30!totalmonto), 0, rs_aux30!totalmonto)
               rs_datos2!anticipo_sueldo = total
              End If
              
'              If rs_aux30!Totaldias > 0 Then
'                total = rs_aux30!anticipo_sueldo + ((rs_aux6!beneficiario_haber_mensual / 30) * rs_aux30!Totaldias)
'                'total = total + rs_datos2!otros_dsctos
'                rs_datos2!anticipo_sueldo = total
'              End If
              
          End If
            'rs_datos2.Update
            'rs_datos2!total_dsctos = "0"
            rs_datos2!total_dsctos = rs_datos2!anticipo_sueldo + rs_datos2!anticipo_refrigerio + rs_datos2!prestamo + rs_datos2!afp1 + rs_datos2!afp2 + rs_datos2!otros_dsctos
                        
            rs_datos2!liquido_pagable_bs = rs_datos2!total_ganado - rs_datos2!total_dsctos
            rs_datos2!liquido_pagable_us = rs_datos2!liquido_pagable_bs / GlTipoCambioOficial
             'rs_datos2!total_dsctos = "0"
            rs_datos2!emite_factura = "N"
             
            rs_datos2!cite_conformidad = "-"
            'rs_datos2!Numero_consultoriaHist = " "
            rs_datos2!fte_financiamientoHist = "-"
            rs_datos2!estado_devengado = "REG"
             '70522788
            rs_datos2!estado_codigo = "REG"
            rs_datos2!fecha_registro = Date
            rs_datos2!usr_codigo = glusuario
            
            rs_datos2!iva_110 = "0"
            rs_datos2!fisco_a_favor = "0"
            rs_datos2!dependiente_a_favor = "0"
            rs_datos2!mes_anterior = "0"
            rs_datos2!mes_anterior_mant = "0"
            rs_datos2!saldo_util = "0"
            rs_datos2!saldo_a_favor_depend = "0"
            rs_datos2!rciva = "0"
            rs_datos2!codigo_empresa = rs_aux6!codigo_empresa
            'ABRIR_TABLA_DET (2)
            rs_datos2!hora_registro = Format(Time, "HH:mm:ss")
            
            rs_datos2.Update
            'Call OptFilGral1_Click
            'rs_datos.MoveLast
            mbDataChanged = False

    
        End If
        End If
   Else 'PARA LAS BAJAS
    rs_aux6!estado_codigo = "REG"
    End If 'PARA LAS BAJAS
        rs_aux6.MoveNext
       Wend
  Else
  sino = MsgBox("No existe personal en esta planilla", vbInformation, "Atención")
   End If 'verifica si existe personal en esa sub_planilla
     
       Call ABRIR_TABLA_DET(2)
       If Ado_datos2.Recordset.RecordCount > 0 Then
       Call numeracion_planilla
       End If
       'rs_datos2.RecordCount
       
   'sino = MsgBox("Se genero correctamente la planilla", vbInformation, "Atención")
    continuar = "SI"
    ProgressBar1.Visible = False
    dtc_buscar_desc.Visible = True
    Label52.Visible = True
    
'        Else
'        sino = MsgBox("ya se GENERO anteriormente Esta Planilla ", vbCritical, "Atención")
'       End If
  Else
    txt_anticipo_sb.Visible = False
    txt_anticipo_refr.Visible = False
    txt_rc_iva.Visible = False
    txt_prestamo.Visible = False
    txt_otros_descuentos.Visible = False
    txt_iva_110.Visible = False
    txt_fisco_a_favor.Visible = False
    txt_dependiente_a_favor.Visible = False
    txt_mes_anterior.Visible = False
    txt_mes_anterior_mant.Visible = False
    txt_saldo_util.Visible = False
    txt_saldo_a_favor_depend.Visible = False
    
       
    txt_anticipo_sb1.Visible = True
    txt_anticipo_refr1.Visible = True
   ' txt_rc_iva1.Visible = True
    txt_prestamo1.Visible = True
    txt_otros_descuentos1.Visible = True
    txt_iva_1101.Visible = True
    txt_fisco_a_favor1.Visible = True
    txt_dependiente_a_favor1.Visible = True
    txt_mes_anterior1.Visible = True
    txt_mes_anterior_mant1.Visible = True
    txt_saldo_util1.Visible = True
    txt_saldo_a_favor_depend1.Visible = True
      
      
      
    txt_anticipo_sb1.Text = 0
    txt_anticipo_refr1.Text = 0
    'txt_rc_iva1.Text = 0
    txt_prestamo1.Text = 0
    txt_otros_descuentos1.Text = 0
    txt_iva_1101.Text = "0"
    txt_fisco_a_favor1.Text = "0"
    txt_dependiente_a_favor1.Text = "0"
    txt_mes_anterior1.Text = "0"
    txt_mes_anterior_mant1.Text = "0"
    txt_saldo_util1.Text = "0"
    txt_saldo_a_favor_depend1.Text = "0"
      
      
      
    VAR_SW = "SI"
    Fra_personal_Ppla.Visible = True
    txt_anticipo_sb.Text = "0"
    txt_anticipo_refr.Text = "0"
    txt_rc_iva.Text = "0"
    txt_afp1.Text = "0"
    txt_afp2.Text = "0"
    txt_prestamo.Text = "0"
    'txt_otros_descuentos.Text = "0" ' HABILITAR DESPUES
    txt_bono_ant.Text = "0"
     txt_nomb_ap.Visible = False
     txt_ci.Visible = False
      txt_sueldo.Visible = False
       txt_refri.Visible = False
       
       
        txt_iva_110.Text = "0"
    txt_fisco_a_favor.Text = "0"
    txt_dependiente_a_favor.Text = "0"
    txt_mes_anterior.Text = "0"
    txt_mes_anterior_mant.Text = "0"
    txt_saldo_util.Text = "0"
    txt_saldo_a_favor_depend.Text = "0"
    txt_total_ganado.Text = "0"
    txt_total_descuento.Text = "0"
       txt_liq_pagable.Text = "0"
    dtc_codigo.Visible = True
     dtc_descripcion.Visible = True
      dtc_sueldo.Visible = True
       dtc_refrigerio.Visible = True
       
    dtc_codigo.Text = ""
     dtc_descripcion.Text = ""
      dtc_sueldo.Text = "0"
       dtc_refrigerio.Text = "0"
       txt_total_ganado.Text = "0"
        txt_total_descuento.Text = "0"
         txt_liq_pagable.Text = "0"
           
           
           FraNavega.Enabled = False
       fraOpciones.Enabled = False
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
         dg_det1.Enabled = False
         fra_opciones_det_1.Enabled = False
fra_opciones_det_2.Enabled = False

        dg_det2.Enabled = False
        
  End If
'       Else
'       sino = MsgBox("No se puede generar si el GRUPO POR UNIDAD no esta aprobado", vbCritical, "Atención")
'End If
 Else
sino = MsgBox("Primero tiene que crear el Sub Grupo Por Unidad", vbCritical, "Atención")
End If
 Exit Sub

UpdateErr:
  
  MsgBox Err.Description
  Else
sino = MsgBox("Primero tiene que crear el Sub Grupo Por Unidad", vbCritical, "Atención")
End If
Else   'APR
      MsgBox "La planilla ya fue APROBADA no se puede hacer ningun cambio", vbExclamation, "Error"
End If 'REG

End Sub


Private Sub Picture15_Click()
If rs_datos!estado_codigo = "REG" Then
 'txt_otros_descuentos.Text = "0" ' TEMPORAL
If Ado_datos.Recordset.RecordCount > 0 Then
If Ado_datos2.Recordset.RecordCount > 0 Then
Call ABRIR_TABLAS_AUX(3)
'dtc_codigo.Text = "0"
'dtc_sueldo.Text = "0"
'txt_bono_ant.Text = "0"
'dtc_refrigerio.Text = "0"
'txt_total_ganado.Text = "0"
'txt_total_ganado.Text = "0"
'txt_anticipo_sb.Text = "0"
'txt_anticipo_sb.Text = "0"
'txt_prestamo.Text = "0"
'txt_afp1.Text = "0"
'txt_afp2.Text = "0"
'txt_rc_iva.Text = "0"
'txt_otros_descuentos.Text = "0"
'txt_total_descuento.Text = "0"
'txt_liq_pagable.Text = "0"
'txt_total_descuento.Text = "0"
 
     txt_anticipo_sb.Visible = True
    txt_anticipo_refr.Visible = True
     txt_rc_iva.Visible = True
     txt_prestamo.Visible = True
    txt_otros_descuentos.Visible = True
        txt_iva_110.Visible = True
    txt_fisco_a_favor.Visible = True
    txt_dependiente_a_favor.Visible = True
    txt_mes_anterior.Visible = True
    txt_mes_anterior_mant.Visible = True
    txt_saldo_util.Visible = True
    txt_saldo_a_favor_depend.Visible = True
    
    
      txt_anticipo_sb1.Visible = False
    txt_anticipo_refr1.Visible = False
     'txt_rc_iva1.Visible = False
     txt_prestamo1.Visible = False
    txt_otros_descuentos1.Visible = False
        txt_iva_1101.Visible = False
    txt_fisco_a_favor1.Visible = False
    txt_dependiente_a_favor1.Visible = False
    txt_mes_anterior1.Visible = False
    txt_mes_anterior_mant1.Visible = False
    txt_saldo_util1.Visible = False
    txt_saldo_a_favor_depend1.Visible = False
    
    
    
     
If rs_datos2.RecordCount > 0 Then
If Ado_datos2.Recordset!estado_codigo = "REG" Then
         FraNavega.Enabled = False
       fraOpciones.Enabled = False
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
         dg_det1.Enabled = False
         fra_opciones_det_1.Enabled = False
          fra_opciones_det_2.Enabled = False
        dg_det2.Enabled = False
VAR_SW = ""
  txt_nomb_ap.Visible = True
     txt_ci.Visible = True
      txt_sueldo.Visible = True
       txt_refri.Visible = True
    dtc_codigo.Visible = False
     dtc_descripcion.Visible = False
      dtc_sueldo.Visible = False
       dtc_refrigerio.Visible = False
Fra_personal_Ppla.Visible = True
'dtc_sueldo.Text = rs_datos2!sueldo_basico
'txt_bono_ant.Text = rs_datos2!bono_antiguedad
'dtc_refrigerio.Text = rs_datos2!monto_refrigerio
' 'rs_datos2!horas_extras = dtc_refrigerio.Text
' 'rs_datos2!bono_transporte = dtc_refrigerio.Text
'   txt_total_ganado.Text = rs_datos2!total_ganado
'   txt_total_ganado.Text = rs_datos2!total_ganado
'   txt_anticipo_sb.Text = rs_datos2!anticipo_sueldo
'   txt_anticipo_sb.Text = rs_datos2!anticipo_refrigerio
'   txt_prestamo.Text = rs_datos2!prestamo
'   txt_afp1.Text = rs_datos2!afp1
'   txt_afp2.Text = rs_datos2!afp2
'   txt_rc_iva.Text = rs_datos2!rciva
'   txt_otros_descuentos.Text = rs_datos2!otros_dsctos
'   txt_total_descuento.Text = rs_datos2!total_dsctos
'   txt_liq_pagable.Text = rs_datos2!liquido_pagable_bs
'   'rs_datos2!liquido_pagable_us = (Val(txt_liq_pagable.Text) / GlTipoCambioOficial)
'   txt_total_descuento.Text = rs_datos2!total_dsctos
    Else
      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
End If
 
    End If
  Exit Sub

EditErr:
  MsgBox Err.Description
  Else
      MsgBox "No existen registros", vbExclamation, "Error"
 End If
Else
      MsgBox "No existen registros", vbExclamation, "Error"
End If
Else    'APR
      MsgBox "La planilla ya fue APROBADA no se puede hacer ningun cambio", vbExclamation, "Error"
End If 'REG
End Sub

Private Sub Picture16_Click()
If rs_datos!estado_codigo = "REG" Then
If Ado_datos.Recordset.RecordCount > 0 Then
If Ado_datos2.Recordset.RecordCount > 0 Then
On Error GoTo UpdateErr

   If rs_datos2!estado_codigo = "APR" Or rs_datos2!estado_devengado = "APR" Then
      sino = MsgBox("No se puede ELIMINAR porque el Registro ya fue utilizado. Desea marcar como ERRADO ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos2!estado_codigo = "ERR"
         rs_datos2!fecha_registro = Date
         rs_datos2!usr_codigo = glusuario
         rs_datos2.UpdateBatch adAffectAll
      End If
   Else
      sino = MsgBox("¿Desea ELIMINAR TODOS los registro? ", vbYesNo + vbQuestion, "Atención") 'unidad_codigo_pla
      If sino = vbYes Then
         db.Execute "DELETE ro_pagos_cronograma_Detalle where ges_gestion = " & Ado_datos2.Recordset!ges_gestion & " AND planilla_codigo = '" & Ado_datos2.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos2.Recordset!mes_grupo & " AND unidad_codigo = '" & Ado_datos2.Recordset!unidad_codigo & "' AND numero_pago = " & Ado_datos2.Recordset!NUMERO_PAGO & ""
      Else
        sino = MsgBox("¿Desea ELIMINAR a Esta Persona? " & vbCrLf & rs_datos2!beneficiario_denominacion, vbYesNo + vbQuestion, "Atención")  'unidad_codigo_pla
      If sino = vbYes Then
         db.Execute "DELETE ro_pagos_cronograma_Detalle where ges_gestion = " & Ado_datos2.Recordset!ges_gestion & " AND planilla_codigo = '" & Ado_datos2.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos2.Recordset!mes_grupo & " AND unidad_codigo = '" & Ado_datos2.Recordset!unidad_codigo & "' AND numero_pago = " & Ado_datos2.Recordset!NUMERO_PAGO & " AND beneficiario_codigo = '" & Ado_datos2.Recordset!beneficiario_codigo & "' and codigo_empresa = " & Ado_datos2.Recordset!codigo_empresa & ""
      Else
      Exit Sub
      End If
      End If
   End If
   Call ABRIR_TABLA_DET(2)
   Call ABRIR_TABLAS_AUX(5)
   Call numeracion_planilla
'    dtc_buscar_desc.Text = ""
'    dtc_buscar_ci.Text = ""
   Exit Sub

UpdateErr:
  MsgBox Err.Description
  Else
      MsgBox "No existen registros", vbExclamation, "Error"
End If
Else
      MsgBox "No existen registros", vbExclamation, "Error"
End If
Else    'APR
      MsgBox "La planilla ya fue APROBADA no se puede hacer ningun cambio", vbExclamation, "Error"
End If 'REG
End Sub

Private Sub BtnGrabar18_Click()
On Error GoTo UpdateErr
  sino = MsgBox("¿Está Seguro de generar Las Planillas con los siguiente Datos?" & vbCrLf & "Gestión = " & cmb_gestion.Text & vbCrLf & "Nro. De planillas + Aguinaldo = " & cmb_nro_planillas.Text, vbYesNo + vbQuestion, "Atención")
  If sino = vbYes Then
    Dim contador As Integer
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "select * from ro_pagos_grupos where ges_gestion = '" & cmb_gestion.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rs_aux1.RecordCount >= 0 Then
        contador = 0
        'contador = 3
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from rc_planilla_grupo where estado_codigo = 'APR'", db, adOpenKeyset, adLockOptimistic
        rs_aux2.MoveFirst
        ProgressBar3.Visible = True
        With ProgressBar3
            .Max = Val(cmb_nro_planillas.Text)
            .Min = 0
            .Value = "0"
        End With
        While Not rs_aux2.EOF
            While (contador <> cmb_nro_planillas.Text)
                contador = contador + 1
                ProgressBar3.Value = contador
                If contador > 12 Then
                    db.Execute "Insert INTO ro_pagos_grupos (ges_gestion, planilla_codigo, mes_grupo, descripcion_grupo, unidad_codigo, depto_codigo, clasif_codigo,doc_codigo,solicitud_tipo, estado_codigo, usr_codigo, fecha_registro) Values ('" & cmb_gestion.Text & "', '" & rs_aux2!planilla_codigo & "', " & contador & ", '" & rs_aux2!planilla_descripcion & " - AGUINALDO " & (contador - 12) & "', 'RRHH', '" & rs_aux2!depto_codigo & "', 'ADM','R-121', '11','REG', '" & glusuario & "',  '" & Date & "')"
                Else
                    db.Execute "Insert INTO ro_pagos_grupos (ges_gestion, planilla_codigo, mes_grupo, descripcion_grupo, unidad_codigo, depto_codigo, clasif_codigo,doc_codigo,solicitud_tipo, estado_codigo, usr_codigo, fecha_registro) Values ('" & cmb_gestion.Text & "', '" & rs_aux2!planilla_codigo & "', " & contador & ", '" & rs_aux2!planilla_descripcion & " - " & UCase(MonthName(contador)) & "', 'RRHH', '" & rs_aux2!depto_codigo & "', 'ADM','R-121', '11','REG', '" & glusuario & "',  '" & Date & "')"
                End If
            Wend
            rs_aux2.MoveNext
            contador = 0
            'contador = 3
        Wend
        sino = MsgBox("Se generaron correctamente las planillas", vbInformation, "Atención")
        ProgressBar3.Visible = False
        Call OptFilGral1_Click
        fra_generar.Visible = False
        FraNavega.Enabled = True
        fraOpciones.Enabled = True
        ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
        dg_det1.Enabled = True
        fra_opciones_det_1.Enabled = True
        fra_opciones_det_2.Enabled = True
    Else
        sino = MsgBox("ya se GENERARON anteriormente las planillas para la GESTIÓN " & cmb_gestion.Text, vbCritical, "Atención")
    End If
  End If
     Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub Picture20_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
If Ado_datos2.Recordset.RecordCount > 0 Then
   Dim iResult As Integer
   CR02.WindowShowPrintSetupBtn = True
   CR02.WindowShowRefreshBtn = True
   var_titulo = "MODULO RECURSOS HUMANOS"
   VAR_SubTitulo = "BOLETA DE PAGO"
   CR02.ReportFileName = App.Path & "\REPORTES\RRHH\rr_boleta_pago.rpt"
   CR02.Formulas(0) = "Titulo = '" & var_titulo & "' "
   CR02.Formulas(1) = "SubTitulo = '" & VAR_SubTitulo & "' "
   ' CR02.Formulas(2) = "periodo = '" & lbl_texto2 & "' "
  
   CR02.StoredProcParam(0) = Ado_datos2.Recordset!ges_gestion
   CR02.StoredProcParam(1) = Ado_datos2.Recordset!planilla_codigo
   CR02.StoredProcParam(2) = Ado_datos2.Recordset!mes_grupo
   CR02.StoredProcParam(3) = Ado_datos2.Recordset!NUMERO_PAGO
   CR02.StoredProcParam(4) = Ado_datos2.Recordset!beneficiario_codigo
   iResult = CR02.PrintReport
   If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
   CR02.WindowState = crptMaximized
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If

End Sub

Private Sub Picture24_Click()
fra_generar.Visible = False
  FraNavega.Enabled = True
       fraOpciones.Enabled = True
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
         dg_det1.Enabled = True
         fra_opciones_det_1.Enabled = True
fra_opciones_det_2.Enabled = True
End Sub

Private Sub Picture26_Click()
Fra_modificar.Visible = False
FraNavega.Enabled = True
       fraOpciones.Enabled = True
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
         dg_det1.Enabled = True
         fra_opciones_det_1.Enabled = True
fra_opciones_det_2.Enabled = True

        dg_det2.Enabled = True
End Sub


Private Sub Picture27_Click()

  On Error GoTo UpdateErr

    rs_datos!descripcion_grupo = UCase(txt_descripcion_grupo.Text)
        
    rs_datos.Update
    'Call OptFilGral1_Click
     'rs_datos.MoveLast
     mbDataChanged = False
      
        Fra_modificar.Visible = False
        FraNavega.Enabled = True
       fraOpciones.Enabled = True
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
         dg_det1.Enabled = True
         fra_opciones_det_1.Enabled = True
fra_opciones_det_2.Enabled = True

        dg_det2.Enabled = True

'  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub Picture29_Click()
fra_modificar2.Visible = False
 FraNavega.Enabled = True
       fraOpciones.Enabled = True
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
         dg_det1.Enabled = True
         fra_opciones_det_1.Enabled = True
fra_opciones_det_2.Enabled = True

        dg_det2.Enabled = True
End Sub

Private Sub BtnGrabar30_Click()
 On Error GoTo UpdateErr

    rs_datos1!antecedente = UCase(txt_antecedente.Text)
    rs_datos1!fecha_estimada_pla = dtp_liq.Value
    rs_datos1.Update
    'Call OptFilGral1_Click
    'rs_datos.MoveLast
    mbDataChanged = False
    
    fra_modificar2.Visible = False
    FraNavega.Enabled = True
    fraOpciones.Enabled = True
    ' FraGrabarCancelar.Visible = True
    dg_datos.Enabled = True
    dg_det1.Enabled = True
    fra_opciones_det_1.Enabled = True
    fra_opciones_det_2.Enabled = True
    dg_det2.Enabled = True
        
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnGrabar32_Click()
On Error GoTo UpdateErr

' txt_total_ganado.Text = Val(dtc_sueldo.Text) + Val(dtc_refrigerio.Text) + Val(txt_bono_ant.Text)
' txt_total_descuento.Text = Val(txt_anticipo_sb.Text) + Val(txt_anticipo_refr.Text) + Val(txt_rc_iva.Text) + Val(txt_afp1.Text) + Val(txt_afp2.Text) + Val(txt_prestamo.Text) + Val(txt_otros_descuentos.Text)
' txt_liq_pagable.Text = Val(txt_total_ganado.Text) - Val(txt_total_descuento.Text)
 If VAR_SW = "SI" Then      '1
 
    If dtc_codigo.Text <> "" And dtc_descripcion.Text <> "" Then    '2
 
     Set rs_datos2 = New ADODB.Recordset
     If rs_datos2.State = 1 Then rs_datos2.Close
     rs_datos2.Open "select * from ro_pagos_cronograma_Detalle where ges_gestion = '" & rs_datos!ges_gestion & "' AND planilla_codigo = '" & rs_datos!planilla_codigo & "' AND mes_grupo = '" & rs_datos1!mes_grupo & "' AND beneficiario_codigo = '" & dtc_codigo.Text & "'", db, adOpenKeyset, adLockOptimistic
     If rs_datos2.RecordCount = 0 Then          '3
        'rs_aux13
        If rs_aux13.State = 1 Then rs_aux13.Close
        rs_aux13.Open "select * from ro_personal_contratado where beneficiario_codigo = '" & dtc_codigo.Text & "'", db, adOpenKeyset, adLockOptimistic
     
        rs_datos2.AddNew
          rs_datos2!cargo_codigo = rs_aux13!cargo_codigo
            rs_datos2!puesto_codigo = rs_aux13!puesto_codigo
            rs_datos2!beneficiario_haber_mensual = rs_aux13!beneficiario_haber_mensual
            rs_datos2!Unidad = rs_aux6!unidad_codigo
            
        rs_datos2!ges_gestion = rs_datos!ges_gestion
        rs_datos2!planilla_codigo = rs_datos!planilla_codigo
        rs_datos2!mes_grupo = rs_datos1!mes_grupo
        rs_datos2!NUMERO_PAGO = rs_datos1!NUMERO_PAGO
        rs_datos2!beneficiario_codigo = dtc_codigo.Text
        VAR_BENEF = dtc_codigo.Text
        rs_datos2!unidad_codigo = rs_datos1!unidad_codigo_pla
        rs_datos2!tipo_moneda = "BOB"
        rs_datos2!tipo_cambio = GlTipoCambioOficial
            
        DIA_IN = Day(rs_aux13!fecha_ingreso)
        MES_IN = Month(rs_aux13!fecha_ingreso)
        ANO_IN = Year(rs_aux13!fecha_ingreso)
        DIA_HOY = Day(Now)
        If MES_IN = rs_datos2!mes_grupo And ANO_IN = rs_datos2!ges_gestion Then
        'If rs_aux6!fecha_ingreso = DateTime.Now().ToShortDateString() Then
         'Call Dias_Del_Mes(rs_aux6!fecha_ingreso)
          rs_datos2!sueldo_basico = (rs_aux13!beneficiario_haber_mensual / 30) * (30 - (DIA_IN - 1))
          rs_datos2!dias_trabajados = 30 - (DIA_IN - 1)
        Else
          rs_datos2!sueldo_basico = rs_aux13!beneficiario_haber_mensual
          rs_datos2!dias_trabajados = 30
        End If
        
        'rs_datos2!sueldo_basico = dtc_sueldo.Text
        rs_datos2!Numero_consultoriaHist = IIf(IsNull(rs_aux13!beneficiario_item), "0", rs_aux13!beneficiario_item)
        rs_datos2!monto_refrigerio = dtc_refrigerio.Text
        fecha_pla = DateSerial(rs_datos!ges_gestion, rs_datos!mes_grupo + 1, 1 - 1)
             'fecha_pla = "29/02/2016"
             'antiguedad
             If rs_aux8.State = 1 Then rs_aux8.Close
             rs_aux8.Open "select * from rc_antiguedad", db, adOpenKeyset, adLockOptimistic, adCmdText
             rs_aux8.MoveFirst
             While Not rs_aux8.EOF
             'f1 = CDate(fecha_pla) - (365 * rs_aux8!parametro_inicial)
             f1 = DateAdd("yyyy", -rs_aux8!parametro_inicial, CDate(fecha_pla))
             'f2 = CDate(fecha_pla) - (365 * rs_aux8!parametro_final)
             f2 = DateAdd("yyyy", -rs_aux8!parametro_final, CDate(fecha_pla))
             If rs_aux6!fecha_ingreso <= CDate(f1) And rs_aux6!fecha_ingreso > CDate(f2) Then
             rs_datos2!bono_antiguedad = rs_aux8!antig_valor
             rs_aux8.MoveLast
             End If
             
             rs_aux8.MoveNext
             
             Wend
 
      rs_datos2!bono_transporte = 0
        'rs_datos2!horas_extras = dtc_refrigerio.Text
        'rs_datos2!bono_transporte = dtc_refrigerio.Text
        rs_datos2!total_ganado = rs_datos2!sueldo_basico + rs_datos2!bono_antiguedad + rs_datos2!bono_transporte
        rs_datos2!provision_aguinaldo = rs_datos2!total_ganado * 0.0833
        rs_datos2!prevision_indemnizacion = rs_datos2!total_ganado * 0.0833
        
        rs_datos2!anticipo_sueldo = txt_anticipo_sb1.Text
        rs_datos2!anticipo_refrigerio = txt_anticipo_refr1.Text
        rs_datos2!prestamo = txt_prestamo1.Text
        Select Case rs_datos4!beneficiario_codigo_afp
            Case "1006803"      'AFP1 FUTURO
                rs_datos2!afp1 = rs_datos2!total_ganado * 0.1271
                rs_datos2!afp2 = "0"       'falta 987654
                VAR_NETO = rs_datos2!total_ganado - rs_datos2!afp1
            Case "987654"       'AFP2 PREVISION
                rs_datos2!afp1 = "0"       'falta 1006803
                rs_datos2!afp2 = rs_datos2!total_ganado * 0.1271
                VAR_NETO = rs_datos2!total_ganado - rs_datos2!afp2
            Case Else
                rs_datos2!afp1 = "0"
                rs_datos2!afp2 = "0"
                VAR_NETO = rs_datos2!total_ganado
        End Select
             '
'        VAR_IVA = 1805 * 4
'        If VAR_NETO > VAR_IVA Then
'            rs_datos2!rciva = rs_datos2!total_ganado * 0.13
'        Else
'            rs_datos2!rciva = "0"        'mayor a 4 SUELDOS BASICOS
'        End If
'        rs_datos2!rciva = txt_rc_iva1.Text   'mayor a 7,000.-
        
        db.Execute "UPDATE ro_controlasistencia SET ges_gestion = year(Fecha_control), Mes_control = month(Fecha_control), Dia_control= day(Fecha_control)"
         'sqlAux = "SELECT '     TOTAL MINUTOS DE RETRASO: ' + CONVERT(VARCHAR, ISNULL(SUM(DATEDIFF(MINUTE, '0:00:00', Tardanza)),0)) AS totHrs FROM ro_controlasistencia WHERE beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' "
        'rs_AsisTT.Open sqlAux, db, adOpenKeyset, adLockOptimistic, adCmdText
        'rs_AsisTT.MoveFirst
        'AdoAsistencia.Caption = CStr(rs_AsisTT!totHrs)
        If rs_aux9.State = 1 Then rs_aux9.Close
        rs_aux9.Open "select sum(AtrasoMin1) as TardanzaMes from ro_controlasistencia where beneficiario_codigo = '" & RTrim(LTrim(VAR_BENEF)) & "' AND ges_gestion = '" & RTrim(LTrim(Ado_datos1.Recordset!ges_gestion)) & "' and Mes_control = '" & RTrim(LTrim(Str(Ado_datos1.Recordset!mes_grupo))) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
         'select sum(convert(int,TardanzaCadena)) as TardanzaMes from ro_controlasistencia where beneficiario_codigo = '6960987' and Mes_control = 7
        If rs_aux14.State = 1 Then rs_aux14.Close
        mesnom = UCase(MonthName(Ado_datos1.Recordset!mes_grupo))
        rs_aux14.Open "select sum(total_minuto) as PermisoMes from ro_permisos where beneficiario_codigo = '" & RTrim(LTrim(VAR_BENEF)) & "' AND ges_gestion = '" & RTrim(LTrim(Ado_datos1.Recordset!ges_gestion)) & "' AND Mes_control = '" & mesnom & "' AND estado_codigo = 'APR' and TipoPermiso <> 'VC'", db, adOpenKeyset, adLockOptimistic, adCmdText
        If rs_aux14!PermisoMes <> "NULL" Then
            permisos = rs_aux14!PermisoMes
        Else
            permisos = "0"
        End If
        If rs_aux9!TardanzaMes <> "NULL" Then       '5
            totalminutos = rs_aux9!TardanzaMes - permisos
            If totalminutos >= 45 And totalminutos <= 60 Then
                rs_datos2!otros_dsctos = ((rs_datos2!sueldo_basico / 30) / 2)
            Else
                If totalminutos >= 61 And totalminutos <= 75 Then
                    rs_datos2!otros_dsctos = (rs_datos2!sueldo_basico / 30)
                Else
                    If totalminutos >= 76 And totalminutos <= 105 Then
                        rs_datos2!otros_dsctos = ((rs_datos2!sueldo_basico / 30) * 2)
                    Else
                        If totalminutos >= 106 Then
                            rs_datos2!otros_dsctos = ((rs_datos2!sueldo_basico / 30) * 3)
                        Else
                            rs_datos2!otros_dsctos = 0
                        End If
                    End If
                End If
            End If
        Else            '5
          If continuar = "SI" Then      '6
            sino = MsgBox("No se Cargo la asistencia del mes de " & UCase(MonthName(rs_datos1!mes_grupo)) & " de: " & dtc_descripcion.Text & vbCrLf & "¿Desea generar de todas maneras?" & vbCrLf & "NOTA: En el campo de otros descuentos se pondra 0 por defecto", vbCritical + vbYesNo, "Atención")
            If sino = vbYes Then
                rs_datos2!otros_dsctos = 0
                continuar = "SI"
            Else
                Fra_personal_Ppla.Visible = False
                FraNavega.Enabled = True
                fraOpciones.Enabled = True
                ' FraGrabarCancelar.Visible = True
                dg_datos.Enabled = True
                dg_det1.Enabled = True
                fra_opciones_det_1.Enabled = True
                fra_opciones_det_2.Enabled = True
        
                dg_det2.Enabled = True
                Call ABRIR_TABLA_DET(2)
                Exit Sub
            End If
          Else                  '6
            rs_datos2!otros_dsctos = 0
            Numero = Numero + 1
            Exit Sub
          End If                '6
        End If                  '5
        'rs_datos2!otros_dsctos = txt_otros_descuentos1.Text   'Atrasos y Faltas
        rs_datos2!r_provision_aguinaldo = "0"
        rs_datos2!r_prevision_indemnizacion = "0"
        If rs_aux15.State = 1 Then rs_aux15.Close
        rs_aux15.Open "select SUM(monto) AS totalmonto, SUM(dias) AS Totaldias from ro_memorandas where ges_gestion = '" & Ado_datos1.Recordset!ges_gestion & "' AND mes_descuento = '" & MonthName(Ado_datos1.Recordset!mes_grupo) & "' AND beneficiario_codigo =  '" & VAR_BENEF & "' AND descuento_pla = 'SI' AND estado_codigo= 'APR'  and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & ", db, adOpenKeyset, adLockOptimistic, adCmdText"
             
        If rs_aux15.RecordCount <> 0 Then
            If rs_aux15!totalmonto > 0 Then
                total = rs_datos2!otros_dsctos + IIf(IsNull(rs_aux15!totalmonto), 0, rs_aux15!totalmonto)
                rs_datos2!otros_dsctos = total
            End If
              
            If rs_aux15!Totaldias > 0 Then
                total = rs_datos2!otros_dsctos + ((rs_aux13!beneficiario_haber_mensual / 30) * rs_aux15!Totaldias)
                'total = total + rs_datos2!otros_dsctos
                  rs_datos2!otros_dsctos = total
            End If
        End If
        rs_datos2!total_dsctos = rs_datos2!anticipo_sueldo + rs_datos2!anticipo_refrigerio + rs_datos2!prestamo + rs_datos2!afp1 + rs_datos2!afp2 + txt_otros_descuentos.Text  'temporal
        rs_datos2!liquido_pagable_bs = rs_datos2!total_ganado - rs_datos2!total_dsctos
        rs_datos2!liquido_pagable_us = rs_datos2!liquido_pagable_bs / GlTipoCambioOficial
        'rs_datos2!total_dsctos = "0"
        rs_datos2!emite_factura = "N"
        
        rs_datos2!cite_conformidad = "DRRHH-"
        'rs_datos2!Numero_consultoriaHist = " "
        rs_datos2!fte_financiamientoHist = " "
        rs_datos2!estado_devengado = "REG"
         '70522788
        rs_datos2!estado_codigo = "REG"
        rs_datos2!fecha_registro = Date
        rs_datos2!usr_codigo = glusuario
        
        rs_datos2!iva_110 = txt_iva_1101.Text
        rs_datos2!fisco_a_favor = txt_fisco_a_favor1.Text
        rs_datos2!dependiente_a_favor = txt_dependiente_a_favor1.Text
        rs_datos2!mes_anterior = txt_mes_anterior1.Text
        rs_datos2!mes_anterior_mant = txt_mes_anterior_mant1.Text
        rs_datos2!saldo_util = txt_saldo_util1.Text
        rs_datos2!saldo_a_favor_depend = txt_saldo_a_favor_depend1.Text
       
        rs_datos2.Update
        Call ABRIR_TABLA_DET(2)
        'Call OptFilGral1_Click
         'rs_datos.MoveLast
        mbDataChanged = False

        Fra_personal_Ppla.Visible = False
        FraNavega.Enabled = True
        fraOpciones.Enabled = True
        ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
        dg_det1.Enabled = True
        dg_det2.Enabled = True
        fra_opciones_det_1.Enabled = True
        fra_opciones_det_2.Enabled = True
      Else          '3
        sino = MsgBox("ya existe " & dtc_descripcion.Text & " en la planilla ", vbCritical, "Atención")
      End If        '3
    End If          '2
Else                '1
  
  Dim otros_desc As Double
  otros_desc = 0
  otros_desc = Val(txt_otros_descuentos.Text)
        If rs_datos4.State = 1 Then rs_datos4.Close
        rs_datos2!ges_gestion = rs_datos!ges_gestion
        rs_datos2!planilla_codigo = rs_datos!planilla_codigo
        rs_datos2!mes_grupo = rs_datos1!mes_grupo
        rs_datos2!NUMERO_PAGO = rs_datos1!NUMERO_PAGO
        'rs_datos2!beneficiario_codigo = dtc_codigo.Text
        VAR_BENEF = txt_ci.Text
        rs_datos2!unidad_codigo = rs_datos1!unidad_codigo_pla
        rs_datos2!tipo_moneda = "BOB"
        'rs_datos2!tipo_cambio = GlTipoCambioOficial
        'rs_datos2!sueldo_basico = rs_datos4!beneficiario_haber_mensual
        'rs_datos2!monto_refrigerio = IIf(IsNull(rs_datos4!beneficiario_otro_mensual), "0", rs_datos4!beneficiario_otro_mensual)
'        If IsNull(rs_datos4!fecha_ingreso) Then
'            VAR_GES = 0
'        Else
'            VAR_GES = DateDiff("yyyy", rs_datos4!fecha_ingreso, Date)
'        End If
        'rc_antiguedad
        'SELECT antig_codigo, parametro_inicial, parametro_final, antig_porcentaje, antig_valor, estado_codigo, fecha_registro, usr_codigo From rc_antiguedad
''        'AÑO ACTUAL - AÑO(fecha_ingreso)
'        If rs_aux8.State = 1 Then rs_aux8.Close
'        rs_aux8.Open "select * from rc_antiguedad where parametro_inicial <= " & VAR_GES & " and parametro_final > " & VAR_GES & " ", db, adOpenKeyset, adLockOptimistic, adCmdText
'        If rs_aux8.RecordCount > 0 Then
'            rs_datos2!bono_antiguedad = rs_aux8!antig_valor
'        Else
'            rs_datos2!bono_antiguedad = "0"
'        End If
        'rs_datos2!horas_extras = dtc_refrigerio.Text
        'rs_datos2!bono_transporte = dtc_refrigerio.Text
        rs_datos2!total_ganado = rs_datos2!sueldo_basico + rs_datos2!bono_antiguedad + rs_datos2!bono_transporte
        rs_datos2!provision_aguinaldo = rs_datos2!total_ganado * 0.0833
        rs_datos2!prevision_indemnizacion = rs_datos2!total_ganado * 0.0833
        rs_datos2!anticipo_sueldo = IIf(IsNull(txt_anticipo_sb.Text) Or txt_anticipo_sb.Text = "", 0, txt_anticipo_sb.Text)
        rs_datos2!anticipo_refrigerio = IIf(IsNull(txt_anticipo_refr.Text) Or txt_anticipo_refr.Text = "", 0, txt_anticipo_refr.Text)
        rs_datos2!prestamo = IIf(IsNull(txt_prestamo.Text) Or txt_prestamo.Text = "", 0, txt_prestamo.Text)
'          Select Case rs_datos4!beneficiario_codigo_afp
'                Case "1006803"      'AFP1
'                    rs_datos2!afp1 = rs_datos2!total_ganado * 0.1271
'                    rs_datos2!afp2 = "0"       'falta 987654
'                    VAR_NETO = rs_datos2!total_ganado - rs_datos2!afp1
'                Case "987654"       'AFP2
'                    rs_datos2!afp1 = "0"       'falta 1006803
'                    rs_datos2!afp2 = rs_datos2!total_ganado * 0.1271
'                    VAR_NETO = rs_datos2!total_ganado - rs_datos2!afp2
'                Case Else
'                    rs_datos2!afp1 = "0"
'                    rs_datos2!afp2 = "0"
'                    VAR_NETO = rs_datos2!total_ganado
'            End Select
'             '
'            VAR_IVA = 1805 * 4
'            If VAR_NETO > VAR_IVA Then
'                rs_datos2!rciva = rs_datos2!total_ganado * 0.13
'            Else
'                rs_datos2!rciva = "0"        'mayor a 4 SUELDOS BASICOA
'            End If
         '
        ' rs_datos2!rciva = txt_rc_iva.Text   'mayor a 7,000.-
        
         db.Execute "UPDATE ro_controlasistencia SET ges_gestion = year(Fecha_control), Mes_control = month(Fecha_control), Dia_control= day(Fecha_control)"
         'sqlAux = "SELECT '     TOTAL MINUTOS DE RETRASO: ' + CONVERT(VARCHAR, ISNULL(SUM(DATEDIFF(MINUTE, '0:00:00', Tardanza)),0)) AS totHrs FROM ro_controlasistencia WHERE beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' "
        'rs_AsisTT.Open sqlAux, db, adOpenKeyset, adLockOptimistic, adCmdText
        'rs_AsisTT.MoveFirst
        'AdoAsistencia.Caption = CStr(rs_AsisTT!totHrs)
        '
            If rs_aux9.State = 1 Then rs_aux9.Close
            rs_aux9.Open "select sum(AtrasoMin1) as TardanzaMes from ro_controlasistencia where beneficiario_codigo = '" & RTrim(LTrim(VAR_BENEF)) & "' AND ges_gestion = '" & RTrim(LTrim(Ado_datos1.Recordset!ges_gestion)) & "' and Mes_control = '" & RTrim(LTrim(Str(Ado_datos1.Recordset!mes_grupo))) & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
             'select sum(convert(int,TardanzaCadena)) as TardanzaMes from ro_controlasistencia where beneficiario_codigo = '6960987' and Mes_control = 7
           
            If rs_aux14.State = 1 Then rs_aux14.Close
            mesnom = UCase(MonthName(Ado_datos1.Recordset!mes_grupo))
            rs_aux14.Open "select sum(total_minuto) as PermisoMes from ro_permisos where beneficiario_codigo = '" & RTrim(LTrim(VAR_BENEF)) & "' AND ges_gestion = '" & RTrim(LTrim(Ado_datos1.Recordset!ges_gestion)) & "' AND Mes_control = '" & mesnom & "' AND estado_codigo = 'APR' and TipoPermiso <> 'VC'", db, adOpenKeyset, adLockOptimistic, adCmdText
           
           
'           If rs_aux14.State = 1 Then rs_aux14.Close
'              rs_aux14.Open "select sum(total_minuto) as PermisoMes from ro_permisos where beneficiario_codigo = '" & RTrim(LTrim(VAR_BENEF)) & "' AND ges_gestion = '" & RTrim(LTrim(Ado_datos1.Recordset!ges_gestion)) & "' and Mes_control = '" & MonthName(Ado_datos1.Recordset!mes_grupo) & "' AND estado_codigo = 'APR'", db, adOpenKeyset, adLockOptimistic, adCmdText
'              'Dim permisos As String
            
           If rs_aux14!PermisoMes <> "NULL" Then
                permisos = rs_aux14!PermisoMes
            Else
                permisos = "0"
            End If
            If rs_aux9!TardanzaMes <> "NULL" Then
             totalminutos = rs_aux9!TardanzaMes - permisos
                If totalminutos >= 45 And totalminutos <= 60 Then
                    rs_datos2!otros_dsctos = ((rs_datos2!sueldo_basico / 30) / 2)
                Else
                    If totalminutos >= 61 And totalminutos <= 75 Then
                        rs_datos2!otros_dsctos = (rs_datos2!sueldo_basico / 30)
                    Else
                        If totalminutos >= 76 And totalminutos <= 105 Then
                            rs_datos2!otros_dsctos = ((rs_datos2!sueldo_basico / 30) * 2)
                        Else
                            If totalminutos >= 106 Then
                                rs_datos2!otros_dsctos = ((rs_datos2!sueldo_basico / 30) * 3)
                            Else
                                rs_datos2!otros_dsctos = 0
                            End If
                        End If
                    End If
                End If
            Else
           If continuar = "SI" Then
          sino = MsgBox("No se Cargo la asistencia del mes de " & UCase(MonthName(rs_datos1!mes_grupo)) & " de algunas personas " & vbCrLf & "¿Desea generar de todas maneras?" & vbCrLf & "NOTA: En el campo de otros descuentos se pondra 0 por defecto", vbCritical + vbYesNo, "Atención")
            If sino = vbYes Then
             rs_datos2!otros_dsctos = 0
             continuar = "SI"
            Else
            
           
            Fra_personal_Ppla.Visible = False
            FraNavega.Enabled = True
            fraOpciones.Enabled = True
            ' FraGrabarCancelar.Visible = True
            dg_datos.Enabled = True
            dg_det1.Enabled = True
            fra_opciones_det_1.Enabled = True
            fra_opciones_det_2.Enabled = True

            dg_det2.Enabled = True
            Call ABRIR_TABLA_DET(2)
            Exit Sub
            End If
            
            
            Else
           rs_datos2!otros_dsctos = 0
           Numero = Numero + 1
           End If
         End If
         'rs_datos2!otros_dsctos = txt_otros_descuentos1.Text   'Atrasos y Faltas
         
         rs_datos2!r_provision_aguinaldo = "0"
         rs_datos2!r_prevision_indemnizacion = "0"
             If rs_aux15.State = 1 Then rs_aux15.Close
              rs_aux15.Open "select SUM(monto) AS totalmonto, SUM(dias) AS Totaldias from ro_memorandas where ges_gestion = '" & Ado_datos1.Recordset!ges_gestion & "' AND mes_descuento = '" & MonthName(Ado_datos1.Recordset!mes_grupo) & "' AND beneficiario_codigo = '" & VAR_BENEF & "' AND descuento_pla = 'SI' AND estado_codigo = 'APR'  and codigo_empresa = " & Ado_datos.Recordset!codigo_empresa & "", db, adOpenKeyset, adLockOptimistic, adCmdText
             
         If rs_aux15.RecordCount <> 0 Then
              If rs_aux15!totalmonto <> "NULL" Then
              total = rs_datos2!otros_dsctos + rs_aux15!totalmonto
              rs_datos2!otros_dsctos = total
              End If
              
              If rs_aux15!Totaldias <> "NULL" Then
              total = rs_datos2!otros_dsctos + ((rs_aux6!beneficiario_haber_mensual / 30) * rs_aux15!Totaldias)
             ' total = total + rs_datos2!otros_dsctos
              rs_datos2!otros_dsctos = total
              End If
      
              'rs_datos2!otros_dsctos = total
         End If
         rs_datos2!otros_dsctos = otros_desc
         rs_datos2!total_dsctos = rs_datos2!anticipo_sueldo + rs_datos2!anticipo_refrigerio + rs_datos2!prestamo + rs_datos2!afp1 + rs_datos2!afp2 + rs_datos2!otros_dsctos   'temporal
         
         rs_datos2!liquido_pagable_bs = rs_datos2!total_ganado - rs_datos2!total_dsctos
         rs_datos2!liquido_pagable_us = rs_datos2!liquido_pagable_bs / GlTipoCambioOficial
         'rs_datos2!total_dsctos = "0"
         rs_datos2!emite_factura = "N"
         
         rs_datos2!cite_conformidad = " "
         'rs_datos2!Numero_consultoriaHist = " "
         rs_datos2!fte_financiamientoHist = " "
         rs_datos2!estado_devengado = "REG"
         '70522788
        rs_datos2!estado_codigo = "REG"
        rs_datos2!fecha_registro = Date
        rs_datos2!usr_codigo = glusuario
        
     rs_datos2!iva_110 = IIf(IsNull(txt_iva_110.Text) Or txt_iva_110.Text = "", 0, txt_iva_110.Text)
     rs_datos2!fisco_a_favor = IIf(IsNull(txt_fisco_a_favor.Text) Or txt_fisco_a_favor.Text = "", 0, txt_fisco_a_favor.Text)
     rs_datos2!dependiente_a_favor = IIf(IsNull(txt_dependiente_a_favor.Text) Or txt_dependiente_a_favor.Text = "", 0, txt_dependiente_a_favor.Text)
     rs_datos2!mes_anterior = IIf(IsNull(txt_mes_anterior.Text) Or txt_mes_anterior.Text = "", 0, txt_mes_anterior.Text)
     rs_datos2!mes_anterior_mant = IIf(IsNull(txt_mes_anterior_mant.Text) Or txt_mes_anterior_mant.Text = "", 0, txt_mes_anterior_mant.Text)
     rs_datos2!saldo_util = IIf(IsNull(txt_saldo_util.Text) Or txt_saldo_util.Text = "", 0, txt_saldo_util.Text)
    rs_datos2!saldo_a_favor_depend = IIf(IsNull(txt_saldo_a_favor_depend.Text) Or txt_saldo_a_favor_depend.Text = "", 0, txt_saldo_a_favor_depend.Text)
        
   
        rs_datos2.Update
     Call ABRIR_TABLA_DET(2)
    End If
           Fra_personal_Ppla.Visible = False
            FraNavega.Enabled = True
            fraOpciones.Enabled = True
            ' FraGrabarCancelar.Visible = True
            dg_datos.Enabled = True
            dg_det1.Enabled = True
            fra_opciones_det_1.Enabled = True
            fra_opciones_det_2.Enabled = True

            dg_det2.Enabled = True
         
fra_opciones_det_2.Enabled = True
fra_opciones_det_1.Enabled = True
        VAR_SW = ""
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub Picture33_Click()
Fra_personal_Ppla.Visible = False
FraNavega.Enabled = True
       fraOpciones.Enabled = True
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
         dg_det1.Enabled = True
         fra_opciones_det_1.Enabled = True
fra_opciones_det_2.Enabled = True

        dg_det2.Enabled = True
        
        
        
         FraNavega.Enabled = True
       fraOpciones.Enabled = True
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
         dg_det1.Enabled = True
         fra_opciones_det_1.Enabled = True
fra_opciones_det_2.Enabled = True
End Sub

Private Sub Picture36_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
      Call tipo_rep

    Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If

End Sub

Private Sub Picture37_Click()
fra_imprimir.Visible = False
 FraNavega.Enabled = True
       fraOpciones.Enabled = True
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
         dg_det1.Enabled = True
         fra_opciones_det_1.Enabled = True
fra_opciones_det_2.Enabled = True

        dg_det2.Enabled = True
End Sub



Private Sub Picture39_Click()
 fra_imprimir.Visible = True
fra_reportes.Visible = False
End Sub

Private Sub Picture4_Click()
'Call ABRIR_TABLAS_AUX(1)
Call ABRIR_TABLAS_AUX(2)

  fra_imprimir.Visible = True
    FraNavega.Enabled = False
       fraOpciones.Enabled = False
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
         dg_det1.Enabled = False
         fra_opciones_det_1.Enabled = False
          fra_opciones_det_2.Enabled = False
        dg_det2.Enabled = False
        
        dtc_rep_cod.Text = ""
         dtc_rep_det.Text = ""
         cbo_mes_rep.Text = ""
        cmb_gestion_rep.Text = Year(Date)
        Option1.Value = False
  
End Sub

Private Sub BtnGrabar40_Click()
    If dt_unidad_cod.Text = "" Then
        sino = MsgBox("Llene todos los datos", vbCritical, "Atención")
    Else
        If rs_aux10.State = 1 Then rs_aux10.Close
        rs_aux10.Open "select * from ro_pagos_cronograma where ges_gestion = '" & rs_datos!ges_gestion & "' AND planilla_codigo = '" & rs_datos!planilla_codigo & "' AND mes_grupo = " & rs_datos!mes_grupo & "AND unidad_codigo_pla = '" & dt_unidad_cod.Text & "'", db, adOpenKeyset, adLockOptimistic
        'If rs_aux10.RecordCount = 0 Then
        db.Execute "Insert INTO ro_pagos_cronograma ( ges_gestion , planilla_codigo, mes_grupo, numero_pago, unidad_codigo_pla, concepto, antecedente, tipo_moneda, monto_us, monto_bs, fecha_estimada_pla, ges_gestion_ANT, codigo_unidad_ANT, planilla_codigo_ANT, numero_pago_ANT, estado_codigo, usr_codigo, fecha_registro) Values ('" & rs_datos!ges_gestion & "', '" & rs_datos!planilla_codigo & "', " & rs_datos!mes_grupo & ", " & rs_aux10.RecordCount + 1 & ", '" & dt_unidad_cod.Text & "', '" & dt_unidad_det.Text & "', '" & rs_datos!descripcion_grupo & " - " & dt_unidad_det.Text & "','BOB','0','0', '" & Date & "','" & rs_datos!ges_gestion & "', '" & rs_datos!planilla_codigo & "', " & rs_datos!mes_grupo & ", '1','REG', '" & glusuario & "',  '" & Date & "')"
        Call ABRIR_TABLA_DET(1)
        fra_sub_grupo_unidad.Visible = False
        FraNavega.Enabled = True
        fraOpciones.Enabled = True
        ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
        dg_det1.Enabled = True
        fra_opciones_det_1.Enabled = True
        fra_opciones_det_2.Enabled = True
        dg_det2.Enabled = True
        ' Else
        '     sino = MsgBox("ya se GENERO anteriormente el Sub Grupo Por Unidad  con estos datos: " & vbCrLf & "Nro. pago: " & cbo_numero_pago.Text & vbCrLf & "Mes:" & UCase(MonthName(rs_datos!mes_grupo)) & vbCrLf & "Unidad: " & dt_unidad_cod.Text & " - " & dt_unidad_det.Text, vbCritical, "Atención")
        ' End If
    End If
End Sub

Private Sub Picture42_Click()
fra_nueva_pla.Visible = False
  FraNavega.Enabled = True
       fraOpciones.Enabled = True
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
         dg_det1.Enabled = True
         fra_opciones_det_1.Enabled = True
fra_opciones_det_2.Enabled = True
End Sub

Private Sub BtnGrabar43_Click()
If txt_concepto_pla.Text <> "" Or dtc_pla_cod.Text <> "" Or cbo_mes_pla.Text <> "" Then
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "select * from ro_pagos_grupos where ges_gestion = '" & cbo_gestion_pla.Text & "' AND planilla_codigo = '" & dtc_pla_cod.Text & "' AND descripcion_grupo = '" & dtc_pla_det.Text & " - " & UCase(txt_concepto_pla.Text) & "'", db, adOpenKeyset, adLockOptimistic
    If rs_aux1.RecordCount = 0 Then
        If rs_aux12.State = 1 Then rs_aux12.Close
        rs_aux12.Open "select * from ro_pagos_grupos where ges_gestion = '" & cbo_gestion_pla.Text & "' AND planilla_codigo = '" & dtc_pla_cod.Text & "'", db, adOpenKeyset, adLockOptimistic
        If rs_aux12.RecordCount > 0 Then
            db.Execute "Insert INTO ro_pagos_grupos (ges_gestion, planilla_codigo, mes_grupo, descripcion_grupo, unidad_codigo, depto_codigo, clasif_codigo,doc_codigo,solicitud_tipo, estado_codigo, usr_codigo, fecha_registro) Values ('" & cbo_gestion_pla.Text & "', '" & dtc_pla_cod.Text & "', " & rs_aux12.RecordCount + 1 & ", '" & dtc_pla_det.Text & " - " & UCase(txt_concepto_pla.Text) & "', 'RRHH', '" & rs_aux11!depto_codigo & "', 'ADM','R-121', '11','REG', '" & glusuario & "',  '" & Date & "')"
            Call OptFilGral1_Click
            fra_nueva_pla.Visible = False
            FraNavega.Enabled = True
            fraOpciones.Enabled = True
            ' FraGrabarCancelar.Visible = True
            dg_datos.Enabled = True
            dg_det1.Enabled = True
            fra_opciones_det_1.Enabled = True
            fra_opciones_det_2.Enabled = True
        Else
            sino = MsgBox("Primero haga que el sistema genere las planillas para la Gestión " & cbo_gestion_pla.Text, vbInformation, "Atención")
            fra_nueva_pla.Visible = False
            FraNavega.Enabled = True
            fraOpciones.Enabled = True
            ' FraGrabarCancelar.Visible = True
            dg_datos.Enabled = True
            dg_det1.Enabled = True
            fra_opciones_det_1.Enabled = True
            fra_opciones_det_2.Enabled = True
        End If
    Else
        sino = MsgBox("Este registro ya existe en la Planilla ", vbCritical, "Atención")
    End If
Else
    sino = MsgBox("Llene todos los datos ", vbInformation, "Atención")
End If
End Sub

Private Sub BtnGrabar45_Click()

    Call generar_rc_iva
    fra_ufv.Visible = False
    FraNavega.Enabled = True
    fraOpciones.Enabled = True
    ' FraGrabarCancelar.Visible = True
    dg_datos.Enabled = True
    dg_det1.Enabled = True
    fra_opciones_det_1.Enabled = True
    fra_opciones_det_2.Enabled = True
    dg_det2.Enabled = True
    FraNavega.Enabled = True
    fraOpciones.Enabled = True
    ' FraGrabarCancelar.Visible = True
    dg_datos.Enabled = True
    dg_det1.Enabled = True
    fra_opciones_det_1.Enabled = True
    fra_opciones_det_2.Enabled = True
End Sub

Private Sub Picture46_Click()
    fra_ufv.Visible = False
    FraNavega.Enabled = True
    fraOpciones.Enabled = True
    ' FraGrabarCancelar.Visible = True
    dg_datos.Enabled = True
    dg_det1.Enabled = True
    fra_opciones_det_1.Enabled = True
    fra_opciones_det_2.Enabled = True
    dg_det2.Enabled = True
    FraNavega.Enabled = True
    fraOpciones.Enabled = True
    ' FraGrabarCancelar.Visible = True
    dg_datos.Enabled = True
    dg_det1.Enabled = True
    fra_opciones_det_1.Enabled = True
    fra_opciones_det_2.Enabled = True
End Sub

Private Sub Picture6_Click()
'If Ado_datos.Recordset.RecordCount > 0 Then
'If Ado_datos1.Recordset.RecordCount > 0 Then
'On Error GoTo UpdateErr
'   If rs_datos!estado_codigo = "REG" Then
'      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'         rs_datos1!estado_codigo = "APR"
'         rs_datos1!Fecha_Registro = Date
'         rs_datos1!usr_codigo = glusuario
'         rs_datos1.UpdateBatch adAffectAll
'      End If
'   Else
'       MsgBox "No se puede APROBAR un registro Anulado (ERR) o Aprobado (APR) anteriormente ...", vbExclamation, "Validación de Registro"
'   End If
'   Exit Sub
'UpdateErr:
'  MsgBox Err.Description
'Else
'      MsgBox "No existen registros", vbExclamation, "Error"
'End If
'Else
'      MsgBox "No existen registros", vbExclamation, "Error"
'End If
If Ado_datos.Recordset.RecordCount > 0 Then
If Ado_datos1.Recordset.RecordCount > 0 Then
On Error GoTo UpdateErr
   If Ado_datos1.Recordset!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        VAR_BENEF1 = "1006803"      'AFP FUTURO     'rs_datos!beneficiario_codigo
        VAR_BENEF2 = "987654"       'AFP PREVISION
        VAR_BENEF3 = "1018595023"   'CAJA  PETROLERA DE SALUD
        VAR_BENEF4 = "29001"        'UNIDAD EJECUTORA DE TITULACION (EX FONVIS)
        VAR_CITE = Ado_datos1.Recordset!ges_gestion + "-" + Ado_datos1.Recordset!unidad_codigo_pla
        VAR_GLOSA = Trim(Ado_datos1.Recordset!antecedente) + " - Gestion: " + Trim(Ado_datos1.Recordset!ges_gestion)
        VAR_DOL2 = Round(IIf(IsNull(Ado_datos1.Recordset!tot_liquido_pagable_us), 0, Ado_datos1.Recordset!tot_liquido_pagable_us), 2)
        VAR_BS2 = Round(IIf(IsNull(Ado_datos1.Recordset!tot_liquido_pagable_bs), 0, Ado_datos1.Recordset!tot_liquido_pagable_bs), 2)
        VAR_AFP1_BS = Round(IIf(IsNull(Ado_datos1.Recordset!tot_afp1), 0, Ado_datos1.Recordset!tot_afp1), 2)
        VAR_AFP2_BS = Round(IIf(IsNull(Ado_datos1.Recordset!tot_afp2), 0, Ado_datos1.Recordset!tot_afp2), 2)
        VAR_DSCTO_BS = Round(IIf(IsNull(Ado_datos1.Recordset!total_dsctos), 0, (Ado_datos1.Recordset!total_dsctos - VAR_AFP1_BS - VAR_AFP2_BS)), 2)
        
        VAR_AFP1_BS2 = Round(IIf(IsNull(Ado_datos1.Recordset!tot_total_ganado), 0, Ado_datos1.Recordset!tot_total_ganado) * 0.471, 2)
        VAR_AFP2_BS2 = Round(IIf(IsNull(Ado_datos1.Recordset!tot_total_ganado), 0, Ado_datos1.Recordset!tot_total_ganado) * 0.471, 2)
        VAR_SS_BS = Round(IIf(IsNull(Ado_datos1.Recordset!tot_total_ganado), 0, Ado_datos1.Recordset!tot_total_ganado) * 0.1, 2)
        VAR_FV_BS = Round(IIf(IsNull(Ado_datos1.Recordset!tot_total_ganado), 0, Ado_datos1.Recordset!tot_total_ganado) * 0.02, 2)
        VAR_INDE_BS = Round(IIf(IsNull(Ado_datos1.Recordset!tot_total_ganado), 0, Ado_datos1.Recordset!tot_total_ganado) * 0.0833, 2)
        VAR_AGUI_BS = Round(IIf(IsNull(Ado_datos1.Recordset!tot_total_ganado), 0, Ado_datos1.Recordset!tot_total_ganado) * 0.0833, 2)
        VAR_CTA = "NN"              'IIf(AdoAux.Recordset!Cta_Codigo = "", "NN", AdoAux.Recordset!Cta_Codigo)
        VAR_PROY2 = "20101-9"       'rs_datos!edif_codigo
        VAR_COD4 = "DRRHH"          'rs_datos!unidad_codigo
        VAR_TIPOV = "E"             'EFECTIVO       'rs_datos!venta_tipo
        
        VAR_FFAC = Ado_datos1.Recordset!fecha_estimada_pla   '"17/11/2016"
        'End If
        NroFactura = "0"            'rs_datos1!cobranza_nro_factura
        NRO_COBR = "0"              'Me.AdoAux.Recordset!cobranza_codigo
        var_literal = Literal(CStr(VAR_BS2)) + " BOLIVIANOS"     'IIf(IsNull(AdoAux.Recordset!Literal), "-", AdoAux.Recordset!Literal)
        VAR_MONEDA = "BOB"   'AdoAux.Recordset!tipo_moneda
        VAR_CODTIPO = "DAP"
        VAR_DOC = "R-111"
        VAR_ETAPA = "FIN-02-04"
        VAR_TCOMP = VAR_CODTIPO      '"DE"        '"RECAUDADO (INGRESOS)"
        VAR_ANIO = Year(VAR_FFAC)
        gestion0 = Year(VAR_FFAC)
        VAR_MES = UCase(MonthName(Month(VAR_FFAC)))
        VAR_PLA = Ado_datos1.Recordset!planilla_codigo
        VAR_SPLA = Ado_datos1.Recordset!unidad_codigo_pla
        VAR_SW = 1
         'Call Contabiliza_gasto
         db.Execute "UPDATE ro_pagos_cronograma SET estado_codigo = 'APR' WHERE ges_gestion='" & VAR_ANIO & "' AND planilla_codigo= '" & Ado_datos1.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos1.Recordset!mes_grupo & " AND numero_pago=" & Ado_datos1.Recordset!NUMERO_PAGO & " AND unidad_codigo_pla='" & Ado_datos1.Recordset!unidad_codigo_pla & "' "
         MsgBox "Se APROBO satisfactoriamente el Registro ...", vbInformation, "Información de Registro"
         'Ado_datos1.Recordset!estado_codigo = "APR"
         'Ado_datos1.Recordset!fecha_registro = Date
         'Ado_datos1.Recordset!usr_codigo = glusuario
         'Ado_datos1.Recordset.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado (ANL) o Aprobado (APR) anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
Else
      MsgBox "No existen registros", vbExclamation, "Error"
End If
Else
      MsgBox "No existen registros", vbExclamation, "Error"
End If

End Sub

Private Sub Picture7_Click()

If Ado_datos.Recordset.RecordCount > 0 Then
If Ado_datos1.Recordset.RecordCount > 0 Then
On Error GoTo UpdateErr
If rs_datos!estado_codigo = "REG" Then
   If ExisteReg("ges_gestion = " & Ado_datos1.Recordset!ges_gestion & " AND planilla_codigo = '" & Ado_datos1.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos1.Recordset!mes_grupo & " AND numero_pago = " & Ado_datos1.Recordset!NUMERO_PAGO & "AND unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo_pla & "'", "ro_pagos_cronograma_Detalle") Then
      sino = MsgBox("No se puede ELIMINAR porque el Registro ya fue utilizado. Desea marcar como ERRADO ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos1!estado_codigo = "ERR"
         rs_datos1!fecha_registro = Date
         rs_datos1!usr_codigo = glusuario
         rs_datos1.UpdateBatch adAffectAll
      End If
   Else
      sino = MsgBox("Está Seguro de ELIMINAR fisicamente el Registro ? ", vbYesNo + vbQuestion, "Atención") 'unidad_codigo_pla
      If sino = vbYes Then
         db.Execute "DELETE ro_pagos_cronograma where ges_gestion = " & Ado_datos1.Recordset!ges_gestion & " AND planilla_codigo = '" & Ado_datos1.Recordset!planilla_codigo & "' AND mes_grupo = " & Ado_datos1.Recordset!mes_grupo & " AND numero_pago = " & Ado_datos1.Recordset!NUMERO_PAGO & " AND unidad_codigo_pla = '" & Ado_datos1.Recordset!unidad_codigo_pla & "'"
      End If
   End If
   Call ABRIR_TABLA_DET(1)
   
Else    'APR
      MsgBox "La planilla ya fue APROBADA no se puede hacer ningun cambio", vbExclamation, "Error"
End If 'REG
   Exit Sub

UpdateErr:
  MsgBox Err.Description
  Else
      MsgBox "No existen registros", vbExclamation, "Error"
End If
Else
      MsgBox "No existen registros", vbExclamation, "Error"
End If
End Sub

Private Sub Picture8_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
If Ado_datos1.Recordset.RecordCount > 0 Then
 On Error GoTo EditErr
  
 'lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
    fra_modificar2.Visible = True
          FraNavega.Enabled = False
       fraOpciones.Enabled = False
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
         dg_det1.Enabled = False
         fra_opciones_det_1.Enabled = False
          fra_opciones_det_2.Enabled = False
        dg_det2.Enabled = False
        swnuevo = 2
    '    BtnVer.Visible = True
    Else    'APR
      MsgBox "La planilla ya fue APROBADA no se puede hacer ningun cambio", vbExclamation, "Error"
End If 'REG
  Exit Sub

EditErr:
  MsgBox Err.Description
  Else
      MsgBox "No existen registros", vbExclamation, "Error"
End If
Else
      MsgBox "No existen registros", vbExclamation, "Error"
End If
End Sub

Private Sub BtnNuevo9_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Call ABRIR_TABLAS_AUX(4)
    On Error GoTo UpdateErr
    If rs_datos!estado_codigo = "REG" Then
        cbo_numero_pago.Text = ""
        dt_unidad_cod.Text = ""
        dt_unidad_det.Text = ""
        sino = MsgBox("¿Quiere generar autamaticamente la(s) SUB PLANILLA(S)?", vbYesNo + vbQuestion, "Atención")
        If sino = vbYes Then
            'If rs_aux3.State = 1 Then rs_aux3.Close
            'rs_aux3.Open "select * from ro_pagos_cronograma where ges_gestion = '" & rs_datos!ges_gestion & "' AND planilla_codigo = '" & rs_datos!planilla_codigo & "' AND mes_grupo = " & rs_datos!mes_grupo & "", db, adOpenKeyset, adLockOptimistic
            'If rs_aux3.RecordCount = 0 Then
            'Dim contador As Integer
            'contador = 0
          If rs_aux4.State = 1 Then rs_aux4.Close
          rs_aux4.Open "select * from rc_planilla_sub_grupo where estado_codigo = 'APR' AND planilla_codigo = '" & rs_datos!planilla_codigo & "'", db, adOpenKeyset, adLockOptimistic
          rs_aux4.MoveFirst
          ProgressBar2.Visible = True
          With ProgressBar2
            .Max = rs_aux4.RecordCount
            .Min = 0
            .Value = 0
          End With
          While Not rs_aux4.EOF
            ProgressBar2.Value = ProgressBar2.Value + 1
            'While (contador <> cmb_nro_planillas.Text)
            'contador = contador + 1
            'If contador > 12 Then
            'db.Execute "Insert INTO ro_pagos_grupos (ges_gestion, planilla_codigo, mes_grupo, descripcion_grupo, unidad_codigo, depto_codigo, clasif_codigo,doc_codigo,solicitud_tipo, estado_codigo, usr_codigo, fecha_registro) Values ('" & cmb_gestion.Text & "', '" & rs_aux2!planilla_codigo & "', " & contador & ", '" & rs_aux2!planilla_descripcion & " - AGUINALDO " & (contador - 12) & "', 'RRHH', '" & rs_aux2!depto_codigo & "', 'ADM','R-121', '11','REG', '" & glusuario & "',  '" & Date & "')"
            'Else
            If rs_aux3.State = 1 Then rs_aux3.Close
            rs_aux3.Open "select * from ro_pagos_cronograma where ges_gestion = '" & rs_datos!ges_gestion & "' AND planilla_codigo = '" & rs_datos!planilla_codigo & "' AND mes_grupo = " & rs_datos!mes_grupo & " AND numero_pago = 1 AND unidad_codigo_pla = '" & rs_aux4!unidad_codigo_pla & "'", db, adOpenKeyset, adLockOptimistic
            If rs_aux3.RecordCount = 0 Then
                db.Execute "Insert INTO ro_pagos_cronograma ( ges_gestion, planilla_codigo, mes_grupo, numero_pago, unidad_codigo_pla, concepto, antecedente, tipo_moneda, monto_us, monto_bs, fecha_estimada_pla, ges_gestion_ANT, codigo_unidad_ANT, planilla_codigo_ANT, numero_pago_ANT, estado_codigo, usr_codigo, fecha_registro) Values ('" & rs_datos!ges_gestion & "', '" & rs_datos!planilla_codigo & "', " & rs_datos!mes_grupo & ", '1', '" & rs_aux4!unidad_codigo_pla & "', '" & rs_aux4!unidad_descripcion_pla & "', '" & rs_datos!descripcion_grupo & " - " & rs_aux4!unidad_descripcion_pla & "','BOB','0','0', '" & Date & "','" & rs_datos!ges_gestion & "', '" & rs_datos!planilla_codigo & "', " & rs_datos!mes_grupo & ", '1','REG', '" & glusuario & "',  '" & Date & "')"
            End If
            rs_aux4.MoveNext
          Wend
      
       'contador = 0
       'Wend
'       rs_datos1.Updat
     Call ABRIR_TABLA_DET(1)
      sino = MsgBox("¿Quiere generar autamaticamente los datos del PERSONAL de todas Las SUB-PLANILLAS?", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
      If rs_datos!mes_grupo > 12 Then
      Call generar_aguinaldo
      Else
      Call generar_personas
      End If
     End If
      
      
      
      
     If NUM_PERS > 0 Then
     sino = MsgBox("Se generaron correctamente las planillas", vbInformation, "Atención")
     Else
     sino = MsgBox("No se agregó ninguna persona a la planilla", vbCritical, "Atención")
     End If
   Call ABRIR_TABLAS_AUX(5)
    ProgressBar2.Visible = False
     'Call OptFilGral1_Click
        fra_generar.Visible = False
        FraNavega.Enabled = True
        fraOpciones.Enabled = True
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = True
        dg_det1.Enabled = True
        dg_det2.Enabled = True
'     Else
'     sino = MsgBox("ya se GENERO anteriormente el Sub Grupo Por Unidad Para el mes de " & UCase(MonthName(rs_datos!mes_grupo)), vbCritical, "Atención")
'
'     End If
    Else
    fra_sub_grupo_unidad.Visible = True
             FraNavega.Enabled = False
       fraOpciones.Enabled = False
       ' FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        dg_det1.Enabled = False
        fra_opciones_det_1.Enabled = False
        fra_opciones_det_2.Enabled = False
        dg_det2.Enabled = False
     End If
     Else    'APR
      MsgBox "La planilla ya fue APROBADA no se puede hacer ningun cambio", vbExclamation, "Error"
    End If 'REG
End If
'
'  Else
'sino = MsgBox("Primero tiene que crear la Planilla", vbCritical, "Atención")
'End If
Exit Sub
UpdateErr:
  MsgBox Err.Description
  
  
End Sub

Private Sub tipo_rep()
Dim iResult As Integer
CR01.Reset
CR01.WindowShowPrintSetupBtn = True
CR01.WindowShowSearchBtn = True
CR01.WindowShowRefreshBtn = True

If cbo_mes_rep.Text = "" Or dtc_rep_cod.Text = "" Or dtc_rep_det.Text = "" Then
sino = MsgBox("Llene todos los datos para el REPORTE por favor", vbCritical, "Atención")
Else
If rb_prevision.Value = False And rb_futuro.Value = False And rb_pla_ministerio.Value = False Then
sino = MsgBox("Elija el reporte que desea imprimir", vbCritical, "Atención")
Else
  
  If rb_prevision.Value = True Then
    CR01.ReportFileName = App.Path & "\REPORTES\RRHH\rr_fondo_pensiones.rpt"
    CR01.StoredProcParam(0) = cmb_gestion_rep.Text
    CR01.StoredProcParam(1) = dtc_rep_cod.Text
    CR01.StoredProcParam(2) = txt_mes.Text
    CR01.StoredProcParam(3) = "1"
  End If
    
  If rb_futuro.Value = True Then
    CR01.ReportFileName = App.Path & "\REPORTES\RRHH\rr_fondo_pensiones.rpt"
    CR01.StoredProcParam(0) = cmb_gestion_rep.Text
    CR01.StoredProcParam(1) = dtc_rep_cod.Text
    CR01.StoredProcParam(2) = txt_mes.Text
    CR01.StoredProcParam(3) = "2"
  End If
 
 If rb_pla_ministerio.Value = True Then
    
    CR01.ReportFileName = App.Path & "\REPORTES\RRHH\rr_planilla_ministerio.rpt"
    CR01.StoredProcParam(0) = cmb_gestion_rep.Text
    CR01.StoredProcParam(1) = dtc_rep_cod.Text
    CR01.StoredProcParam(2) = txt_mes.Text
    
 End If
CR01.WindowState = crptMaximized
iResult = CR01.PrintReport
  If iResult <> 0 Then
      MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
'
  
  'fra_imprimir.Visible = False
  FraNavega.Enabled = True
  fraOpciones.Enabled = True
 'FraGrabarCancelar.Visible = True
  dg_datos.Enabled = True
  dg_det1.Enabled = True
  dg_det2.Enabled = True
  End If
  End If
 

End Sub

Private Sub txt_total_ganado_Click()
'txt_total_ganado.Text = Val(dtc_sueldo.Text) + Val(dtc_refrigerio.Text)
End Sub


Private Sub Contabiliza_gasto()
    'CONTABILIZA PLANILLA
'    'Call graba_proyecto
    'If VAR_SW = 1 Then
    '    Call graba_gasto
    'End If
'    'If VAR_SW = 1 Then
'        Set rstdestino = New ADODB.Recordset
'        If VAR_TIPOV = "L" Then
'            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEY' ", db, adOpenDynamic, adLockOptimistic
'        Else
'            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
'        End If
'        If rstdestino.RecordCount > 0 Then
'            VAR_CODANT = rstdestino!ingreso_codigo
'            VAR_ORG = rstdestino!org_codigo
'            VAR_FTE = rstdestino!fte_codigo
'
'            'Modificar con CASE WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW MAY-2015
'            If VAR_SW = 1 Then
'                VAR_TSOL = VAR_TIPOS
'            Else
'                VAR_TSOL = rstdestino!solicitud_tipo
'                VAR_TIPOS = rstdestino!solicitud_tipo
'                VAR_PARTIDA = rstdestino!rubro_codigo
'            End If
'        Else
'
'        End If
'    'End If
'  '===== Proceso para generar Asientos Contables Automáticos "DEI" y "REC"
'  'sino = MsgBox("¿Está seguro de aprobar el Registro?", vbYesNo + vbQuestion, "CONFIRMAR...")
'  'If sino = vbYes Then

    ' INI CORRECCION 18-JUN-2017
    
    VAR_ORG = "111"         'rstdestino!org_codigo
    VAR_FTE = "10"          'rstdestino!fte_codigo
    VAR_TSOL = "11"         'rstdestino!solicitud_tipo
    VAR_TIPOS = "11"        'rstdestino!solicitud_tipo
    VAR_PARTIDA = "11700"   'rstdestino!par_codigo
    ' de 11100   a   11940
    VAR_CODTIPO = "DAP"

    Dim i As Integer
    Dim j As Integer
    Dim v_Tipo_Comp(1, 2)

    fte_codigo1 = VAR_FTE
    '**** INI VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    Select Case VAR_CODTIPO
        Case "DAP"
            'falta DAC
            'rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            rstdestino.Open "select * from fc_relacionador_gastos where Codigo_Tipo = 'DAP' and par_codigo_I <= " & (VAR_PARTIDA) & " and par_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque la PARTIDA no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
        Case "PAC"
            If VAR_MONEDA = "BOB" Then
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and SubCta_Deb1 = '01' ", db, adOpenKeyset, adLockReadOnly
            'rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            Else
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "  and SubCta_Deb1 = '02' ", db, adOpenKeyset, adLockReadOnly
            End If
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
                        
            If VAR_JQ = "" Then
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
                'rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
                If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
                  If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
                    MsgBox "El monto que está intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
                    'JQA FEB-2016
                    'Exit Sub
                  End If
                End If
                If rs_aux1.State = 1 Then rs_aux1.Close
            End If
        Case "CYD"
            If VAR_VTIPO = "L" Then     'Importación Directa
                If rstdestino.State = 1 Then rstdestino.Close
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = '" & VAR_CODTIPO & "' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " AND correlativo <> '6' ", db, adOpenKeyset, adLockReadOnly
                If rstdestino.RecordCount > 0 Then
                    j = rstdestino.RecordCount
                Else
                  MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
                  Exit Sub
                End If
            End If
            If VAR_VTIPO = "V" Then     'Facturacion Local
                If rstdestino.State = 1 Then rstdestino.Close
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = '" & VAR_CODTIPO & "' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "  ", db, adOpenKeyset, adLockReadOnly
                If rstdestino.RecordCount > 0 Then
                    j = rstdestino.RecordCount
                Else
                  MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
                  Exit Sub
                End If
            End If
            If VAR_VTIPO <> "L" And VAR_VTIPO <> "V" Then       'Mant, Rep, Inst, etc.
                If rstdestino.State = 1 Then rstdestino.Close
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = '" & VAR_CODTIPO & "' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "  ", db, adOpenKeyset, adLockReadOnly
                If rstdestino.RecordCount > 0 Then
                    j = rstdestino.RecordCount
                Else
                  MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
                  Exit Sub
                End If
            End If
            If VAR_JQ = "" Then
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
                'rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
                If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
                  If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
                    MsgBox "El monto que está intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
                     'Exit Sub
                  End If
                End If
                If rs_aux1.State = 1 Then rs_aux1.Close
            End If
        Case "DYP"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DYR' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
            
        Case "ANL"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DES' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "DVL"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "RVT"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DVI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
            
            '' 02/07/2014 VERIFICAR
            'If rstdestino.State = 1 Then rstdestino.Close
            'rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
            'If rstdestino2.State = 1 Then rstdestino2.Close
            'rstdestino2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            'If rstdestino.RecordCount < 1 Or rstdestino2.RecordCount < 1 Then
            '  MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
            '  Exit Sub
            'End If
        Case Else
            MsgBox "No se ha definido el tipo " & vbCrLf & " de registro que está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
            If rstdestino.State = 1 Then rstdestino.Close
            Exit Sub
    End Select
    'If rstdestino.State = 1 Then rstdestino.Close
    '**** FIN VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************

    Dim cta_deb1 As String
    Dim Subcta_deb11 As String
    Dim Subcta_deb21 As String

    Dim cta_credito1 As String
    Dim Subcta_cred11 As String
    Dim Subcta_cred21 As String

    Dim cod_ant As Integer
    Dim org_ant As String

    v_Tipo_Comp(1, 1) = VAR_CODTIPO
    
    db.BeginTrans
'    Frmmensaje.Visible = True
'    LblMensaje.Caption = "Este proceso tomará solo unos segundos, gracias"
    '========================================
    '==== verifica si ya fue contabilizado
      yacontabilizo = 0
      Set rs_aux2 = New ADODB.Recordset
      If rs_aux2.State = 1 Then rs_aux2.Close
      rs_aux2.Open "select * from co_comprobante_m where Cod_trans = '" & VAR_CODANT & "' and org_codigo = '" & VAR_ORG & "' and tipo_comp = '" & VAR_CODTIPO & "' AND estado_codigo = 'APR'", db, adOpenKeyset, adLockOptimistic
      'rs_aux2.Open "select * from co_comprobante_m where Cod_trans = '2' and org_codigo = '111' and tipo_comp = '" & VAR_CODTIPO & "' AND estado_codigo = 'APR'", db, adOpenKeyset, adLockOptimistic
      If rs_aux2.RecordCount > 0 Then
        ' revisar para validar mejor si YA contabilizo !!
        'yacontabilizo = 1
        yacontabilizo = 0
      Else
        yacontabilizo = 0
      End If
      If yacontabilizo = 1 Then
        'MsgBox "aqui recontabilizar" & rstdestino!Cod_trans & " -- " & rstdestino!org_codigo & " / " & rstdestino!Cod_Comp
        Var_Comp = rs_aux2!Cod_Comp
      Else
        '===== ini GENERA EL CODIGO DE COMPROBANTE ====
        Set rstCodComp = New ADODB.Recordset
        rstCodComp.CursorLocation = adUseClient
        If rstCodComp.State = 1 Then rstCodComp.Close
        rstCodComp.Open "select * from fc_Correl where tipo_tramite = 'CMBTE'", db, adOpenDynamic, adLockOptimistic
        If rstCodComp.RecordCount > 0 Then
          Var_Comp = CDbl(rstCodComp!numero_correlativo)
          Var_Comp = Var_Comp + 1
          rstCodComp!numero_correlativo = Trim(Str(Var_Comp))
          rstCodComp.Update
        End If
        If rstCodComp.State = 1 Then rstCodComp.Close
        'CORRELATIVOS   (DEBEN SACAR DE fo_gastos_detalle WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
        correlv = Var_Comp           'rs_datos!gasto_codigo
        nroventa = Var_Comp          'rs_datos!gasto_codigo
        VAR_CODANT = Var_Comp        'rstdestino!gasto_codigo
        VAR_SOL = Var_Comp           'rs_datos!solicitud_codigo
        'Correlativo por Mes y Tipo de Comprobante
        Set rs_aux14 = New ADODB.Recordset
        SQL_FOR = "select numero_correlativo, tipo_tramite FROM fc_correl WHERE (cta_codigo1 = '" & Trim(VAR_MES) & "' and cta_codigo2 = '" & VAR_DOC & "' ) "
        rs_aux14.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux14.RecordCount > 0 Then
              rs_aux14!numero_correlativo = rs_aux14!numero_correlativo + 1
              VAR_COMPM = rs_aux14!numero_correlativo    'VAR_DOCR
              rs_aux14.Update
        End If
        'R-112, R-110, R-111
         Set rs_aux14 = New ADODB.Recordset
         If rs_aux14.State = 1 Then rs_aux14.Close
         SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & VAR_DOC & "'  "  ''R-112' "          '  '" & txt_codigo1 & "' "
         rs_aux14.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
         If rs_aux14.RecordCount > 0 Then
            rs_aux14!correl_doc = rs_aux14!correl_doc + 1
            'VAR_COMPM = rs_aux14!correl_doc
            rs_aux14.Update
         End If
        '===== fin TERMINA GENERACION DE COMPROBANTE =====
      '==== ini registro co_comprobante_m
         rs_aux2.AddNew
         rs_aux2("cod_comp") = Var_Comp
      End If
    '========================================
    'anterior
    '      If rstdestino.State = 1 Then rstdestino.Close
    '      rstdestino.Open "select * from co_comprobante_m where Cod_Comp = 0", db, adOpenKeyset, adLockOptimistic
    '      If rstdestino.RecordCount > 0 Then
    '      End If
    '      rstdestino.AddNew
    
    '      rstdestino("cod_comp") = Var_Comp
    'anterior
      rs_aux2("Tipo_Comp") = VAR_CODTIPO        'v_Tipo_Comp(1, i)
      rs_aux2("cod_trans") = VAR_CODANT
      rs_aux2("org_codigo") = VAR_ORG
      'If yacontabilizo = 0 Then
      '  rs_aux2("Fecha_transacion") = Date
      'Else
        rs_aux2!Fecha_transacion = IIf(Format(VAR_FFAC, "dd/mm/yyyy") = "", Date, Format(VAR_FFAC, "dd/mm/yyyy"))
      'End If
      rs_aux2("mes_trasaccion") = UCase(MonthName(Month(Date)))
      rs_aux2("ges_gestion") = IIf(gestion0 = "", Year(Date), gestion0)  'glGestion
      rs_aux2("beneficiario_codigo") = "1029203026"       'VAR_BENEF
      rs_aux2("glosa") = VAR_TCOMP + "- " + VAR_GLOSA
      rs_aux2("unidad_codigo") = VAR_COD4           'Ado_datos16.Recordset("unidad_codigo")
      rs_aux2("solicitud_codigo") = VAR_SOL         'Ado_datos16.Recordset("solicitud_codigo")
      rs_aux2("tipo_moneda") = VAR_MONEDA
      rs_aux2("unidad_codigo_ant") = VAR_CITE
      'rs_aux2!Cobranza_aux = NRO_COBR
      rs_aux2("proceso_codigo") = Left(VAR_ETAPA, 3)        '"FIN"
      rs_aux2("subproceso_codigo") = Left(VAR_ETAPA, 6)     '"FIN-02"
      rs_aux2("etapa_codigo") = VAR_ETAPA
      
      rs_aux2("clasif_codigo") = "ADM"
      'rs_aux2("doc_codigo") = "R-112"
      rs_aux2("doc_codigo") = VAR_DOC       '"R-110" o "R-112"
      rs_aux2("doc_numero") = VAR_COMPM         'Var_Comp   VAR_COMPM       '20101-9
      rs_aux2("pro_codigo_det") = IIf(VAR_PROY2 = "0", "20101-9", VAR_PROY2)
    
      rs_aux2("estado_codigo") = "APR"
      rs_aux2("usr_codigo") = glusuario
      'JALAR NRO. DE COMPRA WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
      rs_aux2!venta_compra = nroventa
      rs_aux2!cobranza_pago = NRO_COBR
      rs_aux2!glosa_contab = VAR_GLOSA
      'rs_aux2!Factura_cheque= NroFactura
      'If yacontabilizo = 0 Then
      '  rs_aux2("Fecha_registro") = Format(Date, "dd/mm/yyyy")
      '  'rs_aux2("Hora_registro") = Format(Time, "hh:mm:ss")
      'Else
      If Format(VAR_FFAC, "dd/mm/yyyy") = "" Then
        VAR_FFAC = Date
      End If
        rs_aux2("Fecha_registro") = Format(VAR_FFAC, "dd/mm/yyyy")
      'End If
      rs_aux2.Update
      'db.Execute "UPDATE co_comprobante_m SET edificio = gc_edificaciones.edif_descripcion FROM co_comprobante_m INNER JOIN gc_edificaciones ON co_comprobante_m.pro_codigo_det =gc_edificaciones.edif_codigo where co_comprobante_m.edificio Is Null "

      'db.Execute "UPDATE co_comprobante_m SET cliente = gc_beneficiario.beneficiario_denominacion FROM co_comprobante_m INNER JOIN gc_beneficiario ON co_comprobante_m.beneficiario_codigo  =gc_beneficiario.beneficiario_codigo where co_comprobante_m.cliente Is Null "
    
      'db.Execute "UPDATE co_comprobante_m SET departamento = gc_departamento.depto_descripcion FROM co_comprobante_m INNER JOIN gc_departamento ON LEFT(co_comprobante_m.pro_codigo_det,1)  =gc_departamento.depto_codigo where co_comprobante_m.departamento Is Null "

      'If VAR_TCOMP = "REF" Then
      '  db.Execute "UPDATE co_comprobante_m SET glosa_contab = 'Fac: ' NroFactura + ' - '+ unidad_codigo + ' -Edif: ' + rtrim(edificio) + ' - Benef: ' + rtrim(cliente) + ' - ' + departamento + ' - ' + right(glosa,50) where co_comprobante_m.glosa_contab is null "
      'Else
      '  db.Execute "UPDATE co_comprobante_m SET glosa_contab = unidad_codigo + ' -Edif: ' + rtrim(edificio) + ' - Benef: ' + rtrim(cliente) + ' - ' + departamento + ' - ' + right(glosa,50) where co_comprobante_m.glosa_contab is null "
      'End If
      '==== fin registro co_comprobantre_m

    Dim d_cta_nombre_1 As String
    Dim d_aux1_1 As String
    Dim d_aux2_1 As String
    Dim d_aux3_1 As String
    Dim h_cta_nombre_1 As String
    Dim h_aux1_1 As String
    Dim h_aux2_1 As String
    Dim h_aux3_1 As String
    'If rstdestino.State = 1 Then rstdestino.Close
    
    For i = 1 To j
'    ' nuevo ini
      
      If (VAR_CODTIPO = "DAP") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Or (VAR_CODTIPO = "DEY") Or (VAR_CODTIPO = "REF") Then
        cta_deb1 = rstdestino("cta_deb")
        Subcta_deb11 = rstdestino("Subcta_deb1")
        Subcta_deb21 = rstdestino("Subcta_deb2")
        
        cta_credito1 = rstdestino("cta_cred")
        Subcta_cred11 = rstdestino("Subcta_cred1")
        Subcta_cred21 = rstdestino("Subcta_cred2")
        
        VAR_PORC = rstdestino!porcentaje
      Else
        cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
        Subcta_deb11 = rstdestino!Subcta_cred1
        Subcta_deb21 = rstdestino!Subcta_cred2
    
        cta_credito1 = rstdestino!cta_deb
        Subcta_cred11 = rstdestino!Subcta_deb1
        Subcta_cred21 = rstdestino!Subcta_deb2
        
        VAR_PORC = rstdestino!porcentaje
      End If
      
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and SubCta1 = '" & Subcta_deb11 & "' and SubCta2 = '" & Subcta_deb21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        d_cta_nombre_1 = RTrim(rs_aux1("NombreCta"))
        d_aux1_1 = rs_aux1("aux1")
        d_aux2_1 = rs_aux1("aux2")
        d_aux3_1 = rs_aux1("aux3")
        VAR_DCORR = rs_aux1("correl")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and SubCta1 = '" & Subcta_cred11 & "' and SubCta2 = '" & Subcta_cred21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        h_cta_nombre_1 = RTrim(rs_aux1("NombreCta"))
        h_aux1_1 = rs_aux1("aux1")
        h_aux2_1 = rs_aux1("aux2")
        h_aux3_1 = rs_aux1("aux3")
        VAR_HCORR = rs_aux1("correl")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and nivel = '4' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        VAR_NOMD = rs_aux1("NombreCta")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and nivel = '4' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        VAR_NOMH = rs_aux1("NombreCta")
      End If
    ' nuevo fin
    
      '===== ini registra CO_diaRIO =========
      Set rstdestino2 = New ADODB.Recordset
      If rstdestino2.State = 1 Then rstdestino2.Close
      rstdestino2.Open "select * from co_diario where Cod_Comp = " & Var_Comp, db, adOpenKeyset, adLockOptimistic
      'If rstdestino2.RecordCount > 0 Then
      '  MsgBox "Ya Existe el asiento, se reemplazará con los nuevos datos..."
      'Else
        rstdestino2.AddNew
        rstdestino2("Cod_Comp") = Var_Comp
      'End If
        rstdestino2("Cod_Comp_Detalle") = rstdestino2.RecordCount
      'rstdestino2("Tipo_Comp") = "DEI"   'v_Tipo_Comp(1, i)
      'rstdestino2("Cod_Comp_C") = Var_Comp
      'If v_Tipo_Comp(1, i) = "DEI" Or v_Tipo_Comp(1, i) = "REC" Then
      If (VAR_CODTIPO = "DAP") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Or (VAR_CODTIPO = "DEY") Or (VAR_CODTIPO = "REF") Then
        rstdestino2("D_Cuenta") = cta_deb1
        rstdestino2("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino2("D_Subcta1") = Subcta_deb11
        rstdestino2("D_SubCta2") = Subcta_deb21
        rstdestino2("D_Aux1") = d_aux1_1
        rstdestino2("D_Aux2") = d_aux2_1
        rstdestino2("D_Aux3") = d_aux3_1
        rstdestino2("NOMCTADEBE") = VAR_NOMD
        rstdestino2("d_Correl") = VAR_DCORR
        ' para Aux1
        ' ini PARA EL FUTURO ******** REVISAR
'        Set rs_aux4 = New ADODB.Recordset
'        If rs_aux4.State = 1 Then rs_aux4.Close
'        SQL_FOR = "select * from cc_tipo_auxiliar where aux = '" & d_aux1_1 & "' "
'        rs_aux4.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux4.RecordCount > 0 Then
'            Set rs_aux1 = New ADODB.Recordset
'            If rs_aux1.State = 1 Then rs_aux1.Close
'            SQL_FOR = "select * from " + rs_aux4!NombreTabla + " where " + rs_aux4!nombre_codigo + " = " + VAR_COD1
'            rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'            If rs_aux1.RecordCount > 0 Then
'        Else
'        End If
        ' fin PARA EL FUTURO ******** REVISAR
        If cta_deb1 = "5111" And Subcta_deb11 = "01" And cta_credito1 = "2111" And Subcta_cred11 = "01" Then
            rstdestino2("D_MontoBs") = VAR_BS2
            VAR_DOL2 = VAR_BS2 / GlTipoCambioMercado
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "01" And cta_credito1 = "2111" And Subcta_cred11 = "10" Then
            rstdestino2("D_MontoBs") = VAR_DSCTO_BS
            VAR_DOL2 = VAR_DSCTO_BS / GlTipoCambioMercado
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "01" And cta_credito1 = "2111" And Subcta_cred11 = "08" Then
            rstdestino2("D_MontoBs") = VAR_AFP1_BS
            VAR_DOL2 = VAR_AFP1_BS / GlTipoCambioMercado
            VAR_BENEF = VAR_BENEF1
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "01" And cta_credito1 = "2111" And Subcta_cred11 = "09" Then
            rstdestino2("D_MontoBs") = VAR_AFP2_BS
            VAR_DOL2 = VAR_AFP2_BS / GlTipoCambioMercado
            VAR_BENEF = VAR_BENEF2
        End If
        
        If cta_deb1 = "5111" And Subcta_deb11 = "02" And cta_credito1 = "2111" And Subcta_cred11 = "02" Then
            rstdestino2("D_MontoBs") = VAR_AFP1_BS2
            VAR_DOL2 = VAR_AFP1_BS2 / GlTipoCambioMercado
            VAR_BENEF = VAR_BENEF1
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "03" And cta_credito1 = "2111" And Subcta_cred11 = "03" Then
            rstdestino2("D_MontoBs") = VAR_AFP1_BS2
            VAR_DOL2 = VAR_AFP1_BS2 / GlTipoCambioMercado
            VAR_BENEF = VAR_BENEF2
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "04" And cta_credito1 = "2111" And Subcta_cred11 = "04" Then
            rstdestino2("D_MontoBs") = VAR_SS_BS
            VAR_DOL2 = VAR_SS_BS / GlTipoCambioMercado
            VAR_BENEF = VAR_BENEF3
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "05" And cta_credito1 = "2111" And Subcta_cred11 = "05" Then
            rstdestino2("D_MontoBs") = VAR_FV_BS
            VAR_DOL2 = VAR_FV_BS / GlTipoCambioMercado
            VAR_BENEF = VAR_BENEF4
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "06" And cta_credito1 = "2111" And Subcta_cred11 = "06" Then
            rstdestino2("D_MontoBs") = VAR_INDE_BS
            VAR_DOL2 = VAR_INDE_BS / GlTipoCambioMercado
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "07" And cta_credito1 = "2111" And Subcta_cred11 = "07" Then
            rstdestino2("D_MontoBs") = VAR_AGUI_BS
            VAR_DOL2 = VAR_AGUI_BS / GlTipoCambioMercado
        End If

        Select Case d_aux1_1
            Case "01"
                rstdestino2("D_Cta_Aux1") = VAR_BENEF
                Call DESCAUX(d_aux1_1, CStr(VAR_BENEF))    'DESAUX =
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
                Call DESCAUX(d_aux1_1, CStr(VAR_CTA))
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
                Call DESCAUX(d_aux1_1, CStr(VAR_PROY2))
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "04"
                rstdestino2("D_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux1_1, CStr(VAR_COD4))
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "05"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux1") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux1_1, rstdestino2!D_Cta_Aux1)
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "07"
                rstdestino2("D_Cta_Aux1") = VAR_PLA
                Call DESCAUX(d_aux1_1, rstdestino2!D_Cta_Aux1)
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "08"
                rstdestino2("D_Cta_Aux1") = VAR_SPLA
                Call DESCAUX(d_aux1_1, rstdestino2!D_Cta_Aux1)
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "09"
                rstdestino2("D_Cta_Aux1") = VAR_ORG
                Call DESCAUX(d_aux1_1, CStr(VAR_ORG))
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "10"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
        End Select
        rstdestino2!D_Des_Aux1 = DESAUX
        Select Case d_aux2_1
            Case "01"
                rstdestino2("D_Cta_Aux2") = VAR_BENEF
                Call DESCAUX(d_aux2_1, CStr(VAR_BENEF))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
                Call DESCAUX(d_aux2_1, CStr(VAR_CTA))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
                Call DESCAUX(d_aux2_1, CStr(VAR_PROY2))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "04"
                rstdestino2("D_Cta_Aux2") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux2_1, CStr(VAR_COD4))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "05"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux2") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux2_1, rstdestino2!D_Cta_Aux2)
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "07"
                rstdestino2("D_Cta_Aux2") = VAR_PLA
                Call DESCAUX(d_aux2_1, rstdestino2!D_Cta_Aux2)
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "08"
                rstdestino2("D_Cta_Aux2") = VAR_SPLA
                Call DESCAUX(d_aux2_1, rstdestino2!D_Cta_Aux2)
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "09"
                rstdestino2("D_Cta_Aux2") = VAR_ORG
                Call DESCAUX(d_aux2_1, CStr(VAR_ORG))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "10"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
        End Select
        rstdestino2!D_Des_Aux2 = DESAUX
        Select Case d_aux3_1
            Case "01"
                rstdestino2("D_Cta_Aux3") = VAR_BENEF
                Call DESCAUX(d_aux3_1, CStr(VAR_BENEF))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
                Call DESCAUX(d_aux3_1, CStr(VAR_CTA))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
                Call DESCAUX(d_aux3_1, CStr(VAR_PROY2))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "04"
                rstdestino2("D_Cta_Aux3") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux3_1, CStr(VAR_COD4))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux3_1, rstdestino2!D_Cta_Aux3)
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "07"
                rstdestino2("D_Cta_Aux3") = VAR_PLA
                Call DESCAUX(d_aux3_1, rstdestino2!D_Cta_Aux3)
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "08"
                rstdestino2("D_Cta_Aux3") = VAR_SPLA
                Call DESCAUX(d_aux3_1, rstdestino2!D_Cta_Aux3)
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "09"
                rstdestino2("D_Cta_Aux3") = VAR_ORG
                Call DESCAUX(d_aux3_1, CStr(VAR_ORG))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "10"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
        End Select
        rstdestino2!D_Des_Aux3 = DESAUX
        
'        If d_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR

        'VAR_PORC Definido en el Relacionador
'
'        If VAR_PORC = "0.87" Then
'            rstdestino2("D_MontoBs") = VAR_87       'VAR_BS2 * VAR_PORC
'            rstdestino2("D_MontoDl") = VAR_87 * GlTipoCambioOficial  'VAR_DOL2 * VAR_PORC
'        End If
'        If VAR_PORC = "0.13" Then
'            rstdestino2("D_MontoBs") = VAR_13       'VAR_BS2 * VAR_PORC
'            rstdestino2("D_MontoDl") = VAR_13 * GlTipoCambioOficial  'VAR_DOL2 * VAR_PORC
'        End If
'        If VAR_PORC <> "0.87" And VAR_PORC <> "0.13" Then
'            rstdestino2("D_MontoBs") = VAR_BS2 * VAR_PORC
'            rstdestino2("D_MontoDl") = VAR_DOL2 * VAR_PORC
'        End If
        rstdestino2("D_MontoDl") = Round(VAR_DOL2, 2)
        rstdestino2("D_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
        'AQUI MONEDA 02/07/01
        'rstdestino2("D_Cambio") = GlTipoCambioMercado
        'AAAAAAAAAAAAAAQQQQQQQQQQQQQQQQUUUUUUUUUUUUUUUUIIIIIIIIIIIII JQA
        rstdestino2("H_Cuenta") = cta_credito1
        rstdestino2("H_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino2("H_SubCta1") = Subcta_cred11
        rstdestino2("H_SubCta2") = Subcta_cred21
        rstdestino2("H_Aux1") = h_aux1_1
        rstdestino2("H_Aux2") = h_aux2_1
        rstdestino2("H_Aux3") = h_aux3_1
        rstdestino2("NOMCTAHABER") = VAR_NOMH
        rstdestino2("h_Correl") = VAR_HCORR
        'rstdestino2("H_Cta_Aux1") = ""
        If cta_deb1 = "5111" And Subcta_deb11 = "01" And cta_credito1 = "2111" And Subcta_cred11 = "01" Then
            rstdestino2("H_MontoBs") = VAR_BS2
            VAR_DOL2 = VAR_BS2 / GlTipoCambioMercado
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "01" And cta_credito1 = "2111" And Subcta_cred11 = "10" Then
            rstdestino2("H_MontoBs") = VAR_DSCTO_BS
            VAR_DOL2 = VAR_DSCTO_BS / GlTipoCambioMercado
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "01" And cta_credito1 = "2111" And Subcta_cred11 = "08" Then
            rstdestino2("H_MontoBs") = VAR_AFP1_BS
            VAR_DOL2 = VAR_AFP1_BS / GlTipoCambioMercado
            VAR_BENEF = VAR_BENEF1
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "01" And cta_credito1 = "2111" And Subcta_cred11 = "09" Then
            rstdestino2("H_MontoBs") = VAR_AFP2_BS
            VAR_DOL2 = VAR_AFP2_BS / GlTipoCambioMercado
            VAR_BENEF = VAR_BENEF2
        End If
        
        If cta_deb1 = "5111" And Subcta_deb11 = "02" And cta_credito1 = "2111" And Subcta_cred11 = "02" Then
            rstdestino2("H_MontoBs") = VAR_AFP1_BS2
            VAR_DOL2 = VAR_AFP1_BS2 / GlTipoCambioMercado
            VAR_BENEF = VAR_BENEF1
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "03" And cta_credito1 = "2111" And Subcta_cred11 = "03" Then
            rstdestino2("H_MontoBs") = VAR_AFP1_BS2
            VAR_DOL2 = VAR_AFP1_BS2 / GlTipoCambioMercado
            VAR_BENEF = VAR_BENEF2
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "04" And cta_credito1 = "2111" And Subcta_cred11 = "04" Then
            rstdestino2("H_MontoBs") = VAR_SS_BS
            VAR_DOL2 = VAR_SS_BS / GlTipoCambioMercado
            VAR_BENEF = VAR_BENEF3
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "05" And cta_credito1 = "2111" And Subcta_cred11 = "05" Then
            rstdestino2("H_MontoBs") = VAR_FV_BS
            VAR_DOL2 = VAR_FV_BS / GlTipoCambioMercado
            VAR_BENEF = VAR_BENEF4
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "06" And cta_credito1 = "2111" And Subcta_cred11 = "06" Then
            rstdestino2("H_MontoBs") = VAR_INDE_BS
            VAR_DOL2 = VAR_INDE_BS / GlTipoCambioMercado
        End If
        If cta_deb1 = "5111" And Subcta_deb11 = "07" And cta_credito1 = "2111" And Subcta_cred11 = "07" Then
            rstdestino2("H_MontoBs") = VAR_AGUI_BS
            VAR_DOL2 = VAR_AGUI_BS / GlTipoCambioMercado
        End If
        Select Case h_aux1_1
            Case "01"
                rstdestino2("H_Cta_Aux1") = VAR_BENEF
                Call DESCAUX(h_aux1_1, CStr(VAR_BENEF))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
                Call DESCAUX(h_aux1_1, CStr(VAR_CTA))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
                Call DESCAUX(h_aux1_1, CStr(VAR_PROY2))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "04"
                rstdestino2("H_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux1_1, CStr(VAR_COD4))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "05"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux1") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux1_1, rstdestino2!H_Cta_Aux1)
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "07"
                rstdestino2("H_Cta_Aux1") = VAR_PLA
                Call DESCAUX(h_aux1_1, rstdestino2!H_Cta_Aux1)
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "08"
                rstdestino2("H_Cta_Aux1") = VAR_SPLA
                Call DESCAUX(h_aux1_1, rstdestino2!H_Cta_Aux1)
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "09"
                rstdestino2("H_Cta_Aux1") = VAR_ORG
                Call DESCAUX(h_aux1_1, CStr(VAR_ORG))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "10"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
        End Select
        rstdestino2!H_Des_Aux1 = DESAUX
        
        Select Case h_aux2_1
            Case "01"
                rstdestino2("H_Cta_Aux2") = VAR_BENEF
                Call DESCAUX(h_aux2_1, CStr(VAR_BENEF))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
                Call DESCAUX(h_aux2_1, CStr(VAR_CTA))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
                Call DESCAUX(h_aux2_1, CStr(VAR_PROY2))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "04"
                rstdestino2("H_Cta_Aux2") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux2_1, CStr(VAR_COD4))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux2_1, rstdestino2!H_Cta_Aux2)
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "07"
                rstdestino2("H_Cta_Aux2") = VAR_PLA
                Call DESCAUX(h_aux2_1, rstdestino2!H_Cta_Aux2)
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "08"
                rstdestino2("H_Cta_Aux2") = VAR_SPLA
                Call DESCAUX(h_aux2_1, rstdestino2!H_Cta_Aux2)
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "09"
                rstdestino2("H_Cta_Aux2") = VAR_ORG
                Call DESCAUX(h_aux2_1, CStr(VAR_ORG))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "10"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
        End Select
        rstdestino2!H_Des_Aux2 = DESAUX
        Select Case h_aux3_1
            Case "01"
                rstdestino2("H_Cta_Aux3") = VAR_BENEF
                Call DESCAUX(h_aux3_1, CStr(VAR_BENEF))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
                Call DESCAUX(h_aux3_1, CStr(VAR_CTA))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
                Call DESCAUX(h_aux3_1, CStr(VAR_PROY2))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "04"
                rstdestino2("H_Cta_Aux3") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux3_1, CStr(VAR_COD4))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "05"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux3") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux3_1, rstdestino2!H_Cta_Aux3)
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "07"
                rstdestino2("H_Cta_Aux3") = VAR_PLA
                Call DESCAUX(h_aux3_1, rstdestino2!H_Cta_Aux3)
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "08"
                rstdestino2("H_Cta_Aux3") = VAR_SPLA
                Call DESCAUX(h_aux3_1, rstdestino2!H_Cta_Aux3)
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "09"
                rstdestino2("H_Cta_Aux3") = VAR_ORG
                Call DESCAUX(h_aux3_1, CStr(VAR_ORG))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "10"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
        End Select
        rstdestino2!H_Des_Aux3 = DESAUX
        
'        If h_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
 
        rstdestino2("H_MontoDl") = Round(VAR_DOL2, 2)
        rstdestino2("H_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
        rstdestino2!cobranza_pago = NRO_COBR
      End If

      'If (v_Tipo_Comp(1, i) = "DES") Or (v_Tipo_Comp(1, i) = "ANI") Then
      If (VAR_CODTIPO = "DES") Or (VAR_CODTIPO = "ANI") Or (VAR_CODTIPO = "DVI") Then
        'desafecta un devengado
        rstdestino2("D_Cuenta") = cta_credito1
        rstdestino2("D_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino2("D_Subcta1") = Subcta_cred11
        rstdestino2("D_SubCta2") = Subcta_cred21
        rstdestino2("D_Aux1") = h_aux1_1
        rstdestino2("D_Aux2") = h_aux2_1
        rstdestino2("D_Aux3") = h_aux3_1
'        rstdestino2("D_Cta_Aux1") = "VESCT"
        Select Case h_aux1_1
            Case "01"
                rstdestino2("D_Cta_Aux1") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux1") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("D_Cta_Aux1") = ""
            Case "06"
                rstdestino2("D_Cta_Aux1") = ""
            Case "07"
                rstdestino2("D_Cta_Aux1") = ""
            Case "08"
                rstdestino2("D_Cta_Aux1") = ""
            Case "09"
                rstdestino2("D_Cta_Aux1") = VAR_ORG
            Case "10"
                rstdestino2("D_Cta_Aux1") = ""
            Case "11"
                rstdestino2("D_Cta_Aux1") = ""
            Case "12"
                rstdestino2("D_Cta_Aux1") = ""
            Case "00"
                rstdestino2("D_Cta_Aux1") = ""
        End Select
        
        Select Case h_aux2_1
            Case "01"
                rstdestino2("D_Cta_Aux2") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux2") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("D_Cta_Aux2") = ""
            Case "06"
                rstdestino2("D_Cta_Aux2") = ""
            Case "07"
                rstdestino2("D_Cta_Aux2") = ""
            Case "08"
                rstdestino2("D_Cta_Aux2") = ""
            Case "09"
                rstdestino2("D_Cta_Aux2") = VAR_ORG
            Case "10"
                rstdestino2("D_Cta_Aux2") = ""
            Case "11"
                rstdestino2("D_Cta_Aux2") = ""
            Case "12"
                rstdestino2("D_Cta_Aux2") = ""
            Case "00"
                rstdestino2("D_Cta_Aux2") = ""
        End Select
        
        Select Case h_aux3_1
            Case "01"
                rstdestino2("D_Cta_Aux3") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux3") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = ""
            Case "07"
                rstdestino2("D_Cta_Aux3") = ""
            Case "08"
                rstdestino2("D_Cta_Aux3") = ""
            Case "09"
                rstdestino2("D_Cta_Aux3") = VAR_ORG
            Case "10"
                rstdestino2("D_Cta_Aux3") = ""
            Case "11"
                rstdestino2("D_Cta_Aux3") = ""
            Case "12"
                rstdestino2("D_Cta_Aux3") = ""
            Case "00"
                rstdestino2("D_Cta_Aux3") = ""
        End Select
'        If h_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
        rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
        rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
        rstdestino2("D_Cambio") = GlTipoCambioMercado

        rstdestino2("H_Cuenta") = cta_deb1
        rstdestino2("H_Nombre") = d_cta_nombre_1  ' CAMPO PARA ELIMINAR
        rstdestino2("H_SubCta1") = Subcta_deb11
        rstdestino2("H_SubCta2") = Subcta_deb21
        rstdestino2("H_Aux1") = d_aux1_1
        rstdestino2("H_Aux2") = d_aux2_1
        rstdestino2("H_Aux3") = d_aux3_1
'        rstdestino2("H_Cta_Aux1") = "VESCT"
        Select Case d_aux1_1
            Case "01"
                rstdestino2("H_Cta_Aux1") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux1") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("H_Cta_Aux1") = ""
            Case "06"
                rstdestino2("H_Cta_Aux1") = ""
            Case "07"
                rstdestino2("H_Cta_Aux1") = ""
            Case "08"
                rstdestino2("H_Cta_Aux1") = ""
            Case "09"
                rstdestino2("H_Cta_Aux1") = VAR_ORG
            Case "10"
                rstdestino2("H_Cta_Aux1") = ""
            Case "11"
                rstdestino2("H_Cta_Aux1") = ""
            Case "12"
                rstdestino2("H_Cta_Aux1") = ""
            Case "00"
                rstdestino2("H_Cta_Aux1") = ""
        End Select
        
        Select Case d_aux2_1
            Case "01"
                rstdestino2("H_Cta_Aux2") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux2") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = ""
            Case "07"
                rstdestino2("H_Cta_Aux2") = ""
            Case "08"
                rstdestino2("H_Cta_Aux2") = ""
            Case "09"
                rstdestino2("H_Cta_Aux2") = VAR_ORG
            Case "10"
                rstdestino2("H_Cta_Aux2") = ""
            Case "11"
                rstdestino2("H_Cta_Aux2") = ""
            Case "12"
                rstdestino2("H_Cta_Aux2") = ""
            Case "00"
                rstdestino2("H_Cta_Aux2") = ""
        End Select
        
        Select Case d_aux3_1
            Case "01"
                rstdestino2("H_Cta_Aux3") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux3") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("H_Cta_Aux3") = ""
            Case "06"
                rstdestino2("H_Cta_Aux3") = ""
            Case "07"
                rstdestino2("H_Cta_Aux3") = ""
            Case "08"
                rstdestino2("H_Cta_Aux3") = ""
            Case "09"
                rstdestino2("H_Cta_Aux3") = VAR_ORG
            Case "10"
                rstdestino2("H_Cta_Aux3") = ""
            Case "11"
                rstdestino2("H_Cta_Aux3") = ""
            Case "12"
                rstdestino2("H_Cta_Aux3") = ""
            Case "00"
                rstdestino2("H_Cta_Aux3") = ""
        End Select
'        If d_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        rstdestino2("H_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
        rstdestino2("H_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
        rstdestino2("H_Cambio") = GlTipoCambioMercado
        rstdestino2!cobranza_pago = NRO_COBR
      End If

'      '==== INI DVI ====
'      If (VAR_CODTIPO = "DVI") Then
'        rstdestino2("D_Cuenta") = cta_deb1
''        rstdestino2("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino2("D_Subcta1") = Subcta_deb11
'        rstdestino2("D_SubCta2") = Subcta_deb21
'        rstdestino2("D_Aux1") = d_aux1_1
'        rstdestino2("D_Aux2") = d_aux2_1
'        rstdestino2("D_Aux3") = d_aux3_1
'        If d_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
''        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
'        rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
'        rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
'        rstdestino2("D_Cambio") = GlTipoCambioMercado
'        rstdestino2("H_Cuenta") = cta_credito1
''        rstdestino2("H_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino2("H_SubCta1") = Subcta_cred11
'        rstdestino2("H_SubCta2") = Subcta_cred21
'        rstdestino2("H_Aux1") = h_aux1_1
'        rstdestino2("H_Aux2") = h_aux2_1
'        rstdestino2("H_Aux3") = h_aux3_1
'        'rstdestino2("H_Cta_Aux1") = "VESCT"
'        If h_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
''        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
'        rstdestino2("H_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
'        rstdestino2("H_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
'        rstdestino2("H_Cambio") = GlTipoCambioMercado
'      End If
'      '==== FIN DVI ====
      rstdestino2("Usr_codigo") = glusuario
      'If yacontabilizo = 0 Then
      '  rstdestino2("Fecha_registro") = Date
      ''  rstdestino2("Hora_registro") = Format(Time, "hh:mm:ss")
      'Else
        rstdestino2("Fecha_registro") = VAR_FFAC
      'End If
      
      rstdestino2.Update
      If rstdestino2.State = 1 Then rstdestino2.Close
      '======= fin registra co_diario ==========
      rstdestino.MoveNext
    Next i
    '======= inI Actualiza campos de estatus de ingresos ==========
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '" & correlativo1 & "' and org_codigo = '" & VAR_ORG & "' and ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' ", db, adOpenDynamic, adLockOptimistic
'    rstdestino.MoveFirst
'    If Not (rstdestino.EOF) Then
'      rstdestino("estado_aprobacion") = "S"
'        If VAR_CODTIPO = "DEI" Then
'          rstdestino("estado_devengado") = "S"
'        End If
'        If VAR_CODTIPO = "REC" Then
'          rstdestino("estado_recaudado") = "S"
'        End If
'        If VAR_CODTIPO = "DYR" Then
'          rstdestino("estado_devengado") = "S"
'          rstdestino("estado_recaudado") = "S"
'        End If
'
'        If VAR_CODTIPO = "DES" Then
'          rstdestino("estado_desafectado") = "S"
'        End If
'        If VAR_CODTIPO = "ANI" Then
'          rstdestino("estado_anulado") = "S"
'        End If
'        If VAR_CODTIPO = "DVI" Then
'          rstdestino!estado_desafectado = "S"
'          rstdestino!estado_anulado = "S"
'        End If
'       rstdestino.Update
'       If rstdestino.State = 1 Then rstdestino.Close
'    End If
    '======= fin Actualiza campos de estatus de ingresos ==========
    ' AAAAAAAAAQQQQQQQQQQQUUUUUUUUUUUIIIIIIIIIII
    cod_ant = 0
    org_ant = ""
    '======= ini Actualiza el monto recaudado  ==========
    If (VAR_CODTIPO = "REC") Then
      '      If rstdestino.State = 1 Then rstdestino.Close
      '      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      '      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
      '        cod_ant = rstdestino("ingreso_codigo_anterior")
      '        org_ant = rstdestino("org_codigo")
      '      End If
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      'rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") + VAR_DOL2
          rstdestino("monto_recaudado_bolivianos") = rstdestino("monto_recaudado_bolivianos") + VAR_BS2
          rstdestino.Update
      End If
      If rstdestino.State = 1 Then rstdestino.Close
    End If

    If (VAR_CODTIPO = "DES") Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
'      Print VAR_CODANT
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'        cod_ant = IIf(IsNull(rstdestino("ingreso_codigo_anterior")), 0, rstdestino("ingreso_codigo_anterior"))
'        org_ant = rstdestino("org_codigo")
'      End If

      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        If rstdestino("codigo_tipo") = "DEI" Then 'And VAR_CODTIPO = "DES"
'          rstdestino!estado_desafectado = "S" 02/07/01
          rstdestino!estado_codigo = "DES"
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
        Else
          rstdestino("estado_codigo") = "DES"
'          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - VAR_DOL2
          cod_ant = IIf(IsNull(rstdestino("ingreso_codigo_anterior")), 0, rstdestino("ingreso_codigo_anterior"))
          org_ant = rstdestino("org_codigo")
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
          'rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & cod_ant & " and org_codigo = '" & org_ant & "' ", db, adOpenKeyset, adLockOptimistic
          rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
          If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
            rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - VAR_DOL2
            rstdestino("monto_recaudado_bolivianos") = rstdestino("monto_recaudado_bolivianos") - VAR_BS2
          End If
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
        End If
      End If
    End If

    If (VAR_CODTIPO = "ANI") Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        If rstdestino("codigo_tipo") = "REC" Then
'          rstdestino("estado_desafectado") = ""
          rstdestino("estado_codigo") = "ANI"
'          rstdestino("estado_devengado") = "S" 02/07/01
'          rstdestino("estado_anulado") = ""
'          rstdestino("codigo_tipo") = "DEI" 02/07/01
          rstdestino("monto_recaudado_dolares") = 0
        End If
      End If
      rstdestino.Update
'      Print rstdestino!ingreso_codigo_anterior
'      Print rstdestino!monto_recaudado
      cod_ant = 0
      org_ant = ""
      
      'Call f_actual_rec(rstdestino!org_codigo, rstdestino!ingreso_codigo_anterior)
      If rstdestino.State = 1 Then rstdestino.Close
    End If
    If (VAR_CODTIPO = "DVI") Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        rstdestino!estado_codigo = "DVI"
      End If
      rstdestino.Update
      If rstdestino.State = 1 Then rstdestino.Close
    End If
    '======= fin Actualiza el monto recaudado  ==========

    '======= ini Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
    If VAR_CODTIPO = "REC" Or VAR_CODTIPO = "DYR" Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & VAR_CTA & "'", db, adOpenKeyset, adLockOptimistic
      If Not rstdestino.EOF Then
        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + VAR_BS2
        rstdestino.Update
      End If
    End If
    If VAR_CODTIPO = "ANI" Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & VAR_CTA & "'", db, adOpenKeyset, adLockOptimistic
      If Not rstdestino.EOF Then
        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + VAR_BS2
        rstdestino.Update
      End If
    End If
    '======= fin Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
    'LblMensaje.Caption = "El proceso concluyó exitosamente, gracias"
    'Frmmensaje.Visible = False
    db.CommitTrans
  'End If
'  'marca1 = Ado_datos.Recordset.Bookmark
'  rs_datos.Update
'  rs_datos.Requery
'  Set Ado_datos.Recordset = rs_datos
'  If rs_datos.RecordCount > 0 Then
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
'  'db.Execute "EXEC ts_mf_ActualizaCtaBancaria"

End Sub

Private Function DESCAUX(VARAUX As String, VARCODIG As String)
    'CONTABILIZA PLANILLA (AUXILIARES)
    Set rsAuxDetalle = New ADODB.Recordset
    If rsAuxDetalle.State = 1 Then rsAuxDetalle.Close
    Select Case VARAUX
        Case "01"
            rsAuxDetalle.Open "SELECT beneficiario_denominacion AS DESAUX2 FROM gc_beneficiario where beneficiario_codigo = '" & VARCODIG & "' ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT beneficiario_denominacion AS DESAUX FROM gc_beneficiario where beneficiario_codigo = '" & VARCODIG & "' "
        Case "02"
            rsAuxDetalle.Open "SELECT cta_descripcion AS DESAUX2 FROM fc_cuenta_bancaria where Cta_codigo = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT cta_descripcion AS DESAUX FROM fc_cuenta_bancaria where Cta_codigo = '" & VARCODIG & "' "
        Case "03"
            rsAuxDetalle.Open "SELECT pro_codigo_det_descripcion AS DESAUX2 FROM fo_proyectos_ejecucion where pro_codigo_det = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT pro_codigo_det_descripcion AS DESAUX FROM fo_proyectos_ejecucion where pro_codigo_det = '" & VARCODIG & "' "
        Case "04"
            rsAuxDetalle.Open "SELECT unidad_descripcion AS DESAUX2 FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "05"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "06"
            rsAuxDetalle.Open "SELECT depto_descripcion AS DESAUX2 FROM gc_departamento where depto_codigo = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT depto_descripcion AS DESAUX FROM gc_departamento where depto_codigo = '" & VARCODIG & "' "
        Case "07"
            'DESAUX = ""
            rsAuxDetalle.Open "SELECT planilla_descripcion AS DESAUX2 FROM rc_planilla_grupo where planilla_codigo = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
        Case "08"
            'DESAUX = ""
            rsAuxDetalle.Open "SELECT unidad_descripcion_pla AS DESAUX2 FROM rc_planilla_sub_grupo where unidad_codigo_pla = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
        Case "09"
            rsAuxDetalle.Open "SELECT Org_descripcion AS DESAUX2 FROM fc_organismo_financiamiento where org_codigo = '" & VARCODIG & "' ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT Org_descripcion AS DESAUX FROM fc_organismo_financiamiento where org_codigo = '" & VARCODIG & "' "
        Case "10"
            'db.Execute "SELECT impuesto_descripcion AS DESAUX FROM fc_impuestos where impuesto_codigo = '" & VARCODIG & "' "
        Case "11"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "12"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "00"
            DESAUX = ""
            DESAUX2 = ""
    End Select
    If rsAuxDetalle.RecordCount > 0 Then
      DESAUX = RTrim(rsAuxDetalle!DESAUX2)
    Else
      DESAUX = ""
    End If
End Function



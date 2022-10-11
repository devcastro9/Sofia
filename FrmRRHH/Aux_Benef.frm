VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Aux_benef 
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGrabaCto 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8880
      Picture         =   "Aux_Benef.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Nuevo Registro"
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Fra_ABM 
      Caption         =   "Registro de Contratos"
      ForeColor       =   &H00C00000&
      Height          =   4335
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8775
      Begin VB.Frame Frame13 
         Height          =   690
         Left            =   8040
         TabIndex        =   45
         Top             =   120
         Width           =   615
         Begin VB.Image Image2 
            Height          =   540
            Left            =   15
            Picture         =   "Aux_Benef.frx":058A
            Top             =   105
            Width           =   555
         End
      End
      Begin VB.TextBox TxtAprob 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "cod_est_contrato"
         DataSource      =   "Ado_Contrato"
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
         Height          =   315
         Left            =   2805
         TabIndex        =   22
         Text            =   "NO"
         Top             =   520
         Width           =   495
      End
      Begin VB.ComboBox Txtestado 
         DataField       =   "fechas_confirmado"
         DataSource      =   "Ado_Contrato"
         Height          =   315
         ItemData        =   "Aux_Benef.frx":0912
         Left            =   1800
         List            =   "Aux_Benef.frx":091C
         TabIndex        =   21
         Text            =   "SI"
         Top             =   520
         Width           =   660
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "codigo_contrato"
         DataSource      =   "Ado_Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   520
         Width           =   1335
      End
      Begin VB.TextBox txtObjContrato 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "objeto_contrato"
         DataSource      =   "Ado_Contrato"
         Height          =   525
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1000
         Width           =   6855
      End
      Begin VB.TextBox TxtForm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "id_contrato"
         DataSource      =   "Ado_Contrato"
         Enabled         =   0   'False
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
         Left            =   5280
         TabIndex        =   18
         Top             =   520
         Width           =   855
      End
      Begin VB.TextBox TxtBs 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "monto_totalBS"
         DataSource      =   "Ado_Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   520
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo Dtc_descrip 
         Bindings        =   "Aux_Benef.frx":0928
         DataField       =   "codigo_unidad"
         DataSource      =   "Ado_Contrato"
         Height          =   315
         Left            =   4440
         TabIndex        =   23
         Top             =   2400
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "Uni_descripcion_larga"
         BoundColumn     =   "codigo_unidad"
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
      Begin MSDataListLib.DataCombo DtcPryDes 
         Bindings        =   "Aux_Benef.frx":0940
         DataField       =   "Pro_proyecto"
         DataSource      =   "Ado_Contrato"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   3795
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "Pro_descripcion_larga"
         BoundColumn     =   "Pro_proyecto"
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
      Begin MSComCtl2.DTPicker DTPFFirma 
         DataField       =   "fecha_firma"
         DataSource      =   "Ado_Contrato"
         Height          =   285
         Left            =   1695
         TabIndex        =   25
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   101253121
         CurrentDate     =   40471
      End
      Begin MSDataListLib.DataCombo DtcPuestoDes 
         Bindings        =   "Aux_Benef.frx":0955
         DataField       =   "codigo_puesto"
         DataSource      =   "Ado_Contrato"
         Height          =   315
         Left            =   4440
         TabIndex        =   26
         Top             =   3795
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "denominacion_puesto"
         BoundColumn     =   "codigo_puesto"
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
      Begin MSDataListLib.DataCombo DtcPuesto 
         Bindings        =   "Aux_Benef.frx":0970
         DataField       =   "codigo_puesto"
         DataSource      =   "Ado_Contrato"
         Height          =   315
         Left            =   6240
         TabIndex        =   27
         Top             =   3405
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "codigo_puesto"
         BoundColumn     =   "codigo_puesto"
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
      Begin MSComCtl2.DTPicker DTPFInicio 
         DataField       =   "fecha_inicio"
         DataSource      =   "Ado_Contrato"
         Height          =   285
         Left            =   4560
         TabIndex        =   28
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   101253121
         CurrentDate     =   40471
      End
      Begin MSComCtl2.DTPicker DTPFFin 
         DataField       =   "fecha_fin"
         DataSource      =   "Ado_Contrato"
         Height          =   285
         Left            =   7200
         TabIndex        =   29
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   101253121
         CurrentDate     =   40471
      End
      Begin MSDataListLib.DataCombo Dtc_codigo 
         Bindings        =   "Aux_Benef.frx":098B
         DataField       =   "codigo_unidad"
         DataSource      =   "Ado_Contrato"
         Height          =   315
         Left            =   6240
         TabIndex        =   30
         Top             =   2060
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "codigo_unidad"
         BoundColumn     =   "codigo_unidad"
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
      Begin MSDataListLib.DataCombo DtcOrgDes 
         Bindings        =   "Aux_Benef.frx":09A3
         DataField       =   "Codigo_Convenio"
         DataSource      =   "Ado_Contrato"
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   3120
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "Denominacion_Convenio"
         BoundColumn     =   "Codigo_Convenio"
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
      Begin MSDataListLib.DataCombo DtcOrg 
         Bindings        =   "Aux_Benef.frx":09B8
         DataField       =   "Codigo_Convenio"
         DataSource      =   "Ado_Contrato"
         Height          =   315
         Left            =   2520
         TabIndex        =   32
         Top             =   2780
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "Codigo_Convenio"
         BoundColumn     =   "Codigo_Convenio"
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
      Begin MSDataListLib.DataCombo DtcPry 
         Bindings        =   "Aux_Benef.frx":09CD
         DataField       =   "Pro_proyecto"
         DataSource      =   "Ado_Contrato"
         Height          =   315
         Left            =   2520
         TabIndex        =   33
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "Pro_proyecto"
         BoundColumn     =   "Pro_proyecto"
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
      Begin MSDataListLib.DataCombo DtcCargoDes 
         Bindings        =   "Aux_Benef.frx":09E2
         DataField       =   "codigo_cargo"
         DataSource      =   "Ado_Contrato"
         Height          =   315
         Left            =   4440
         TabIndex        =   34
         Top             =   3120
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "descripcion_cargo"
         BoundColumn     =   "codigo_cargo"
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
      Begin MSDataListLib.DataCombo DtcCargo 
         Bindings        =   "Aux_Benef.frx":09F9
         DataField       =   "codigo_cargo"
         DataSource      =   "Ado_Contrato"
         Height          =   315
         Left            =   6240
         TabIndex        =   35
         Top             =   2780
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "codigo_cargo"
         BoundColumn     =   "codigo_cargo"
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
      Begin MSDataListLib.DataCombo DtcFteDes 
         Bindings        =   "Aux_Benef.frx":0A10
         DataField       =   "Fte_codigo"
         DataSource      =   "Ado_Contrato"
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   2400
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "Fte_descripcion_larga"
         BoundColumn     =   "Fte_codigo"
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
      Begin MSDataListLib.DataCombo DTcFte 
         Bindings        =   "Aux_Benef.frx":0A28
         DataField       =   "Fte_codigo"
         DataSource      =   "Ado_Contrato"
         Height          =   315
         Left            =   2520
         TabIndex        =   37
         Top             =   2060
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         ListField       =   "Fte_codigo"
         BoundColumn     =   "Fte_codigo"
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
      Begin VB.Label lblARCH 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         DataField       =   "ARCHIVO"
         DataSource      =   "Ado_Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   7320
         TabIndex        =   44
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Proyecto                                                                                   Puesto que Ocupa "
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   43
         Top             =   3560
         Width           =   5745
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Organismo Financiador                                                             Cargo que Ocupa"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   42
         Top             =   2880
         Width           =   5625
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Código Contrato            Vigente       Aprobado       Monto Contrato             Nro.Reg."
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   41
         Top             =   280
         Width           =   5865
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "FONDO (Fuente Financiamiento)                                              Area Organizacional"
         Height          =   195
         Index           =   28
         Left            =   120
         TabIndex        =   40
         Top             =   2175
         Width           =   5805
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Objeto del Contrato . . ."
         Height          =   195
         Index           =   29
         Left            =   120
         TabIndex        =   39
         Top             =   1095
         Width           =   1635
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Firma Contrato:                                      Fecha de Inicio:                                     Fecha de Fin:"
         Height          =   195
         Index           =   31
         Left            =   120
         TabIndex        =   38
         Top             =   1725
         Width           =   7050
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800080&
         BorderWidth     =   2
         X1              =   4365
         X2              =   4365
         Y1              =   2040
         Y2              =   4305
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00800080&
         X1              =   0
         X2              =   8760
         Y1              =   2040
         Y2              =   2040
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   8895
      Begin VB.Frame Frame14 
         Caption         =   "IV. TOTAL BENEFICIOS SOCIALES"
         ForeColor       =   &H00C00000&
         Height          =   1005
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   8655
         Begin VB.ComboBox Combo7 
            DataField       =   "Forma_pago"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            ItemData        =   "Aux_Benef.frx":0A40
            Left            =   120
            List            =   "Aux_Benef.frx":0A4D
            TabIndex        =   14
            Text            =   "CHEQUE"
            Top             =   540
            Width           =   1620
         End
         Begin VB.TextBox Text14 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Monto_Total"
            DataSource      =   "Ado_Auxiliar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7080
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   540
            Width           =   1455
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Num_chq_cmpbte"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   1800
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   540
            Width           =   855
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Deducciones"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   5880
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   540
            Width           =   1095
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   2760
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   540
            Width           =   3015
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   $"Aux_Benef.frx":0A71
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   300
            Width           =   8460
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "III. TOTAL REMUNERACION PROMEDIO INDEMNIZABLE"
         ForeColor       =   &H00C00000&
         Height          =   1935
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8655
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Otros_Pagos"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   5880
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   1485
            Width           =   1455
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Prima_Legal"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            Left            =   4080
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   1485
            Width           =   1455
         End
         Begin VB.ComboBox Combo3 
            DataField       =   "Años"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            IntegralHeight  =   0   'False
            ItemData        =   "Aux_Benef.frx":0AF9
            Left            =   5640
            List            =   "Aux_Benef.frx":0B2A
            TabIndex        =   6
            Text            =   "0"
            Top             =   525
            Width           =   900
         End
         Begin VB.ComboBox Combo2 
            DataField       =   "Años"
            DataSource      =   "Ado_Auxiliar"
            Height          =   315
            IntegralHeight  =   0   'False
            ItemData        =   "Aux_Benef.frx":0B6A
            Left            =   3480
            List            =   "Aux_Benef.frx":0B92
            TabIndex        =   5
            Text            =   "0"
            Top             =   525
            Width           =   900
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Utimo"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   5640
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   900
            Width           =   1695
         End
         Begin VB.TextBox Text15 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Penul"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   3480
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   900
            Width           =   1695
         End
         Begin VB.TextBox Text16 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "OtroPago_Antep"
            DataSource      =   "Ado_Auxiliar"
            Height          =   285
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   900
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "Aux_benef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Set rs_Puesto_Org = New ADODB.Recordset
  rs_Puesto_Org.Open "select * from rc_PUESTO_organizacional  ", DB, adOpenKeyset, adLockOptimistic
  Set AdoPuestoOrg.Recordset = rs_Puesto_Org.DataSource
  DtcPuestoDes.BoundText = DtcPuesto.BoundText

  Set rs_UNIDAD = New ADODB.Recordset
  rs_UNIDAD.Open "select * from fc_unidad_ejecutora  ", DB, adOpenKeyset, adLockOptimistic
  Set AdoUnidad.Recordset = rs_UNIDAD.DataSource
  Dtc_descrip.BoundText = Dtc_codigo.BoundText
  
  Set rsfuente = New ADODB.Recordset
  rsfuente.Open "select * from fc_fuente_financiamiento WHERE fte_activo=1 ", DB, adOpenKeyset, adLockOptimistic
  Set AdoFuente.Recordset = rsfuente
  DtcFteDes.BoundText = DtcFte.BoundText
    
  Set rs_Org = New ADODB.Recordset
  rs_Org.Open "select * from fc_convenios  ", DB, adOpenKeyset, adLockOptimistic
  Set AdoOrg.Recordset = rs_Org.DataSource
  DtcOrgDes.BoundText = DtcOrg.BoundText
  
  Set rs_Pry = New ADODB.Recordset
  rs_Pry.Open "select * from fc_estructura_programatica  ", DB, adOpenKeyset, adLockOptimistic
  Set AdoPry.Recordset = rs_Pry.DataSource
  DtcPryDes.BoundText = DtcPry.BoundText
  
  Set rs_CARGO = New ADODB.Recordset
  rs_CARGO.Open "select * from RC_CARGOS  ", DB, adOpenKeyset, adLockOptimistic
  Set AdoCargo.Recordset = rs_CARGO.DataSource
  DtcCargoDes.BoundText = DtcCargo.BoundText


' PARA MOVECOPLETE
If swnuevo = "M" Then
    If rs_contrato!cod_est_contrato = "NO" Then
        TxtAprob.ForeColor = &H80&
        CmdAddCto.Visible = True
        CmdModCto.Visible = True
        CmdGrabaCto.Visible = False
        CmdAprCto.Visible = True
    Else
        TxtAprob.ForeColor = &H4000&
        CmdAddCto.Visible = True
        CmdModCto.Visible = False
        CmdGrabaCto.Visible = False
        CmdAprCto.Visible = False
    End If
  Else
    If rs_contrato!cod_est_contrato = "NO" Then
        TxtAprob.ForeColor = &H80&
        lblARCH.ForeColor = &H80&
    Else
        TxtAprob.ForeColor = &H4000&
        lblARCH.ForeColor = &H4000&
    End If
  End If
'FIN MOVECOMPLETE
End Sub

Private Sub Frame13_DragDrop(Source As Control, X As Single, y As Single)
 If lblARCH.Caption = "Cargar_Archivo" Then
    MsgBox ("No Existe el Archivo Asociado al Contrato, debe Cargarlo ...")
 Else
    If GlServidor <> GlMaquina Then      ' "-" Then
        e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(adoLista.Recordset!iniciales) & "_" & Trim(Ado_Contrato.Recordset!codigo_beneficiario) & "\CONTRATOS\" & Trim(Ado_Contrato.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    Else
        e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(adoLista.Recordset!iniciales) & "_" & Trim(Ado_Contrato.Recordset!codigo_beneficiario) & "\CONTRATOS\" & Trim(Ado_Contrato.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    End If
 End If
End Sub

Private Sub CmdGrabaCto_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
'  Call valida_campos
  If VAR_VAL = "OK" Then
    If GlSW = "ADD" Then
      rs_contrato!codigo_contrato = TxtCodigo.Text
      rs_contrato!codigo_beneficiario = adoLista.Recordset("codigo_Beneficiario") 'DtcBenef.Text
      rs_contrato!ges_gestion = glGestion
      rs_contrato!codigo_solicitud = rs_contrato.RecordCount
      
      Set rs_correlativo = New ADODB.Recordset
      rs_correlativo.Open "select * from ao_contrato_c WHERE codigo_beneficiario = '" & adoLista.Recordset("codigo_Beneficiario") & "'  ", DB, adOpenKeyset, adLockOptimistic
      If rs_correlativo.RecordCount > 0 Then
            rs_contrato!numero_consultoria = rs_correlativo.RecordCount
'            rs_correlativo!correlativo = rs_correlativo!correlativo + 1
'            rs_correlativo.Update
'            rs_M1!Numero_FA = rs_correlativo!correlativo
      Else
            rs_contrato!numero_consultoria = 1
      End If
      rs_contrato!ARCHIVO = "Cargar_Archivo"
      rs_contrato!ARCHIVO_NOMB = Trim(adoLista.Recordset("iniciales")) & "_Contrato_" & rs_contrato!numero_consultoria & ".pdf"
      TxtAprob.Text = "NO"
    End If
      rs_contrato!objeto_contrato = txtObjContrato.Text
      rs_contrato!codigo_puesto = DtcPuesto.Text
      rs_contrato!codigo_unidad = Dtc_codigo.Text
      rs_contrato!codigo_convenio = DtcOrg.Text
      rs_contrato!pro_proyecto = DtcPry.Text
      rs_contrato!fechas_confirmado = txtEstado
      rs_contrato!cod_est_contrato = TxtAprob
      rs_contrato!fecha_firma = DTPFFirma.Value
      rs_contrato!fecha_inicio = DTPFInicio.Value
      rs_contrato!fecha_fin = DTPFFin.Value
      rs_contrato!monto_totalbs = TxtBs.Text
      If GlTipoCambioOficial > 0 Then
        rs_contrato!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
      Else
        GlTipoCambioOficial = 7.05
        rs_contrato!monto_totalus = CDbl(TxtBs.Text) / GlTipoCambioOficial
      End If
      rs_contrato!observacion_contrato = "-"
      rs_contrato!establece_multas = "N"
      rs_contrato!cod_forma_inicio = "1"
      rs_contrato!tiempo_num = 0
      rs_contrato!tiempo_dmy = "-"
      rs_contrato!tipo_moneda = "Bs"
      rs_contrato!tc_us = GlTipoCambioOficial
      
      rs_contrato!org_codigo = "111"
      rs_contrato!porc_orgfin = 100
      rs_contrato!porc_contra = 0
      'rs_contrato!fechas_confirmado = "N"
      rs_contrato!hora_registro = "8:00"
      rs_contrato!fecha_registro = Date
      rs_contrato!usr_usuario = "ADMIN" 'GlUsuario
      rs_contrato.Update    'Batch adAffectAll
      
'      mbDataChanged = False
      CmdAddCto.Visible = True
      CmdModCto.Visible = True
      CmdGrabaCto.Visible = False
      CmdAprCto.Visible = True
      TxtAprob.Enabled = True
      Fra_ABM.Enabled = False
      DtG_Auxiliar.Enabled = False
      GlSW = " "

  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub


Private Sub Dtc_codigo_Click(Area As Integer)
    Dtc_descrip.BoundText = Dtc_codigo.BoundText
End Sub

Private Sub Dtc_descrip_Click(Area As Integer)
    Dtc_codigo.BoundText = Dtc_descrip.BoundText
End Sub

Private Sub DtcCargo_Click(Area As Integer)
    DtcCargoDes.BoundText = DtcCargo.BoundText
End Sub

Private Sub DtcCargoDes_Click(Area As Integer)
    DtcCargo.BoundText = DtcCargoDes.BoundText
End Sub

Private Sub DTcFte_Click(Area As Integer)
   DtcFteDes.BoundText = DtcFte.BoundText
   Call pOrganismo(DtcFteDes.BoundText)
End Sub

Private Sub DtcFteDes_Click(Area As Integer)
    DtcFte.BoundText = DtcFteDes.BoundText
    Call pOrganismo(DtcFte.BoundText)
End Sub

Private Sub pOrganismo(CodFuente As String)
   Dim strConsultaF As String
   strConsultaF = "select * from fc_convenios where fte_codigo='" & CodFuente & "'"
   Set DtcOrg.RowSource = Nothing
   Set DtcOrg.RowSource = DB.Execute(strConsultaF, , adCmdText)
   DtcOrg.ReFill
   DtcOrg.BoundText = Empty
   Set DtcOrgDes.RowSource = Nothing
   Set DtcOrgDes.RowSource = DB.Execute(strConsultaF, , adCmdText)
   DtcOrgDes.ReFill
   DtcOrgDes.BoundText = Empty
End Sub

Private Sub DtcOrg_Click(Area As Integer)
    DtcOrgDes.BoundText = DtcOrg.BoundText
    Call pCat(DtcOrgDes.BoundText)
End Sub

Private Sub DtcOrgDes_Click(Area As Integer)
    DtcOrg.BoundText = DtcOrgDes.BoundText
    Call pCat(DtcOrg.BoundText)
End Sub

Private Sub pCat(CodOrganismo As String)
   Dim strConsulta As String
   
   strConsulta = "select * from fc_estructura_programatica where codigo_convenio='" & CodOrganismo & "'"
   
   Set DtcPry.RowSource = Nothing
   Set DtcPry.RowSource = DB.Execute(strConsulta, , adCmdText)
   DtcPry.ReFill
   DtcPry.BoundText = Empty
   
   Set DtcPryDes.RowSource = Nothing
   Set DtcPryDes.RowSource = DB.Execute(strConsulta, , adCmdText)
   DtcPryDes.ReFill
   DtcPryDes.BoundText = Empty

End Sub

Private Sub DtcPry_Click(Area As Integer)
    DtcPryDes.BoundText = DtcPry.BoundText
End Sub

Private Sub DtcPryDes_Click(Area As Integer)
    DtcPry.BoundText = DtcPryDes.BoundText
End Sub


Private Sub DtcPuesto_Click(Area As Integer)
    DtcPuestoDes.BoundText = DtcPuesto.BoundText
End Sub

Private Sub DtcPuestoDes_Click(Area As Integer)
    DtcPuesto.BoundText = DtcPuestoDes.BoundText
End Sub

Private Sub Image2_Click()
 If lblARCH.Caption = "Cargar_Archivo" Then
    MsgBox ("No Existe el Archivo Asociado al Contrato, debe Cargarlo ...")
 Else
    If GlServidor <> GlMaquina Then      ' "-" Then
        e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(adoLista.Recordset!iniciales) & "_" & Trim(Ado_Contrato.Recordset!codigo_beneficiario) & "\CONTRATOS\" & Trim(Ado_Contrato.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    Else
        e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(adoLista.Recordset!iniciales) & "_" & Trim(Ado_Contrato.Recordset!codigo_beneficiario) & "\CONTRATOS\" & Trim(Ado_Contrato.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
    End If
 End If

End Sub

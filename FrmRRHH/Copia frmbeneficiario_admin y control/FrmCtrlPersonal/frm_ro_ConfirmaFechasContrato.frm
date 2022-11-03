VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frm_ro_ConfirmaFechasContrato 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comfirma fechas y datos de consultoría"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   840
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Sale del formulario"
      Top             =   7920
      Width           =   885
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos de Consultoría"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   1080
      TabIndex        =   51
      Top             =   720
      Width           =   8655
      Begin VB.TextBox txtNroConsultoria 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   69
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtDesConsultoria 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   65
         Top             =   1440
         Width           =   6855
      End
      Begin VB.TextBox txtFileConsultoria 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   53
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txtCodPrism 
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   52
         Top             =   2040
         Width           =   2895
      End
      Begin MSDataListLib.DataCombo cboEstadoProceso 
         Height          =   315
         Left            =   5640
         TabIndex        =   67
         Top             =   1080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         BackColor       =   14737632
         ForeColor       =   4210752
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
      Begin VB.Label Label4 
         Caption         =   "Nro. Consultoría:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado Proceso:"
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
         Height          =   255
         Left            =   3720
         TabIndex        =   68
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción consultoría:"
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
         Height          =   495
         Left            =   120
         TabIndex        =   66
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "No. Solic.:"
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
         Height          =   255
         Left            =   6480
         TabIndex        =   64
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "Uni. Sol.:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblCodUnidad 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   62
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblUniSol 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   61
         Top             =   720
         Width           =   6255
      End
      Begin VB.Label Label30 
         Caption         =   "Gestión:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblGestion 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   59
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblFormulario 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   58
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblNroSol 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   57
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Form.:"
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
         Height          =   255
         Left            =   1920
         TabIndex        =   56
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label38 
         Caption         =   "File Consultoria:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label35 
         Caption         =   "Código PRISM:"
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
         Height          =   255
         Left            =   4320
         TabIndex        =   54
         Top             =   2040
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   0
      TabIndex        =   42
      Top             =   840
      Width           =   1020
      Begin VB.CommandButton cmdEditarAdj 
         Caption         =   "Modificar"
         Height          =   840
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Modifica datos de consultoría"
         Top             =   240
         Width           =   765
      End
      Begin VB.CommandButton cmdGuardaAdj 
         Caption         =   "Grabar"
         Height          =   840
         Left            =   120
         MousePointer    =   4  'Icon
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   240
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmdCancelaAdj 
         Caption         =   "Cancelar"
         Height          =   840
         Left            =   120
         MousePointer    =   4  'Icon
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   1320
         Visible         =   0   'False
         Width           =   765
      End
   End
   Begin VB.Frame fraContrato 
      Caption         =   "Datos de Contrato Original"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   1080
      TabIndex        =   2
      Top             =   3120
      Width           =   8655
      Begin VB.TextBox tdnOrgContraBS 
         Height          =   285
         Left            =   7320
         TabIndex        =   80
         Text            =   "Text6"
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox tdnOrgContraUS 
         Height          =   285
         Left            =   6120
         TabIndex        =   79
         Text            =   "Text5"
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox tdnOrgBaseBS 
         Height          =   285
         Left            =   7320
         TabIndex        =   78
         Text            =   "Text4"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox tdnOrgBaseUS 
         Height          =   285
         Left            =   6120
         TabIndex        =   77
         Text            =   "Text3"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox tdnMontoBS 
         Height          =   285
         Left            =   7320
         TabIndex        =   76
         Text            =   "Text2"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox tdnMontoUS 
         Height          =   285
         Left            =   6120
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox tdnTcUS 
         Height          =   285
         Left            =   4800
         TabIndex        =   74
         Text            =   "Text1"
         Top             =   3360
         Width           =   615
      End
      Begin MSComCtl2.DTPicker tddFFirma 
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CalendarBackColor=   12640511
         Format          =   92143617
         CurrentDate     =   39638
      End
      Begin VB.Frame Frame4 
         Height          =   420
         Left            =   240
         TabIndex        =   7
         Top             =   5520
         Width           =   8295
         Begin VB.Label lblNroAddendum 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nro. registros:"
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
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   3255
         End
         Begin VB.Label Label21 
            Caption         =   "Totales:"
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
            Left            =   3600
            TabIndex        =   9
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblTotalUSAdd 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "$US:"
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
            Left            =   4440
            TabIndex        =   10
            Top             =   120
            Width           =   1830
         End
         Begin VB.Label lblTotalBSAdd 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Bs.:"
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
            Left            =   6360
            TabIndex        =   11
            Top             =   120
            Width           =   1830
         End
      End
      Begin VB.TextBox txtObjetoContrato 
         BackColor       =   &H00C0E0FF&
         Height          =   525
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Top             =   600
         Width           =   7335
      End
      Begin VB.TextBox txtCodContrato 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   45
         Top             =   240
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         Caption         =   "Plazo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   3120
         TabIndex        =   27
         Top             =   1080
         Width           =   5415
         Begin VB.CheckBox chkConfirma 
            Caption         =   "CONFIRMAR FECHAS DE INICIO Y FINALIZACIÓN"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   35
            Top             =   1320
            Width           =   4695
         End
         Begin VB.Frame fraTiempo 
            Caption         =   "Tiempo: "
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   2520
            TabIndex        =   28
            Top             =   240
            Width           =   2535
            Begin VB.TextBox txtTiempo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   240
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   32
               Top             =   360
               Width           =   735
            End
            Begin VB.OptionButton optTiempo 
               Caption         =   "Año (s)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   1200
               TabIndex        =   31
               Top             =   720
               Width           =   1095
            End
            Begin VB.OptionButton optTiempo 
               Caption         =   "Mes (es)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   1200
               TabIndex        =   30
               Top             =   480
               Width           =   1095
            End
            Begin VB.OptionButton optTiempo 
               Caption         =   "Día (s)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   1200
               TabIndex        =   29
               Top             =   240
               Width           =   1095
            End
         End
         Begin MSComCtl2.DTPicker dtpFInicio 
            Height          =   315
            Left            =   960
            TabIndex        =   36
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   92143617
            CurrentDate     =   36882
         End
         Begin MSComCtl2.DTPicker dtpFFin 
            Height          =   315
            Left            =   960
            TabIndex        =   37
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   92143617
            CurrentDate     =   36882
         End
         Begin VB.Label Label74 
            Alignment       =   1  'Right Justify
            Caption         =   "F. Inicio:"
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
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "F. Fin:"
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
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.OptionButton optMultas 
         Caption         =   "SI establece multas"
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
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Top             =   1920
         Width           =   2055
      End
      Begin VB.OptionButton optMultas 
         Caption         =   "NO establece multas"
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
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   2280
         Value           =   -1  'True
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo cboTipoMoneda 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   3360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         BackColor       =   12640511
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboFormaInicio 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   2880
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         BackColor       =   12640511
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboEstadoCont 
         Height          =   315
         Left            =   5760
         TabIndex        =   46
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         BackColor       =   12640511
         ForeColor       =   4210752
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
      Begin TrueOleDBGrid60.TDBGrid grdAddendum 
         Height          =   1320
         Left            =   240
         OleObjectBlob   =   "frm_ro_ConfirmaFechasContrato.frx":0000
         TabIndex        =   81
         Top             =   4320
         Width           =   8295
      End
      Begin VB.Label Label73 
         Caption         =   "Objeto de Contrato:"
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
         Height          =   495
         Left            =   120
         TabIndex        =   50
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado Contrato:"
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
         Height          =   255
         Left            =   4080
         TabIndex        =   48
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label75 
         Caption         =   "Código:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto Bs."
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
         Height          =   255
         Left            =   7320
         TabIndex        =   26
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto $US."
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
         Height          =   255
         Left            =   6120
         TabIndex        =   25
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
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
         Height          =   255
         Left            =   5520
         TabIndex        =   24
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label lblPorcOrgContra 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   5520
         TabIndex        =   23
         Top             =   4005
         Width           =   615
      End
      Begin VB.Label lblDesOrgContra 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Organismo contraparte: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   1560
         TabIndex        =   22
         Top             =   4005
         Width           =   3975
      End
      Begin VB.Label lblPorcOrgBase 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   5520
         TabIndex        =   21
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblDesOrgBase 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Organismo base: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   1560
         TabIndex        =   20
         Top             =   3720
         Width           =   3975
      End
      Begin VB.Label lblPorcTotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5520
         TabIndex        =   19
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "TOTAL DE CONTRATO ORIGINAL"
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
         Height          =   180
         Left            =   5520
         TabIndex        =   18
         Top             =   2880
         Width           =   3015
      End
      Begin VB.Label Label37 
         Caption         =   "Moneda:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label72 
         Alignment       =   2  'Center
         Caption         =   "TC $US:"
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
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Forma de Inicio:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "Firma de Contrato:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.Frame fraToolBarGuarda 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Width           =   1020
      Begin VB.CommandButton cmdCancela 
         Caption         =   "Cancelar"
         Height          =   840
         Left            =   120
         MousePointer    =   4  'Icon
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   1920
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Modificar"
         Height          =   840
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Modifica la cofirmación de las fechas de consultoría"
         Top             =   480
         Width           =   765
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Grabar"
         Height          =   840
         Left            =   120
         MousePointer    =   4  'Icon
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   480
         Visible         =   0   'False
         Width           =   765
      End
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "CONFIRMACIÓN DATOS DE CONSULRORÍA Y CONTRATO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Beneficiario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1080
      TabIndex        =   39
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblBeneficiario 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Left            =   2640
      TabIndex        =   38
      Top             =   360
      Width           =   6975
   End
End
Attribute VB_Name = "frm_ro_ConfirmaFechasContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQLs As String ' usado para la elaboración de los querys
Dim rs_AdjCont As ADODB.Recordset ' usado para la carga de los datos
Dim CodBenef As String

Private Sub chkConfirma_Click()
    If chkConfirma.Value = 1 Then
        dtpFInicio.Enabled = False
        dtpFFin.Enabled = False
      Else
        dtpFInicio.Enabled = True
        dtpFFin.Enabled = True
    End If
End Sub

Private Sub cmdCancela_Click()
    If vbYes = MsgBox("Desea mostrar los valores originales, perdiendo cualquier modificación realizada?", vbDefaultButton2 + vbYesNo + vbQuestion, "Aviso") Then
        Call pl_RefrescaContratoAdj
        pl_HabilitaOpcion (False)
      Else
        chkConfirma.SetFocus
    End If

End Sub

Private Sub cmdCancelaAdj_Click()
    If vbYes = MsgBox("Desea mostrar los valores originales, perdiendo cualquier modificación realizada?", vbDefaultButton2 + vbYesNo + vbQuestion, "Aviso") Then
        Call pl_RefrescaContratoAdj
        cmdGuardaAdj.Visible = False
        cmdCancelaAdj.Visible = False
        cmdEditarAdj.Visible = True
        txtFileConsultoria.Locked = True
        txtCodPrism.Locked = True
        If IsNull(rs_AdjCont!id_contrato) Then
            cmdEditar.Enabled = False
          Else
            cmdEditar.Enabled = True
        End If
      Else
        txtFileConsultoria.SetFocus
    End If
   
End Sub

Private Sub cmdEditar_Click()
    pl_HabilitaOpcion (True)
    chkConfirma.SetFocus
End Sub

Private Sub cmdEditarAdj_Click()
    cmdGuardaAdj.Visible = True
    cmdCancelaAdj.Visible = True
    cmdEditarAdj.Visible = False
    cmdEditar.Enabled = False
    txtFileConsultoria.Locked = False
    txtCodPrism.Locked = False
    txtFileConsultoria.SetFocus
End Sub

Private Sub cmdGuardaAdj_Click()
    Dim sw As Boolean
    On Error GoTo EtiqError
    sw = True
    
    If Len(Trim(txtFileConsultoria.Text)) = 0 Then
        MsgBox "File consultoria no es valido." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        txtFileConsultoria.SetFocus
        sw = False
        Exit Sub
    End If
    
    If Len(Trim(txtCodPrism.Text)) = 0 Then
        MsgBox "El codigo PRISM no es valido." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        txtCodPrism.SetFocus
        sw = False
        Exit Sub
    End If
    
    If sw = True Then
        Screen.MousePointer = vbHourglass
        If vbYes = MsgBox("Se confirmará los datos de FILE y código PRISM de la consultoría correspondiente a [" & lblBeneficiario.Caption & "]", vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación") Then
            SQLs = "UPDATE ao_adjudica_c SET file_consultoria = '" & pg_ReemplazaCarater(pg_QuitaEspBlanco(txtFileConsultoria.Text), Chr(34), Chr(39)) & "', codigo_prism = '" & pg_ReemplazaCarater(pg_QuitaEspBlanco(txtCodPrism.Text), Chr(34), Chr(39)) & "' "
            SQLs = SQLs & " WHERE ges_gestion = '" & rs_AdjCont!ges_gestion & "' "
            SQLs = SQLs & "AND numero_consultoria = " & rs_AdjCont!numero_consultoria & " and codigo_beneficiario ='" & CodBenef & "'"
            'JQ QR
            'De.dbo_apGeneralSearching SQLs
            txtFileConsultoria.Locked = True
            txtCodPrism.Locked = True
            cmdEditarAdj.Visible = True
            cmdGuardaAdj.Visible = False
            cmdCancelaAdj.Visible = False
            
            If IsNull(rs_AdjCont!id_contrato) Then
                cmdEditar.Enabled = False
              Else
                cmdEditar.Enabled = True
            End If

          Else
            txtFileConsultoria.SetFocus
        End If
        Screen.MousePointer = vbDefault
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo EtiqError
    
    If fl_VerificaFechas Then ' verificamos si la información está correcta antes de actualizar la BD
        Screen.MousePointer = vbHourglass
        If vbYes = MsgBox("Se confirmará las fechas de inicio y de finalización de la consultoría correspondiente a [" & lblBeneficiario.Caption & "]", vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmación") Then
            SQLs = "UPDATE ao_contrato_c "
            SQLs = SQLs & "SET fechas_confirmado = 'S', fecha_inicio = '" & dtpFInicio.Value & "', fecha_fin = '" & dtpFFin.Value & "', tiempo_num = " & Val(txtTiempo.Text) & ", tiempo_dmy = '" & IIf(optTiempo(0).Value = True, "dd", IIf(optTiempo(1).Value = True, "mm", "yy")) & "'"
            SQLs = SQLs & " WHERE ges_gestion = '" & rs_AdjCont!ges_gestion & "' "
            SQLs = SQLs & "AND numero_consultoria = " & rs_AdjCont!numero_consultoria & " AND id_contrato = " & rs_AdjCont!id_contrato & " AND codigo_beneficiario ='" & CodBenef & "'"
            'JQ QR
            'De.dbo_apGeneralSearching SQLs
            pl_HabilitaOpcion (False)
        End If
        Call pl_RefrescaContratoAdj
        Screen.MousePointer = vbDefault
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub dtpFFin_Change()
    Select Case True
      Case optTiempo(0).Value
        Call optTiempo_Click(0)
      Case optTiempo(1).Value
        Call optTiempo_Click(1)
      Case optTiempo(2).Value
        Call optTiempo_Click(2)
    End Select

End Sub

Private Sub dtpFFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub dtpFinicio_Change()
    Select Case True
      Case optTiempo(0).Value
        Call optTiempo_Click(0)
      Case optTiempo(1).Value
        Call optTiempo_Click(1)
      Case optTiempo(2).Value
        Call optTiempo_Click(2)
    End Select

End Sub


Private Sub dtpFinicio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    lblBeneficiario.Caption = frm_ro_LiquidaMain.lblEstadoBeneficiario.Caption   ' des de beneficario
    CodBenef = frm_ro_LiquidaMain.lblEstadoBeneficiario.Tag ' codigo de beneficario
    
    Call pl_Llena_Combos_Base
    Call pl_RefrescaContratoAdj
    
	Call SeguridadSet(Me)
End Sub

Private Sub pl_RefrescaContratoAdj()
    Dim MontoTotalBs As Double ' usado para guardar el monto total BS limite de adjudicacion
    Dim MontoTotalUS As Double ' usado para guardar el monto total US limite de adjudicacion
    Dim MontoBsAdjudicado As Double ' usado para guardar el monto BS subtotal adjudicado
    Dim MontoUSAdjudicado As Double ' usado para guardar el monto US subtotal adjudicado
    Dim DesOrgBase As String ' usado para guardar la descripcion del organismo base
    Dim PorcOrgBase As Double ' usado para guardar el el porcentaje del organismo base
    Dim DesOrgContra As String ' usado para guardar la descripcion del organismo contraparte
    Dim PorcOrgContra As Double ' usado para guardar el el porcentaje del organismo contraparte
    
    On Error GoTo EtiqError
    
    ' obtiene datos de consultoria y contrato
    SQLs = "SELECT  ao_contrato_c.objeto_contrato, ao_contrato_c.observacion_contrato, ao_contrato_c.establece_multas, ao_contrato_c.cod_forma_inicio,"
    SQLs = SQLs & "ao_contrato_c.fecha_inicio, ao_contrato_c.fecha_fin, ao_contrato_c.tiempo_num, ao_contrato_c.tiempo_dmy, ao_contrato_c.tipo_moneda,"
    SQLs = SQLs & "ao_contrato_c.tc_us, ao_contrato_c.monto_totalUS, ao_contrato_c.monto_totalBS, ao_contrato_c.cod_est_contrato, ao_contrato_c.fechas_confirmado,"
    SQLs = SQLs & "ao_contrato_c.id_contrato, ao_no_objecion_c.ges_gestion, ao_no_objecion_c.codigo_solicitud, ao_no_objecion_c.codigo_unidad,"
    SQLs = SQLs & "ao_no_objecion_c.numero_consultoria, ao_no_objecion_c.formulario, ac_Tipo_Tramite.Denominacion_Tipo, ao_no_objecion_c.des_consultoria,"
    SQLs = SQLs & "ao_no_objecion_c.cod_est_no_objecion , ao_adjudica_c.codigo_prism, ao_adjudica_c.file_consultoria, ao_contrato_c.codigo_contrato, ao_contrato_c.fecha_firma, fc_unidad_ejecutora.Uni_descripcion_larga "
    SQLs = SQLs & "FROM ao_adjudica_c INNER JOIN ao_no_objecion_c ON ao_adjudica_c.numero_consultoria = ao_no_objecion_c.numero_consultoria AND ao_adjudica_c.ges_gestion = ao_no_objecion_c.ges_gestion AND ao_adjudica_c.codigo_solicitud = ao_no_objecion_c.codigo_solicitud AND "
    SQLs = SQLs & "ao_adjudica_c.codigo_unidad = ao_no_objecion_c.codigo_unidad INNER JOIN ac_Tipo_Tramite ON ao_no_objecion_c.formulario = ac_Tipo_Tramite.Tipo_Formulario LEFT OUTER JOIN fc_unidad_ejecutora ON ao_no_objecion_c.codigo_unidad = fc_unidad_ejecutora.codigo_unidad LEFT OUTER JOIN "
    SQLs = SQLs & "ao_contrato_c ON ao_adjudica_c.ges_gestion = ao_contrato_c.ges_gestion AND ao_adjudica_c.codigo_unidad = ao_contrato_c.codigo_unidad AND ao_adjudica_c.codigo_solicitud = ao_contrato_c.codigo_solicitud AND ao_adjudica_c.numero_consultoria = ao_contrato_c.numero_consultoria AND "
    SQLs = SQLs & "ao_adjudica_c.codigo_beneficiario = ao_contrato_c.codigo_beneficiario "
    SQLs = SQLs & "WHERE ao_adjudica_c.gp_ges_gestion = '" & frm_ro_LiquidaMain.lblGestion.Caption & "' AND ao_adjudica_c.gp_codigo_unidad = '" & frm_ro_LiquidaMain.lblCodUniSol & "' AND ao_adjudica_c.gp_codigo_grupo = " & Val(frm_ro_LiquidaMain.lblCodGrupo.Caption) & " AND ao_adjudica_c.codigo_beneficiario = '" & CodBenef & "'"
        
    Set rs_AdjCont = New ADODB.Recordset
    rs_AdjCont.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rs_AdjCont.RecordCount = 0 Then
    
        cmdEditarAdj.Enabled = False
        cmdEditar.Enabled = False
        
        lblBeneficiario.Caption = ""
        lblGestion.Caption = ""
        lblFormulario.Caption = ""
        lblNroSol.Caption = ""
        lblCodUnidad.Caption = ""
        lblUniSol.Caption = ""
        txtNroConsultoria.Text = ""
        txtDesConsultoria.Text = ""
        cboEstadoProceso.BoundText = ""
        
        txtFileConsultoria.Text = ""
        txtCodPrism.Text = ""
        
        txtCodContrato.Text = ""
        txtObjetoContrato.Text = ""
        dtpFInicio.Value = Date
        dtpFFin.Value = Date
        tddFFirma.Value = Date
        txtTiempo.Text = ""

        cboFormaInicio.BoundText = ""
        cboTipoMoneda.BoundText = ""
        tdnTcUS.Text = 0
        tdnMontoUS.Text = 0
        tdnMontoBS.Text = 0
        cboEstadoCont.BoundText = ""
        
        lblDesOrgBase.Caption = "Organismo externo: "
        lblPorcOrgBase.Caption = "0.00"
    
        lblDesOrgContra.Caption = "Contraparte Nacional: "
        lblPorcOrgContra.Caption = "0.00"
        lblPorcTotal.Caption = "0.00"
        
        tdnOrgBaseUS.Text = 0
        tdnOrgBaseBS.Text = 0
        
        tdnOrgContraUS.Text = 0
        tdnOrgContraBS.Text = 0
      
      Else
        lblGestion.Caption = rs_AdjCont!ges_gestion & ""
        lblFormulario.Caption = rs_AdjCont!formulario & " - " & rs_AdjCont!Denominacion_Tipo & ""
        lblNroSol.Caption = rs_AdjCont!codigo_solicitud & ""
        lblCodUnidad.Caption = rs_AdjCont!codigo_unidad & ""
        lblUniSol.Caption = rs_AdjCont!Uni_descripcion_larga & ""
        txtNroConsultoria.Text = rs_AdjCont!numero_consultoria & ""
        txtDesConsultoria.Text = rs_AdjCont!des_consultoria & ""
        cboEstadoProceso.BoundText = rs_AdjCont!cod_est_no_objecion & ""
        
        txtFileConsultoria.Text = rs_AdjCont!file_consultoria & ""
        txtCodPrism.Text = rs_AdjCont!codigo_prism & ""
        
        If Not IsNull(rs_AdjCont!id_contrato) Then ' determina si existe registro de contrato
            fraContrato.Enabled = True
            
            txtCodContrato.Text = rs_AdjCont!codigo_contrato & ""
            txtObjetoContrato.Text = rs_AdjCont!objeto_contrato & ""
                    
            dtpFInicio.Value = IIf(IsNull(rs_AdjCont!fecha_inicio), Date, rs_AdjCont!fecha_inicio)
            dtpFFin.Value = IIf(IsNull(rs_AdjCont!fecha_fin), Date, rs_AdjCont!fecha_fin)
            tddFFirma.Value = IIf(IsNull(rs_AdjCont!fecha_firma), Date, rs_AdjCont!fecha_firma)     'JQA JUL/2008
            'tddFFirma.Text = IIf(IsNull(rs_AdjCont!fecha_firma), "__/__/____", rs_AdjCont!fecha_firma)
            txtTiempo.Text = rs_AdjCont!tiempo_num & ""
    
            Select Case rs_AdjCont!tiempo_dmy & ""
              Case "dd"
                optTiempo(0).Value = True
              Case "mm"
                optTiempo(1).Value = True
              Case "yy"
                optTiempo(2).Value = True
            End Select
            
            Select Case rs_AdjCont!establece_multas & ""
              Case "S"
                optMultas(0).Value = True
              Case "N"
                optMultas(1).Value = True
            End Select
    
            cboTipoMoneda.BoundText = rs_AdjCont!tipo_moneda & ""
            tdnTcUS.Text = IIf(IsNull(rs_AdjCont!tc_us), 0, rs_AdjCont!tc_us)
            tdnMontoUS.Text = IIf(IsNull(rs_AdjCont!monto_totalUS), 0, rs_AdjCont!monto_totalUS)
            tdnMontoBS.Text = IIf(IsNull(rs_AdjCont!monto_totalBS), 0, rs_AdjCont!monto_totalBS)
            cboEstadoCont.BoundText = rs_AdjCont!cod_est_contrato & ""
            cboFormaInicio.BoundText = rs_AdjCont!cod_forma_inicio & ""
            
            If rs_AdjCont!fechas_confirmado & "" = "S" Then
                chkConfirma.Value = 1
              Else
                chkConfirma.Value = 0
            End If
            
            ' calcula montos adjudicados y porcentajes segun organismo
            'JQ QR
            'De.dbo_ap_SumaMontosLimiteBen rs_AdjCont!ges_gestion, rs_AdjCont!codigo_unidad, rs_AdjCont!codigo_solicitud, rs_AdjCont!numero_consultoria, MontoTotalBs, MontoTotalUS, MontoBsAdjudicado, MontoUSAdjudicado, DesOrgBase, PorcOrgBase, DesOrgContra, PorcOrgContra
            
            lblDesOrgBase.Caption = DesOrgBase & ": " ' organismo externo
            lblPorcOrgBase.Caption = Format(PorcOrgBase, "######0.00")
            
            If PorcOrgContra > 0 Then ' si tiene financiamiento de organismo contraparte (nacional)
                lblDesOrgContra.Caption = DesOrgContra & ": "
                lblPorcOrgContra.Caption = Format(PorcOrgContra, "######0.00")
              Else
                lblDesOrgContra.Caption = "Contraparte Nacional: "
                lblPorcOrgContra.Caption = Format(PorcOrgContra, "######0.00")
            End If
            lblPorcTotal.Caption = PorcOrgBase + PorcOrgContra
            
            ' calcula porcentajes por organismo de financiamiento
            tdnOrgBaseUS.Text = tdnMontoUS.Text * Val(lblPorcOrgBase.Caption) / 100
            tdnOrgBaseBS.Text = tdnMontoBS.Text * Val(lblPorcOrgBase.Caption) / 100
            
            tdnOrgContraUS.Text = tdnMontoUS.Text * Val(lblPorcOrgContra.Caption) / 100
            tdnOrgContraBS.Text = tdnMontoBS.Text * Val(lblPorcOrgContra.Caption) / 100
          
            ' obtiene datos de addendums al contrato
            SQLs = "SELECT ao_contrato_addendum_c.codigo_solicitud, ac_tipo_ord_cambio.des_corta_ord_cambio, ao_contrato_addendum_c.fecha_inicio, ao_contrato_addendum_c.fecha_fin, ao_contrato_addendum_c.tipo_moneda, ao_contrato_addendum_c.tc_us, "
            SQLs = SQLs & "ao_contrato_addendum_c.monto_totalUS, ao_contrato_addendum_c.monto_totalBS, CAST(ao_contrato_addendum_c.tiempo_num AS VARCHAR(10)) + ' '+ case when ao_contrato_addendum_c.tiempo_dmy = 'dd' then 'Días' else case when ao_contrato_addendum_c.tiempo_dmy = 'mm' then 'Meses' else 'Años' end end AS tiempo, "
            SQLs = SQLs & "ao_contrato_addendum_c.justificacion, ao_contrato_addendum_c.fecha_firma, ao_contrato_addendum_c.tiempo_num, ao_contrato_addendum_c.tiempo_dmy, ao_contrato_addendum_c.cod_est_contrato, ao_contrato_addendum_c.id_contrato, ao_contrato_addendum_c.id_addendum, ao_solicitud.cod_tipo_ord_cambio "
            SQLs = SQLs & "FROM ao_contrato_addendum_c INNER JOIN ao_Solicitud ON ao_contrato_addendum_c.ges_gestion = ao_Solicitud.Ges_Gestion AND ao_contrato_addendum_c.codigo_unidad = ao_Solicitud.codigo_unidad AND "
            SQLs = SQLs & "ao_contrato_addendum_c.codigo_solicitud_oc = ao_Solicitud.codigo_solicitud AND ao_contrato_addendum_c.codigo_beneficiario = ao_Solicitud.CI_aprueba INNER JOIN ac_tipo_ord_cambio ON ao_Solicitud.cod_tipo_ord_cambio = ac_tipo_ord_cambio.cod_tipo_ord_cambio "
            SQLs = SQLs & "WHERE ao_contrato_addendum_c.cod_est_contrato not IN(19,29,39) "
            SQLs = SQLs & " AND ao_contrato_addendum_c.id_contrato =" & IIf(IsNull(rs_AdjCont!id_contrato), 0, rs_AdjCont!id_contrato)
            SQLs = SQLs & " ORDER BY ao_contrato_addendum_c.codigo_solicitud DESC"
            
            Set rstTemp = New ADODB.Recordset
            rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
            
            Set grdAddendum.DataSource = rstTemp
            lblNroAddendum.Caption = "Nro. de addendums: " & rstTemp.RecordCount
            Call pl_PersonalizaGridAddendum
            
            SQLs = "SELECT 'monto_totalUS' = sum(ao_contrato_addendum_c.monto_totalUS), 'monto_totalBS' = sum(ao_contrato_addendum_c.monto_totalBS) "
            SQLs = SQLs & "FROM ao_contrato_addendum_c INNER JOIN ao_Solicitud ON ao_contrato_addendum_c.ges_gestion = ao_Solicitud.Ges_Gestion AND ao_contrato_addendum_c.codigo_unidad = ao_Solicitud.codigo_unidad AND "
            SQLs = SQLs & "ao_contrato_addendum_c.codigo_solicitud_oc = ao_Solicitud.codigo_solicitud AND ao_contrato_addendum_c.codigo_beneficiario = ao_Solicitud.CI_aprueba "
            SQLs = SQLs & "WHERE ao_contrato_addendum_c.cod_est_contrato not IN(19,29,39) AND "
            SQLs = SQLs & "ao_contrato_addendum_c.id_contrato =" & IIf(IsNull(rs_AdjCont!id_contrato), 0, rs_AdjCont!id_contrato)
            
            Set rstTemp = New ADODB.Recordset
            rstTemp.Open SQLs, db, adOpenStatic, adLockReadOnly
            If rstTemp.RecordCount > 0 Then
                lblTotalUSAdd.Caption = IIf(IsNull(rstTemp!monto_totalUS), 0, rstTemp!monto_totalUS) & " $US"
                lblTotalBSAdd.Caption = IIf(IsNull(rstTemp!monto_totalBS), 0, rstTemp!monto_totalBS) & " Bs"
            End If

          Else
            
            MsgBox "No existe registro de contrato." & Chr(13) & "Corrija el error e intente procesar nuevamente.", vbInformation, "Aviso"
            fraContrato.Enabled = False
            cmdEditar.Enabled = False
            
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub

EtiqError:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Sub

Private Sub pl_PersonalizaGridAddendum()
    'TITULO:                Procedimiento pl_PersonalizaGridAddendum
    'PROPOSITO:             Personalizar los captions, anchos, etc. del grid
    'EJEMPLO DE LLAMADA:    call pl_PersonalizaGridAddendum
        
    Dim i As Integer
    
    ' define ancho de columnas y titulo de la cabecera
    grdAddendum.Columns(0).Width = 1000 ' codigo solicitud OC
    grdAddendum.Columns(0).Caption = "Cod. Sol.OC"
    grdAddendum.Columns(1).Width = 1700 ' tipo orden cambio
    grdAddendum.Columns(1).Caption = "Tipo Ord.Camb."
    grdAddendum.Columns(2).Width = 1000 ' fecha inicio
    grdAddendum.Columns(2).Caption = "F. Inicio"
    grdAddendum.Columns(3).Width = 1000 ' fecha final
    grdAddendum.Columns(3).Caption = "F. Final"
    grdAddendum.Columns(4).Width = 600 ' tipo moneda
    grdAddendum.Columns(4).Caption = "Moneda"
    grdAddendum.Columns(5).Width = 600 ' ts us
    grdAddendum.Columns(5).Caption = "TC US"
    grdAddendum.Columns(6).Width = 1000 ' monto us
    grdAddendum.Columns(6).Caption = "Monto $US"
    grdAddendum.Columns(7).Width = 1000 ' monto bs
    grdAddendum.Columns(7).Caption = "Monto Bs"
    grdAddendum.Columns(8).Width = 1000 ' tiempo
    grdAddendum.Columns(8).Caption = " Tiempo"
    grdAddendum.Columns(9).Width = 1500 ' justidica
    grdAddendum.Columns(9).Caption = "Justificación"
    
    For i = 10 To rstTemp.Fields.Count - 1
        grdAddendum.Columns(i).Visible = False
        grdAddendum.Columns(i).AllowSizing = False
    Next i
    
End Sub

Private Sub pl_HabilitaOpcion(swModo As Boolean)
    'TITULO:                Procedimiento pl_HabilitaOpcion
    
    ' habilitamos o deshabilitamos las opciones del menu
    cmdEditar.Enabled = Not swModo
    cmdEditar.Visible = Not swModo
    dtpFInicio.Enabled = swModo
    dtpFFin.Enabled = swModo
    chkConfirma.Enabled = swModo
    fraTiempo.Enabled = swModo
    cmdGuardar.Visible = swModo
    cmdCancela.Visible = swModo
    cmdEditar.Visible = Not swModo
    
    cmdEditarAdj.Enabled = Not swModo
    Call chkConfirma_Click
End Sub

Private Function fl_VerificaFechas() As Boolean
    'TITULO:                Función fl_VerificaFechas
    'PROPOSITO:             Verifica los datos para el registro de una liquidacion
    'EJEMPLO DE LLAMADA:    fl_VerificaFechas
    
    fl_VerificaFechas = True ' asuminos que se cuenta con los datos mnimos para grabar
    
    On Error GoTo EtiqError
     
    ' verificamos dechas del plazo
    If dtpFInicio.Value >= dtpFFin.Value Then
        MsgBox "No existe coherencia en el plazo de la consultoría." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        dtpFInicio.SetFocus
        fl_VerificaFechas = False
        Exit Function
    End If
    
    ' verificamos dechas del plazo
    If chkConfirma.Value = False Then
        MsgBox "NO se confirmo las fechas de inicio y finalización." & Chr(13) & "Corrija el error e intente guardar nuevamente.", vbInformation, "Aviso"
        dtpFInicio.SetFocus
        fl_VerificaFechas = False
        Exit Function
    End If
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Function

EtiqError:
    MsgBox "Error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description & Chr(13) & Chr(13) & "Anote el error y comuniquese con el soporte técnico.", vbCritical, "Error"

End Function

Private Sub pl_Llena_Combos_Base()
    ' llena los combos y listas base para la carga del formulario
    
    ' estado proceso = consultoria
    Set rs_AdjCont = New ADODB.Recordset
    SQLs = "SELECT cod_est_no_objecion, des_corta_est_no, activo FROM ac_estado_no_objecion_c where activo='S'"
    rs_AdjCont.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rs_AdjCont.RecordCount > 0 Then
        Set cboEstadoProceso.RowSource = rs_AdjCont
        cboEstadoProceso.BoundColumn = "cod_est_no_objecion"
        cboEstadoProceso.ListField = "des_corta_est_no"
      Else
        MsgBox "El catálogo de estados de proceso de consultoría no esta actualizado", vbInformation, "Aviso"
    End If
    
    ' tipo de moneda base
    Set rs_AdjCont = New ADODB.Recordset
    SQLs = "select tipo_moneda, tipo_moneda + ' - ' + denominacion_moneda as des_tipo_moneda from tipo_moneda where activo='S'"
    
    rs_AdjCont.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rs_AdjCont.RecordCount > 0 Then
        Set cboTipoMoneda.RowSource = rs_AdjCont
        cboTipoMoneda.BoundColumn = "tipo_moneda"
        cboTipoMoneda.ListField = "des_tipo_moneda"
        
      Else
        MsgBox "El catálogo de tipo de moneda no esta actualizado", vbInformation, "Aviso"
    End If
    
    ' forma de inicio
    Set rs_AdjCont = New ADODB.Recordset
    SQLs = "SELECT * FROM ac_forma_inicio_c where activo='S'"
    
    rs_AdjCont.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rs_AdjCont.RecordCount > 0 Then
        Set cboFormaInicio.RowSource = rs_AdjCont
        cboFormaInicio.BoundColumn = "cod_forma_inicio"
        cboFormaInicio.ListField = "des_forma_inicio"
      Else
        MsgBox "El catálogo de formas de inicio no esta actualizado", vbInformation, "Aviso"
    End If
    
    ' estado contrado
    Set rs_AdjCont = New ADODB.Recordset
    SQLs = "select cod_est_contrato, des_corta_cont from ac_estado_contrato_c where activo='S'"
    
    rs_AdjCont.Open SQLs, db, adOpenStatic, adLockReadOnly
    If rs_AdjCont.RecordCount > 0 Then
        Set cboEstadoCont.RowSource = rs_AdjCont
        cboEstadoCont.BoundColumn = "cod_est_contrato"
        cboEstadoCont.ListField = "des_corta_cont"
      Else
        MsgBox "El catálogo de estados de contrato no esta actualizado", vbInformation, "Aviso"
    End If
    
    Set rs_AdjCont = Nothing

End Sub

Private Sub lblDesOrgBase_Change()
    lblDesOrgBase.ToolTipText = lblDesOrgBase.Caption
End Sub

Private Sub lblDesOrgContra_Change()
    lblDesOrgContra.ToolTipText = lblDesOrgContra.Caption
End Sub

Private Sub optTiempo_Click(Index As Integer)
    On Error GoTo EtiqError
        
    Select Case Index
      Case 0 ' dia
        txtTiempo.Text = DateDiff("d", dtpFInicio.Value, dtpFFin.Value)
      Case 1 ' mes
        txtTiempo.Text = DateDiff("M", dtpFInicio.Value, dtpFFin.Value)
      Case 2 ' año
        txtTiempo.Text = DateDiff("yyyy", dtpFInicio.Value, dtpFFin.Value)
    End Select
    
    On Error GoTo 0 ' desactivamos el manejador de errores
    Exit Sub
    
EtiqError:

End Sub

Private Sub txtCodPrism_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 And txtCodPrism.Locked = False Then
        Call cmdCancelaAdj_Click
      Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

End Sub

Private Sub txtFileConsultoria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
      ElseIf KeyAscii = 27 And txtFileConsultoria.Locked = False Then
        Call cmdCancelaAdj_Click
      Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

End Sub

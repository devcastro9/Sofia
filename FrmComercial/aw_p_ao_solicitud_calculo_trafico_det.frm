VERSION 5.00
Begin VB.Form aw_p_ao_solicitud_calculo_trafico_det 
   Caption         =   "Cálculo de Tráfico - Datos Calculados"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   Icon            =   "aw_p_ao_solicitud_calculo_trafico_det.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   9315
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   8565
      Left            =   0
      ScaleHeight     =   8505
      ScaleWidth      =   9285
      TabIndex        =   0
      Top             =   0
      Width           =   9345
      Begin VB.Label lbl_campoc24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   76
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label lbl_campoc23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   75
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label lbl_campoc22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   74
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label lbl_campoc21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   73
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label lbl_campoc14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   72
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label lbl_campoc13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   71
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label lbl_campoc12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   70
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label lbl_campoc11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   69
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C. PARAMETROS ADOPTADOS"
         DataSource      =   "adoLista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Width           =   2910
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "D. VALORES CALCULADOS"
         DataSource      =   "adoLista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   67
         Top             =   2085
         Width           =   2580
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T-Asc/Desaceleración"
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
         Left            =   120
         TabIndex        =   66
         Top             =   740
         Width           =   2040
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "T-Entrada Salida"
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
         Left            =   120
         TabIndex        =   65
         Top             =   1700
         Width           =   1530
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "T-Apertura/Cierre"
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
         Left            =   120
         TabIndex        =   64
         Top             =   1220
         Width           =   1560
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo Recorrido"
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
         Left            =   120
         TabIndex        =   63
         Top             =   3240
         Width           =   1650
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Paradas Probables"
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
         Left            =   120
         TabIndex        =   62
         Top             =   2480
         Width           =   2040
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Capacid. Tiempo (CTi)"
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
         Left            =   120
         TabIndex        =   61
         Top             =   7580
         Width           =   2040
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Capacidad Tot.Arreglo"
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
         Left            =   120
         TabIndex        =   60
         Top             =   8060
         Width           =   2055
      End
      Begin VB.Label Label47 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "E. TIEMPOS CALCULADOS"
         DataSource      =   "adoLista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   59
         Top             =   2835
         Width           =   2520
      End
      Begin VB.Label Label48 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. TIEMPOS CALCULADOS - PORCENTAJE"
         DataSource      =   "adoLista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   58
         Top             =   5040
         Width           =   3990
      End
      Begin VB.Label Label49 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "G. RESULTADOS OBTENIDOS POR CADA ARREGLO"
         DataSource      =   "adoLista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   57
         Top             =   7200
         Width           =   4875
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T-Asc/Desaceleración"
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
         Left            =   120
         TabIndex        =   56
         Top             =   3700
         Width           =   2040
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "T-Entrada Salida"
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
         Left            =   120
         TabIndex        =   55
         Top             =   4660
         Width           =   1530
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "T-Apertura/Cierre"
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
         Left            =   120
         TabIndex        =   54
         Top             =   4180
         Width           =   1560
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo Recorrido"
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
         Left            =   120
         TabIndex        =   53
         Top             =   5420
         Width           =   1650
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T-Asc/Desaceleración"
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
         Left            =   120
         TabIndex        =   52
         Top             =   5860
         Width           =   2040
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "T-Entrada Salida"
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
         Left            =   120
         TabIndex        =   51
         Top             =   6820
         Width           =   1530
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "T-Apertura/Cierre"
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
         Left            =   120
         TabIndex        =   50
         Top             =   6340
         Width           =   1560
      End
      Begin VB.Label lbl_campog13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_g_capacidad_tiempo_cti3"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   49
         Top             =   7560
         Width           =   1605
      End
      Begin VB.Label lbl_campog12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_g_capacidad_tiempo_cti2"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   48
         Top             =   7560
         Width           =   1605
      End
      Begin VB.Label lbl_campog11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_g_capacidad_tiempo_cti"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   47
         Top             =   7560
         Width           =   1605
      End
      Begin VB.Label lbl_campog14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_g_capacidad_tiempo_cti4"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   46
         Top             =   7560
         Width           =   1605
      End
      Begin VB.Label lbl_campog23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_g_capacidad_total_arreglo3"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   45
         Top             =   8040
         Width           =   1605
      End
      Begin VB.Label lbl_campog22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_g_capacidad_total_arreglo2"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   44
         Top             =   8040
         Width           =   1605
      End
      Begin VB.Label lbl_campog21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_g_capacidad_total_arreglo"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   43
         Top             =   8040
         Width           =   1605
      End
      Begin VB.Label lbl_campog24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_g_capacidad_total_arreglo4"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   42
         Top             =   8040
         Width           =   1605
      End
      Begin VB.Label lbl_campof43 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_time_entrada_salida3"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   41
         Top             =   6800
         Width           =   1605
      End
      Begin VB.Label lbl_campof42 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_time_entrada_salida2"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   40
         Top             =   6800
         Width           =   1605
      End
      Begin VB.Label lbl_campof41 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_time_entrada_salida"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   39
         Top             =   6800
         Width           =   1605
      End
      Begin VB.Label lbl_campof44 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_time_entrada_salida4"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   38
         Top             =   6800
         Width           =   1605
      End
      Begin VB.Label lbl_campof33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_time_apertura_cierre3"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   37
         Top             =   6320
         Width           =   1605
      End
      Begin VB.Label lbl_campof32 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_time_apertura_cierre2"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   36
         Top             =   6320
         Width           =   1605
      End
      Begin VB.Label lbl_campof31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_time_apertura_cierre"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   35
         Top             =   6320
         Width           =   1605
      End
      Begin VB.Label lbl_campof34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_time_apertura_cierre4"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   34
         Top             =   6320
         Width           =   1605
      End
      Begin VB.Label lbl_campoc33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_c_time_entrada_salida3"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   33
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label lbl_campoc32 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_c_time_entrada_salida2"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   32
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label lbl_campoc31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_c_time_entrada_salida"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   31
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label lbl_campoc34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_c_time_entrada_salida4"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   30
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label lbl_campod13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_d_num_paradas_probables3"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   29
         Top             =   2440
         Width           =   1605
      End
      Begin VB.Label lbl_campod12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_d_num_paradas_probables2"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   28
         Top             =   2440
         Width           =   1605
      End
      Begin VB.Label lbl_campod11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_d_num_paradas_probables"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   27
         Top             =   2440
         Width           =   1605
      End
      Begin VB.Label lbl_campod14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_d_num_paradas_probables4"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   26
         Top             =   2440
         Width           =   1605
      End
      Begin VB.Label lbl_campoe13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_recorrido3"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   25
         Top             =   3200
         Width           =   1605
      End
      Begin VB.Label lbl_campoe12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_recorrido2"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   24
         Top             =   3200
         Width           =   1605
      End
      Begin VB.Label lbl_campoe11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_recorrido"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   23
         Top             =   3200
         Width           =   1605
      End
      Begin VB.Label lbl_campoe14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_recorrido4"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   22
         Top             =   3200
         Width           =   1605
      End
      Begin VB.Label lbl_campoe23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_asc_desaceleracion3"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   21
         Top             =   3680
         Width           =   1605
      End
      Begin VB.Label lbl_campoe22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_asc_desaceleracion2"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   20
         Top             =   3680
         Width           =   1605
      End
      Begin VB.Label lbl_campoe21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_asc_desaceleracion"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   19
         Top             =   3680
         Width           =   1605
      End
      Begin VB.Label lbl_campoe24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_asc_desaceleracion4"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   18
         Top             =   3680
         Width           =   1605
      End
      Begin VB.Label lbl_campoe33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_apertura_cierre3"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   17
         Top             =   4160
         Width           =   1605
      End
      Begin VB.Label lbl_campoe32 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_apertura_cierre2"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   16
         Top             =   4160
         Width           =   1605
      End
      Begin VB.Label lbl_campoe31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_apertura_cierre"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   15
         Top             =   4160
         Width           =   1605
      End
      Begin VB.Label lbl_campoe34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_apertura_cierre4"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   14
         Top             =   4160
         Width           =   1605
      End
      Begin VB.Label lbl_campoe43 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_entrada_salida3"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   13
         Top             =   4640
         Width           =   1605
      End
      Begin VB.Label lbl_campoe42 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_entrada_salida2"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   12
         Top             =   4640
         Width           =   1605
      End
      Begin VB.Label lbl_campoe41 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_entrada_salida"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   11
         Top             =   4640
         Width           =   1605
      End
      Begin VB.Label lbl_campoe44 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_e_tiempo_entrada_salida4"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   10
         Top             =   4640
         Width           =   1605
      End
      Begin VB.Label lbl_campof23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_time_asc_desaceleracion3"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   9
         Top             =   5840
         Width           =   1605
      End
      Begin VB.Label lbl_campof22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_time_asc_desaceleracion2"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   8
         Top             =   5840
         Width           =   1605
      End
      Begin VB.Label lbl_campof21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_time_asc_desaceleracion"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   7
         Top             =   5840
         Width           =   1605
      End
      Begin VB.Label lbl_campof24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_time_asc_desaceleracion4"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   6
         Top             =   5840
         Width           =   1605
      End
      Begin VB.Label lbl_campof13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_tiempo_recorrido3"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5780
         TabIndex        =   5
         Top             =   5400
         Width           =   1605
      End
      Begin VB.Label lbl_campof12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_tiempo_recorrido2"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4000
         TabIndex        =   4
         Top             =   5400
         Width           =   1605
      End
      Begin VB.Label lbl_campof11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_tiempo_recorrido"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2220
         TabIndex        =   3
         Top             =   5400
         Width           =   1605
      End
      Begin VB.Label lbl_campof14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label52"
         DataField       =   "trafico_f_tiempo_recorrido4"
         DataSource      =   "aw_p_ao_solicitud_calculo_trafico.Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   2
         Top             =   5400
         Width           =   1605
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "     Arreglo 1              Arreglo 2               Arreglo 3              Arreglo 4     "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   2160
         TabIndex        =   1
         Top             =   30
         Width           =   7035
      End
   End
End
Attribute VB_Name = "aw_p_ao_solicitud_calculo_trafico_det"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    'Call aw_p_ao_solicitud_calculo_trafico.OptFilGral1     '.OptFilGral1_Click
'    Call aw_p_ao_solicitud_calculo_trafico.OptFilGral1_Click
    Unload Me
End Sub


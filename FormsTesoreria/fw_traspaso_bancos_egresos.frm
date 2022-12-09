VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fw_traspaso_bancos_egresos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tesoreria - Traspasos Egresos"
   ClientHeight    =   10410
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   10815
   Icon            =   "fw_traspaso_bancos_egresos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   4.33764e5
   ScaleMode       =   0  'User
   ScaleWidth      =   2.19418e7
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra_reporte 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Elija un item del Extracto Bancario ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3135
      Left            =   1560
      TabIndex        =   102
      Top             =   3840
      Visible         =   0   'False
      Width           =   16935
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   3000
         TabIndex        =   119
         Text            =   "0"
         Top             =   1800
         Visible         =   0   'False
         Width           =   9015
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
         ScaleWidth      =   16680
         TabIndex        =   109
         Top             =   2040
         Width           =   16680
         Begin VB.PictureBox BtnImprimir2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   112
            ToolTipText     =   "Imprimir el Listado de los Registros"
            Top             =   120
            Width           =   1455
         End
         Begin VB.PictureBox BtnCancelar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   8760
            Picture         =   "fw_traspaso_bancos_egresos.frx":0A02
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   111
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6480
            Picture         =   "fw_traspaso_bancos_egresos.frx":12EE
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   110
            Top             =   0
            Width           =   1280
         End
         Begin VB.Label Label22 
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
            TabIndex        =   113
            Top             =   195
            Visible         =   0   'False
            Width           =   1005
         End
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   108
         Text            =   "0"
         Top             =   1320
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo DctOrigina18 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":1AC4
         DataField       =   "cmpbte_deposito_bco"
         DataSource      =   "Ado_datos02"
         Height          =   315
         Left            =   12960
         TabIndex        =   103
         Top             =   720
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "descripcion"
         BoundColumn     =   "cod_bancarizacion"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DctDeposita18 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":1ADE
         DataField       =   "cmpbte_deposito_bco"
         DataSource      =   "Ado_datos02"
         Height          =   315
         Left            =   9720
         TabIndex        =   104
         Top             =   720
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "nombre_depositante"
         BoundColumn     =   "cod_bancarizacion"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DctCliente18 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":1AF8
         DataField       =   "cmpbte_deposito_bco"
         DataSource      =   "Ado_datos02"
         Height          =   315
         Left            =   8640
         TabIndex        =   105
         Top             =   720
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "cod_cliente"
         BoundColumn     =   "cod_bancarizacion"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DctCuenta18 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":1B12
         DataField       =   "cmpbte_deposito_bco"
         DataSource      =   "Ado_datos02"
         Height          =   315
         Left            =   6360
         TabIndex        =   106
         Top             =   720
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "cuenta"
         BoundColumn     =   "cod_bancarizacion"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DctMontoDol18 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":1B2C
         DataField       =   "cmpbte_deposito_bco"
         DataSource      =   "Ado_datos02"
         Height          =   315
         Left            =   5040
         TabIndex        =   107
         Top             =   720
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "monto_dol"
         BoundColumn     =   "cod_bancarizacion"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTP_Finicio 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   8280
         TabIndex        =   114
         Top             =   1320
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   119144449
         CurrentDate     =   44457
      End
      Begin MSComCtl2.DTPicker DTP_Ffin 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   8280
         TabIndex        =   115
         Top             =   1320
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   119144449
         CurrentDate     =   42880
      End
      Begin MSDataListLib.DataCombo DctMonto18 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":1B46
         DataField       =   "cmpbte_deposito_bco"
         DataSource      =   "Ado_datos02"
         Height          =   315
         Left            =   3720
         TabIndex        =   116
         Top             =   720
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "monto_bs"
         BoundColumn     =   "cod_bancarizacion"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DctCod18 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":1B60
         DataField       =   "cmpbte_deposito_bco"
         DataSource      =   "Ado_datos02"
         Height          =   315
         Left            =   240
         TabIndex        =   117
         Top             =   720
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "cod_bancarizacion"
         BoundColumn     =   "cod_bancarizacion"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DctFecha18 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":1B7A
         DataField       =   "cmpbte_deposito_bco"
         DataSource      =   "Ado_datos02"
         Height          =   315
         Left            =   2160
         TabIndex        =   118
         Top             =   720
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "fecha_transaccion"
         BoundColumn     =   "cod_bancarizacion"
         Text            =   "Todos"
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA TRANSACCION"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6360
         TabIndex        =   130
         Top             =   1365
         Width           =   1725
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "# COMPROBANTE DEPOSITO"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   165
         TabIndex        =   129
         Top             =   1365
         Width           =   2295
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCIAS DEL DEPOSITANTE"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   128
         Top             =   1800
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   13080
         TabIndex        =   127
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre.Depositante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   9960
         TabIndex        =   126
         Top             =   360
         Width           =   1875
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cod.Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   8640
         TabIndex        =   125
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Importe.Bs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   3960
         TabIndex        =   124
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo.Bancarizacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   240
         TabIndex        =   123
         Top             =   360
         Width           =   1980
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha.Extracto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   2400
         TabIndex        =   122
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Importe.Dol."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   5160
         TabIndex        =   121
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta.Bancaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   6480
         TabIndex        =   120
         Top             =   360
         Width           =   1485
      End
   End
   Begin VB.Frame frm_benef 
      BackColor       =   &H00808080&
      Caption         =   "Elije un Beneficiario y Registra su Cuenta Bancaria"
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
      Height          =   2535
      Left            =   4560
      TabIndex        =   34
      Top             =   720
      Visible         =   0   'False
      Width           =   12015
      Begin VB.TextBox dtc_nom8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3840
         MultiLine       =   -1  'True
         TabIndex        =   47
         Text            =   "fw_traspaso_bancos_egresos.frx":1B94
         Top             =   1680
         Width           =   5805
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9345
         TabIndex        =   45
         Top             =   1215
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox dtc_cta8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   720
         MultiLine       =   -1  'True
         TabIndex        =   43
         Text            =   "fw_traspaso_bancos_egresos.frx":1B96
         Top             =   1680
         Width           =   3045
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9360
         TabIndex        =   37
         Top             =   855
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton BtnCancelarBen 
         BackColor       =   &H80000015&
         Height          =   635
         Left            =   10200
         MaskColor       =   &H00000000&
         Picture         =   "fw_traspaso_bancos_egresos.frx":1B98
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Cancelar"
         Top             =   1560
         Width           =   1365
      End
      Begin VB.CommandButton BtnGrabarBen 
         BackColor       =   &H80000015&
         Height          =   635
         Left            =   10200
         Picture         =   "fw_traspaso_bancos_egresos.frx":2484
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   600
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":2C72
         DataField       =   "beneficiario_codigo_fac"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   7320
         TabIndex        =   38
         Top             =   1200
         Visible         =   0   'False
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8421504
         ForeColor       =   16777215
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "00000000000004"
      End
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":2C8B
         DataField       =   "beneficiario_codigo_fac"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   720
         TabIndex        =   39
         Top             =   840
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux8 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":2CA4
         DataField       =   "beneficiario_codigo_fac"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   7320
         TabIndex        =   40
         Top             =   840
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8421504
         ForeColor       =   16777215
         ListField       =   "beneficiario_nit"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "00000000000004"
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Número de la Cuenta Bancaria         Nombre de la Cuenta Bancaria del Proveedor"
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
         Left            =   720
         TabIndex        =   44
         Top             =   1365
         Width           =   7245
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de Proveedor"
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
         Left            =   720
         TabIndex        =   42
         Top             =   585
         Width           =   2025
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "NIT"
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
         Left            =   7440
         TabIndex        =   41
         Top             =   600
         Width           =   2025
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      FillStyle       =   2  'Horizontal Line
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   18600
      TabIndex        =   88
      Top             =   0
      Width           =   18600
      Begin VB.PictureBox BtnBuscar1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   10200
         Picture         =   "fw_traspaso_bancos_egresos.frx":2CBD
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   93
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17040
         Picture         =   "fw_traspaso_bancos_egresos.frx":3472
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   92
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8715
         Picture         =   "fw_traspaso_bancos_egresos.frx":3C34
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   90
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7440
         Picture         =   "fw_traspaso_bancos_egresos.frx":4520
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   89
         Top             =   0
         Visible         =   0   'False
         Width           =   1280
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ORIGEN - Ordenes de Cancelación (OC) a Proveedores (PENDIENTES)"
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
         Left            =   255
         TabIndex        =   91
         Top             =   180
         Width           =   8025
      End
   End
   Begin VB.Frame FrmCabecera 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Left            =   4680
      TabIndex        =   48
      Top             =   720
      Visible         =   0   'False
      Width           =   11895
      Begin VB.PictureBox FraGrabarCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   11640
         TabIndex        =   82
         Top             =   3960
         Width           =   11640
         Begin VB.PictureBox BtnGrabar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4680
            Picture         =   "fw_traspaso_bancos_egresos.frx":4CF6
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   84
            Top             =   0
            Width           =   1280
         End
         Begin VB.PictureBox BtnCancelar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5955
            Picture         =   "fw_traspaso_bancos_egresos.frx":54CC
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   83
            Top             =   0
            Width           =   1455
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
            ForeColor       =   &H00FFFF80&
            Height          =   285
            Left            =   8775
            TabIndex        =   85
            Top             =   180
            Width           =   1005
         End
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   8025
         TabIndex        =   67
         Top             =   390
         Width           =   270
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   8715
         TabIndex        =   66
         Top             =   30
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Frame Fra_datos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "-------------------------------- ORIGEN "
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
         Height          =   2445
         Left            =   45
         TabIndex        =   58
         Top             =   1395
         Width           =   5895
         Begin VB.ComboBox cmd_unimed2 
            DataField       =   "unimed_codigo_cobr"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6210
            TabIndex        =   60
            Text            =   "ANUAL"
            Top             =   1080
            Visible         =   0   'False
            Width           =   555
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "fw_traspaso_bancos_egresos.frx":5DB8
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1260
            TabIndex        =   59
            Top             =   660
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "fw_traspaso_bancos_egresos.frx":5DD1
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   61
            Top             =   660
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   14737632
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo dtc_desc21 
            Bindings        =   "fw_traspaso_bancos_egresos.frx":5DEA
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   62
            Top             =   1920
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   14737632
            ListField       =   "cta_descripcion"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo21 
            Bindings        =   "fw_traspaso_bancos_egresos.frx":5E04
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   63
            Top             =   1560
            Width           =   2970
            _ExtentX        =   5239
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "cta_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Bancaria o Caja ORIGEN:"
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
            Top             =   1245
            Width           =   2985
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Realizado por:"
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
            Top             =   360
            Width           =   1320
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "----------------------------- DESTINO "
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
         Height          =   2445
         Left            =   5960
         TabIndex        =   49
         Top             =   1395
         Width           =   5895
         Begin VB.CommandButton BtnNuevaCta 
            BackColor       =   &H00C0E0FF&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5120
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   1080
            Width           =   645
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "fw_traspaso_bancos_egresos.frx":5E1E
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1395
            TabIndex        =   51
            Top             =   660
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "fw_traspaso_bancos_egresos.frx":5E37
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   195
            TabIndex        =   52
            Top             =   660
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   14737632
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo dtc_desc22 
            Bindings        =   "fw_traspaso_bancos_egresos.frx":5E50
            DataField       =   "cta_codigo_destino"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   195
            TabIndex        =   53
            Top             =   1920
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   14737632
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo22 
            Bindings        =   "fw_traspaso_bancos_egresos.frx":5E6A
            DataField       =   "cta_codigo_destino"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   195
            TabIndex        =   54
            Top             =   1560
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "cta_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_moneda22 
            Bindings        =   "fw_traspaso_bancos_egresos.frx":5E84
            DataField       =   "cta_codigo_destino"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3400
            TabIndex        =   55
            Top             =   1560
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   ""
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Verificado por:"
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
            Left            =   195
            TabIndex        =   57
            Top             =   360
            Width           =   1305
         End
         Begin VB.Label lbl_Rdestino 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Bancaria o Caja DESTINO -  Proveedor"
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
            Left            =   195
            TabIndex        =   56
            Top             =   1245
            Width           =   4215
         End
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":5E9E
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7065
         TabIndex        =   68
         Top             =   375
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   14737632
         ForeColor       =   0
         ListField       =   "doc_codigo"
         BoundColumn     =   "doc_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":5EB7
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1725
         TabIndex        =   69
         Top             =   370
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   14737632
         ForeColor       =   0
         ListField       =   "doc_descripcion"
         BoundColumn     =   "doc_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":5ED0
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6960
         TabIndex        =   70
         Top             =   240
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "clasif_codigo"
         BoundColumn     =   "doc_codigo"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTPfechasol 
         DataField       =   "fecha_traspaso"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   1755
         TabIndex        =   71
         Top             =   960
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         Format          =   119144449
         CurrentDate     =   44856
         MaxDate         =   55153
         MinDate         =   2
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Total Dolares"
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
         Height          =   285
         Left            =   8460
         TabIndex        =   81
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Total Bs."
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
         Index           =   0
         Left            =   4440
         TabIndex        =   80
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "Total_bs"
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5880
         TabIndex        =   79
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Documento ISO"
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
         Left            =   180
         TabIndex        =   78
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label txt_campo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "correl_doc"
         DataSource      =   "Ado_datos"
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
         Height          =   300
         Left            =   10395
         TabIndex        =   77
         Top             =   370
         Width           =   1365
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Traspaso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   13
         Left            =   8760
         TabIndex        =   76
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label txt_venta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "total_dol"
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   10440
         TabIndex        =   75
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label lbl_cerrado 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "TRASPASO CONCILIADO !!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   3480
         TabIndex        =   74
         Top             =   -30
         Width           =   4875
      End
      Begin VB.Label Label21 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Traspaso"
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
         Left            =   180
         TabIndex        =   73
         Top             =   960
         Width           =   1710
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         DataField       =   "Correl_doc"
         DataSource      =   "Ado_datos11"
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
         Height          =   300
         Left            =   10395
         TabIndex        =   72
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame FraDet3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija una de las 2 Opciones ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2280
      Left            =   4800
      TabIndex        =   24
      Top             =   3120
      Visible         =   0   'False
      Width           =   11820
      Begin VB.CommandButton BtnCancelar2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Todos de la Orden de Pago"
         Height          =   735
         Left            =   5280
         Picture         =   "fw_traspaso_bancos_egresos.frx":5EE9
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Acepta todos los Registro del Recibo de Tesorería ""RboTes"""
         Top             =   1320
         Width           =   1485
      End
      Begin VB.CommandButton BtnGrabar2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Registro Elegido"
         Height          =   735
         Left            =   960
         Picture         =   "fw_traspaso_bancos_egresos.frx":68EB
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Acepta SOLO el Registro elegido..."
         Top             =   1320
         Width           =   1365
      End
      Begin VB.TextBox txtRecTes 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "fw_traspaso_bancos_egresos.frx":6AF5
         Top             =   720
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton BtnCancelar1 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   9360
         Picture         =   "fw_traspaso_bancos_egresos.frx":6AF7
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Acepta SOLO el Registro elegido..."
         Top             =   1440
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo dtc_desc6 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":74B6
         DataField       =   "IdRecibo"
         DataSource      =   "ado_datos14"
         Height          =   315
         Left            =   4380
         TabIndex        =   29
         Top             =   1680
         Visible         =   0   'False
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "IdRecibo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo6 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":74CF
         DataField       =   "IdRecibo"
         DataSource      =   "ado_datos14"
         Height          =   315
         Left            =   6720
         TabIndex        =   30
         Top             =   720
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "IdRecibo"
         BoundColumn     =   "IdRecibo"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_recibo6 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":74E8
         DataField       =   "IdRecibo"
         DataSource      =   "ado_datos14"
         Height          =   315
         Left            =   5520
         TabIndex        =   31
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "IdRecibo"
         BoundColumn     =   "IdRecibo"
         Text            =   "0"
      End
      Begin VB.Label lbl_orden 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "1.Registro Elegido"
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
         Height          =   240
         Left            =   840
         TabIndex        =   33
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label lbl_orden_camb 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "2. Registros de Orden de Pago . . . (Ord.Pago) :"
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
         Height          =   555
         Left            =   4320
         TabIndex        =   32
         Top             =   375
         Width           =   3375
      End
   End
   Begin VB.Frame FrmDetalle2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE DE PAGOS"
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
      Height          =   2025
      Left            =   1560
      TabIndex        =   19
      Top             =   7320
      Width           =   17055
      Begin MSDataGridLib.DataGrid DtGLista11 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":7501
         Height          =   1740
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   16860
         _ExtentX        =   29739
         _ExtentY        =   3069
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
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
         ColumnCount     =   20
         BeginProperty Column00 
            DataField       =   "adjudica_codigo"
            Caption         =   "#CorrelPago"
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
         BeginProperty Column01 
            DataField       =   "IdRecibo"
            Caption         =   "#.OC"
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
            DataField       =   "fecha_recibo"
            Caption         =   "Fecha.OC"
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
            DataField       =   "venta_tipo"
            Caption         =   "Tipo.Contrato"
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
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cite.Contrato"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "trans_descripcion"
            Caption         =   "Tipo.Transac."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "cmpbte_deposito"
            Caption         =   "#Cheque/Transf."
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
         BeginProperty Column07 
            DataField       =   "cmpbte_deposito_bco"
            Caption         =   "#Cmpbte.Deposito"
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
            DataField       =   "fecha_registro_bco"
            Caption         =   "Fecha.Deposito"
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
            DataField       =   "adjudica_bs"
            Caption         =   "Pagado Bs."
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
            DataField       =   "adjudica_dol"
            Caption         =   "Pagado Dol."
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
            DataField       =   "edif_codigo_corto"
            Caption         =   "Edificio"
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
            DataField       =   "edif_descripcion"
            Caption         =   "Nombre.Edificio"
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
            DataField       =   "observaciones"
            Caption         =   "Concepto"
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
            DataField       =   "estado_destino"
            Caption         =   "Pagado"
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
         BeginProperty Column16 
            DataField       =   "IdTraspasoBancos"
            Caption         =   "Id.Trp."
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
         BeginProperty Column17 
            DataField       =   "estado_codigo"
            Caption         =   "Aceptado"
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
            DataField       =   "usr_codigo"
            Caption         =   "Usuario"
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
         BeginProperty Column19 
            DataField       =   "cta_codigo"
            Caption         =   "Cuenta.Bancaria"
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
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   3495.118
            EndProperty
            BeginProperty Column14 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   3075.024
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column16 
               Alignment       =   2
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column17 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column18 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column19 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1319.811
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
      ScaleWidth      =   18600
      TabIndex        =   9
      Top             =   3120
      Width           =   18600
      Begin VB.PictureBox BtnAprobar1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   14400
         Picture         =   "fw_traspaso_bancos_egresos.frx":751B
         ScaleHeight     =   735
         ScaleWidth      =   1320
         TabIndex        =   20
         ToolTipText     =   "Verifica Comprobante de Traspaso"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnDesAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   14400
         Picture         =   "fw_traspaso_bancos_egresos.frx":7D53
         ScaleHeight     =   735
         ScaleWidth      =   1320
         TabIndex        =   21
         ToolTipText     =   "Aprueba Comprobante de Traspaso"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   13080
         Picture         =   "fw_traspaso_bancos_egresos.frx":874A
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   10
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   15720
         Picture         =   "fw_traspaso_bancos_egresos.frx":8EFF
         ScaleHeight     =   735
         ScaleWidth      =   1320
         TabIndex        =   11
         ToolTipText     =   "Aprueba Comprobante de Traspaso"
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   11880
         Picture         =   "fw_traspaso_bancos_egresos.frx":9732
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   12
         ToolTipText     =   "Anula Registro"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   17040
         Picture         =   "fw_traspaso_bancos_egresos.frx":9E7E
         ScaleHeight     =   735
         ScaleWidth      =   1395
         TabIndex        =   15
         ToolTipText     =   "Comprobante de Arqueo de Traspasos"
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   9360
         Picture         =   "fw_traspaso_bancos_egresos.frx":A74B
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   14
         ToolTipText     =   "Nuevo Arqueo de Traspasos"
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   10545
         Picture         =   "fw_traspaso_bancos_egresos.frx":AF0A
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   13
         ToolTipText     =   "Modifica datos del arqueo"
         Top             =   0
         Width           =   1430
      End
      Begin VB.Label lbl_titulo 
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
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   360
         TabIndex        =   16
         Top             =   195
         Width           =   1815
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTA"
      ForeColor       =   &H00C00000&
      Height          =   3480
      Left            =   1560
      TabIndex        =   4
      Top             =   3840
      Width           =   17025
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00C0E0FF&
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   9600
         TabIndex        =   7
         Top             =   3195
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00C0E0FF&
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   5280
         TabIndex        =   6
         Top             =   3195
         Value           =   -1  'True
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   2850
         Left            =   75
         TabIndex        =   5
         Top             =   240
         Width           =   16785
         _ExtentX        =   29607
         _ExtentY        =   5027
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
         ColumnCount     =   16
         BeginProperty Column00 
            DataField       =   "clasif_codigo"
            Caption         =   "Clasificacion"
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
            DataField       =   "doc_codigo"
            Caption         =   "Doc.ISO"
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
         BeginProperty Column02 
            DataField       =   "IdTraspasoBancos"
            Caption         =   "Nro.Traspaso"
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
         BeginProperty Column03 
            DataField       =   "fecha_traspaso"
            Caption         =   "Fecha.Traspaso"
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
            DataField       =   "total_bs"
            Caption         =   "Total.Bs."
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
            DataField       =   "total_dol"
            Caption         =   "Total.Dolares"
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
            DataField       =   "unidad_codigo_resp"
            Caption         =   "Unidad.Origen"
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
            DataField       =   "unidad_codigo"
            Caption         =   "Unidad.Destino"
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
         BeginProperty Column08 
            DataField       =   "cta_codigo"
            Caption         =   "Cuenta.Origen"
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
            DataField       =   "Cta_codigo_destino"
            Caption         =   "Cuenta.Destino"
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "CI_Entrega"
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
            DataField       =   "cta_descripcion"
            Caption         =   "Nombre.de.la.Cuenta"
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
            DataField       =   "estado_verificado"
            Caption         =   "Tesoreria"
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
            DataField       =   "estado_codigo"
            Caption         =   "Supervisor"
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
            DataField       =   "fecha_registro"
            Caption         =   "Fecha.Registro"
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
            DataField       =   "usr_codigo"
            Caption         =   "Usuario"
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
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   3300.095
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column13 
               Alignment       =   2
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   75
         Top             =   3120
         Width           =   16785
         _ExtentX        =   29607
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
   Begin VB.Frame FrmDetalle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE DE PAGOS"
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
      Height          =   2385
      Left            =   1320
      TabIndex        =   3
      Top             =   660
      Width           =   17295
      Begin MSDataGridLib.DataGrid DtgLista 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":B81F
         Height          =   2100
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   17100
         _ExtentX        =   30163
         _ExtentY        =   3704
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
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
         ColumnCount     =   23
         BeginProperty Column00 
            DataField       =   "adjudica_codigo"
            Caption         =   "#Correl.Pago"
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
            DataField       =   "IdRecibo"
            Caption         =   "#.OC"
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
            DataField       =   "fecha_recibo"
            Caption         =   "Fecha.OC."
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
            DataField       =   "sigla_emprea"
            Caption         =   "Empresa"
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
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cite.Contrato"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "trans_descripcion_fac"
            Caption         =   "Tipo.Transac."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "nro_nota_remision"
            Caption         =   "#Factura.Recibo"
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
         BeginProperty Column07 
            DataField       =   "adjudica_fecha"
            Caption         =   "Fecha.Factura.Recibo"
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
         BeginProperty Column09 
            DataField       =   "adjudica_bs"
            Caption         =   "Pagado Bs."
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
         BeginProperty Column10 
            DataField       =   "adjudica_dol"
            Caption         =   "Pagado Dol."
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
            DataField       =   "Observaciones"
            Caption         =   "Proveedor"
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
            DataField       =   "edif_codigo_corto"
            Caption         =   "Edificio"
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
            DataField       =   "edif_descripcion"
            Caption         =   "Nombre.Edificio"
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
            DataField       =   "observaciones"
            Caption         =   "Concepto"
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
            DataField       =   "estado_destino"
            Caption         =   "Pagado"
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
         BeginProperty Column16 
            DataField       =   "venta_tipo"
            Caption         =   "Tipo.Contrato"
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
         BeginProperty Column17 
            DataField       =   "IdTraspasoBancos"
            Caption         =   "Id.Trp."
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
         BeginProperty Column18 
            DataField       =   "estado_codigo"
            Caption         =   "Aceptado"
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
            DataField       =   "usr_codigo"
            Caption         =   "Usuario"
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
         BeginProperty Column20 
            DataField       =   "cta_codigo"
            Caption         =   "Cuenta.Bancaria"
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
            DataField       =   "adjudica_codigo"
            Caption         =   "Id.Adjudica"
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
         BeginProperty Column22 
            DataField       =   "unidad_codigo_adm"
            Caption         =   "Unidad.Organizacional"
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
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   2819.906
            EndProperty
            BeginProperty Column14 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   3075.024
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column17 
               Alignment       =   2
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column18 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column19 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column20 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column21 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column22 
            EndProperty
         EndProperty
      End
   End
   Begin Crystal.CrystalReport CryV01 
      Left            =   0
      Top             =   9360
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6720
      Top             =   8760
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
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   11280
      Top             =   9480
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
   Begin MSAdodcLib.Adodc ado_datos14 
      Height          =   330
      Left            =   11280
      Top             =   9120
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "ado_datos14"
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
   Begin MSAdodcLib.Adodc ado_datos17 
      Height          =   330
      Left            =   9000
      Top             =   9120
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
      Caption         =   "ado_datos17"
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
      Left            =   -120
      Top             =   9840
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
   Begin MSAdodcLib.Adodc Ado_datos16 
      Height          =   330
      Left            =   13560
      Top             =   9120
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Ado_datos16"
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
   Begin MSAdodcLib.Adodc ado_datos15 
      Height          =   330
      Left            =   6720
      Top             =   9120
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
      Caption         =   "ado_datos15"
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
      Left            =   11280
      Top             =   8760
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
   Begin MSAdodcLib.Adodc Ado_Datos12 
      Height          =   330
      Left            =   2160
      Top             =   9120
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
      Caption         =   "Ado_Datos12"
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
   Begin MSAdodcLib.Adodc Ado_datos13 
      Height          =   330
      Left            =   4440
      Top             =   9120
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
      Caption         =   "Ado_datos13"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   13560
      Top             =   8760
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
      Caption         =   "AdoAux"
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
      Left            =   4440
      Top             =   8760
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   13560
      Top             =   9480
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
   Begin MSAdodcLib.Adodc ado_datos6 
      Height          =   330
      Left            =   9000
      Top             =   9480
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
      Caption         =   "ado_datos6"
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
   Begin Crystal.CrystalReport CryR01 
      Left            =   480
      Top             =   9360
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
   Begin MSAdodcLib.Adodc Ado_datos20 
      Height          =   330
      Left            =   -120
      Top             =   9480
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
      Caption         =   "Ado_datos20"
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
      Left            =   2160
      Top             =   9480
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
   Begin MSAdodcLib.Adodc Ado_datos22 
      Height          =   330
      Left            =   4440
      Top             =   9480
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
      Caption         =   "Ado_datos22"
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
   Begin MSAdodcLib.Adodc AdoAux9 
      Height          =   330
      Left            =   6720
      Top             =   9480
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
      Caption         =   "AdoAux9"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ORIGEN - Ordenes de Pago a Proveedores"
      ForeColor       =   &H00C00000&
      Height          =   2040
      Left            =   1800
      TabIndex        =   86
      Top             =   720
      Visible         =   0   'False
      Width           =   9225
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "fw_traspaso_bancos_egresos.frx":B839
         Height          =   1410
         Left            =   75
         TabIndex        =   87
         Top             =   240
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   2487
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
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "clasif_codigo"
            Caption         =   "Clasificacion"
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
            DataField       =   "doc_codigo"
            Caption         =   "Doc.ISO"
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
         BeginProperty Column02 
            DataField       =   "correl_doc"
            Caption         =   "Orden.Pago"
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
         BeginProperty Column03 
            DataField       =   "fecha_recibo"
            Caption         =   "Fecha.Recibo"
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
            DataField       =   "total_bs"
            Caption         =   "Total.Bs."
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
         BeginProperty Column06 
            DataField       =   "beneficiario_codigo"
            Caption         =   "CI_Entrega"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "CI_Recibe"
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
         BeginProperty Column08 
            DataField       =   "total_dol"
            Caption         =   "Total.Dolares"
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
         BeginProperty Column09 
            DataField       =   "fecha_registro"
            Caption         =   "Fecha.Registro"
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
            DataField       =   "usr_codigo"
            Caption         =   "Usuario"
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
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   -1  'True
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   75
         Top             =   1680
         Width           =   8985
         _ExtentX        =   15849
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
         BackColor       =   -2147483624
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
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   2625
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   10.688
      ScaleMode       =   4  'Character
      ScaleWidth      =   9.625
      TabIndex        =   94
      Top             =   480
      Width           =   1215
      Begin VB.PictureBox BtnAddDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   0
         Picture         =   "fw_traspaso_bancos_egresos.frx":B851
         ScaleHeight     =   1095
         ScaleWidth      =   1200
         TabIndex        =   95
         ToolTipText     =   "Aprueba Comprobante de Traspaso"
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Elije los registros de la lista y ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   0
         TabIndex        =   99
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PASO 2. --->"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Left            =   0
         TabIndex        =   98
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   5625
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   23.188
      ScaleMode       =   4  'Character
      ScaleWidth      =   11.625
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
      Begin VB.PictureBox BtnAnlDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         Picture         =   "fw_traspaso_bancos_egresos.frx":C398
         ScaleHeight     =   1095
         ScaleWidth      =   1335
         TabIndex        =   23
         ToolTipText     =   "Anula Registro"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.PictureBox BtnModDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "fw_traspaso_bancos_egresos.frx":CFE6
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   22
         ToolTipText     =   "Modifica Fecha y Código de Bancarización"
         Top             =   4680
         Width           =   1430
      End
      Begin VB.CommandButton BtnImprimir1 
         BackColor       =   &H80000018&
         Height          =   525
         Left            =   0
         Picture         =   "fw_traspaso_bancos_egresos.frx":D8FB
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Imprime Kardex del Bien"
         Top             =   2790
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Verifica (Aprueba) el Traspaso ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   0
         TabIndex        =   101
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PASO 3.   --->"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Left            =   0
         TabIndex        =   100
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crea y/o Identifica el Traspaso  ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   930
         Left            =   0
         TabIndex        =   97
         Top             =   480
         Width           =   1425
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PASO 1.   --->"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Left            =   60
         TabIndex        =   96
         Top             =   240
         Width           =   1305
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos18 
      Height          =   330
      Left            =   2160
      Top             =   9840
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
      Caption         =   "Ado_datos18"
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
   Begin VB.Label LblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label LblUni_descripcion_larga 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   225
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.Label lblUni_codigo 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "fw_traspaso_bancos_egresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************
'Ventas
Dim rs_datos As New ADODB.Recordset     'VENTAS
Dim rs_datos1 As New ADODB.Recordset    'UNIDAD EJECUTORA
Dim rs_datos2 As New ADODB.Recordset    'Beneficiario Personas Nat. y Juridicas (menos de CGI)
Dim rs_datos3 As New ADODB.Recordset    'Proyecto de Edificacion
Dim rs_datos4 As New ADODB.Recordset    'Beneficiario Funcionario de CGI (Vendedor, Cobrador, Admin, etc.)
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset
Dim rs_datos12 As New ADODB.Recordset
Dim rs_datos13 As New ADODB.Recordset
Dim rs_datos14 As New ADODB.Recordset   'Ventas_detalle
Dim rs_datos15 As New ADODB.Recordset
Dim rs_datos16 As New ADODB.Recordset   'Ventas cobranzas
Dim rs_datos17 As New ADODB.Recordset
Dim rs_datos18 As New ADODB.Recordset

Dim rs_datos19 As New ADODB.Recordset   'Acumula Cobranzas
Dim rs_datos20 As New ADODB.Recordset
Dim rs_datos21 As New ADODB.Recordset
Dim rs_datos22 As New ADODB.Recordset

'AUXILIARES
Dim rs_Ventas_lista As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset
Dim rs_aux9 As New ADODB.Recordset
Dim rstdestino As New ADODB.Recordset
Dim rstcorrel_ing As New ADODB.Recordset
Dim rs_precio As New ADODB.Recordset

Dim rsNada As ADODB.Recordset
'OTROS
'Dim rstdetsalalm As New ADODB.Recordset
Dim RS_BENEF As New ADODB.Recordset
Dim rs_TipoCambio As New ADODB.Recordset
Dim rs_almacen2 As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim rsAuxDetalle As New ADODB.Recordset

'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir, accion As String
'Dim queryinicial As String
Public queryinicial2 As String

'Almacenes
Dim descri_bien As String
Dim VAR_ALMX As String
Dim VAR_ALMT As String
Dim tipo_alm As String
Dim VAR_DOC As String
Dim VAR_DA As String
Dim VAR_ALMD As String
Dim VAR_ORIGEN As String
Dim VAR_DOCI, VAR_DOCR, VAR_DOCH, VAR_DOCA As String
Dim VAR_BENI, VAR_BENR, VAR_BENH, VAR_BENA As String
Dim VAR_BENDI, VAR_BENDR, VAR_BENDH, VAR_BENDA As String
Dim VAR_NUMI, VAR_NUMR, VAR_NUMH, VAR_NUMA As String

Dim Cant_Alm, VAR_CANT As Integer
Dim correlativo1, VAR_RECIBO As Integer
Dim VAR_IDTRP As String
'Dim VAR_ALMI, VAR_ALMR, VAR_ALMH, VAR_ALMA As Integer
'Dim VAR_ALMDI, VAR_ALMDR, VAR_ALMDH, VAR_ALMDA As Integer

'VARIABLES
Dim marca1 As Variant

Dim swgrabar, swnuevo, deta2, CONT_MED As Integer
Dim nroventa, correlv, correldet2, corrprog As Integer
Dim VAR_PARTIDA, VAR_PROY, correldetalle As Integer
Dim VAR_CODANT, Var_Comp, VAR_SOL, CANTOT, var_cod5 As Integer
Dim CONT2, CONT3, CONT4, VAR_TIPO As Integer
Dim fdia, fmes, fanio, Dias_Mes, TimeD  As Integer
Dim VAR_COBR1, VAR_COBR2, VAR_CONTR As Integer
Dim VAR_NUM, var_cod, VAR_COD2 As Integer
Dim VARFILTRO, SWFILTRO As Integer

Dim Cobrobs, VAR_COBR, VAR_AUX, VAR_AUX2 As Double
Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2, VAR_MBS2, VAR_MDOL2 As Double

Dim VAR_DET As String
Dim gestion0, var_literal, VAR_PROY2, VAR_CITE, VAR_CTA As String
Dim VAR_CODTIPO, VAR_ORG, VAR_FTE, VAR_BENEF, VAR_GLOSA, VAR_GLOSA2, VAR_MONEDA As String
Dim VAR_BEND, VAR_EDIFD, VARG_ORGD, VAR_CTAD, VAR_UNID, VAR_DPTO, VAR_DPTOD As String
Dim VAR_COD1, VAR_BIEN2, VAR_COD3, VAR_COD4 As String
Dim VAR_MED, VAR_MED2 As String
Dim VAR_TIPOV, VAR_VAL As String
Dim VAR_FEC2, MControl, VAR_MES2 As String
Dim VAR_BEN2, VAR_BEN3, VAR_ALM As String
Dim VAR_BIEN, VAR_R As String
Dim VAR_N1, VAR_N2, VAR_N3, VAR_POA As String
Dim VAR_LITERAL1, VAR_LITERAL2 As String

Dim FInicio, FFin, FControl, FVenta As Date
Dim precio_tot, precio_uni As Double


Private Sub CmdDetalle_Click()
'    FrmCobranza.Visible = True
End Sub

'Private Sub adosalalm_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    If pRecordset.EOF Or pRecordset.BOF Then Exit Sub
'        Select Case pRecordset.EditMode
'        Case adEditNone
'            If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'            rstdetsalalm.Open "Select * from ao_detallesalidaalmacen where correlativo_salida = '" & pRecordset("correlativo_salida") & "'", db, adOpenDynamic, adLockOptimistic
'            Set DataGrid2.DataSource = Nothing
'            Set DataGrid2.DataSource = rstdetsalalm
'            DataGrid2.ReBind
'        End Select
'End Sub

Private Sub Adodetallesolicitud_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If (Not adoDetalleSolicitud.Recordset.BOF) And (Not adoDetalleSolicitud.Recordset.EOF) Then
        If Not IsNull(adoDetalleSolicitud.Recordset("correlativo_solicitud")) Then
            txtnosolicitud1.Text = adoDetalleSolicitud.Recordset("correlativo_solicitud")
            txtcorrdet.Text = adoDetalleSolicitud.Recordset("correlativo_detalle")
        Else
            txtnosolicitud1.Text = Ado_datos.Recordset("codigo_solicitud")
            txtcorrdet.Text = " "
            dtccodpar.Text = " "
            dtcdescripar.Text = " "
            txtsolpeso.Text = 0
        End If
    End If
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  Dim descri_bien As String
  Dim Cant_Alm As Integer
  If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then

    If parametro <> Ado_datos.Recordset!unidad_codigo Then
'        BtnAnlDetalle.Visible = False
    Else
'        BtnAnlDetalle.Visible = True
    End If
    If Not IsNull(Ado_datos.Recordset!IdTraspasoBancos) Then
        If buscados = 0 Then
           OptFilGral1.Visible = True
           OptFilGral2.Visible = True
        Else
           OptFilGral1.Visible = False
           OptFilGral2.Visible = False
        End If
        If (Ado_datos.Recordset!estado_verificado = "APR") And (Ado_datos.Recordset!estado_codigo = "APR") And (Ado_datos.Recordset!estado_conciliado = "APR") Then
            BtnAprobar1.Visible = False
            BtnAprobar.Visible = False
            BtnModificar.Visible = False
            BtnDesAprobar.Visible = False
            BtnEliminar.Visible = False
            lbl_cerrado.Caption = "CONCILIADO CONTABILIDAD"
            FrmABMDet.Visible = False
            FrmABMDet2.Visible = False
            FrmDetalle.Visible = False
        End If
        If (Ado_datos.Recordset!estado_verificado = "APR") And (Ado_datos.Recordset!estado_codigo = "APR") And (Ado_datos.Recordset!estado_conciliado = "REG") Then
            BtnAprobar1.Visible = False
            BtnAprobar.Visible = False
            BtnModificar.Visible = False
            BtnDesAprobar.Visible = True
            BtnEliminar.Visible = False
            lbl_cerrado.Caption = "APROBADO SUPERVISOR"
            FrmABMDet.Visible = False
            FrmABMDet2.Visible = False
            FrmDetalle.Visible = False
        End If
        If (Ado_datos.Recordset!estado_verificado = "APR") And (Ado_datos.Recordset!estado_codigo = "REG") And (Ado_datos.Recordset!estado_conciliado = "REG") Then
            BtnAprobar1.Visible = False
            BtnAprobar.Visible = True
            BtnModificar.Visible = False
            BtnDesAprobar.Visible = True
            BtnEliminar.Visible = False
            lbl_cerrado.Caption = "VERIFICADO TESORERIA"
            FrmABMDet.Visible = False
            FrmABMDet2.Visible = False
            FrmDetalle.Visible = False
        End If
        If (Ado_datos.Recordset!estado_verificado = "REG") And (Ado_datos.Recordset!estado_codigo = "REG") And (Ado_datos.Recordset!estado_conciliado = "REG") Then
            BtnAprobar1.Visible = True
            BtnAprobar.Visible = False
            BtnModificar.Visible = True
            BtnDesAprobar.Visible = False
            BtnEliminar.Visible = False
            lbl_cerrado.Caption = ""
            FrmABMDet.Visible = True
            FrmABMDet2.Visible = True
            FrmDetalle.Visible = True
        End If
        
        Call AbrirDetalle
        
'        FrmDetalle.Caption = "ORIGEN - Ordenes de Pago a Proveedores " '+ Str((IIf(IsNull(Ado_datos.Recordset!correl_doc), 0, Ado_datos.Recordset!correl_doc)))
        FrmDetalle2.Caption = "DESTINO del TRASPASO Nro. " + Str((IIf(IsNull(Ado_datos.Recordset!correl_doc), 0, Ado_datos.Recordset!correl_doc)))
    End If
        'FrmDetalle.Visible = True
'        FrmCobranza.Visible = True
  Else
        FrmABMDet.Visible = False
        FrmABMDet2.Visible = False
        FrmDetalle.Visible = False
        If buscados = 0 Then
           OptFilGral1.Visible = True
           OptFilGral2.Visible = True
        Else
           OptFilGral1.Visible = False
           OptFilGral2.Visible = False
        End If
  End If
'        BtnEliminar.Visible = True
End Sub

Private Sub AbrirOrigen()
    'ORIGEN
    Set rs_datos14 = New ADODB.Recordset
    If rs_datos14.State = 1 Then rs_datos14.Close
    'If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        ' fv_tes_adjudica_recibos   'ESTADO_CODIGO='APR'
        DtGLista.Visible = True
'        If dtc_moneda22.Text = "" Then
'            queryinicial2 = "select * from fv_tes_adjudica_recibos_trp where IdTraspasoBancos = '0' AND estado_codigo = 'REG' AND estado_destino = 'REG' "      'beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND
'        Else
'            queryinicial2 = "select * from fv_tes_adjudica_recibos_trp where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND beneficiario_codigo_prov = '" & Ado_datos.Recordset!beneficiario_codigo_prov & "' AND IdTraspasoBancos = '0' AND estado_codigo = 'APR' AND estado_destino = 'REG' "
'        End If
'        ', db, adOpenKeyset, adLockOptimistic
'    Else
'        ' fv_tes_adjudica_recibos   'ESTADO_DESTINO='APR'
'        DtGLista.Visible = False
'        queryinicial2 = "select * from fv_tes_adjudica_recibos_trp where estado_codigo_tes = 'APR' AND (IdTraspasoBancos is NULL or IdTraspasoBancos ='0') and estado_codigo_cont = 'REG'     "
'        'order by  doc_numero
'    End If
    
    'DESDE AQUI
    queryinicial2 = "select * from fv_tes_adjudica_recibos_trp where IdTraspasoBancos = '0' AND estado_codigo = 'APR' AND estado_destino = 'REG' "
    rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
    rs_datos14.Sort = "IdRecibo, Observaciones"
    'HASTA AQUI
    Set ado_datos14.Recordset = rs_datos14.DataSource
    ado_datos14.Recordset.Requery
    If ado_datos14.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista.Visible = True
        Set DtGLista.DataSource = ado_datos14.Recordset
        'Call AbreOrigen
    Else
        deta2 = 0
        DtGLista.Visible = False
    End If
End Sub

Private Sub AbrirDetalle()
'    'ORIGEN
'    Set rs_datos14 = New ADODB.Recordset
'    If rs_datos14.State = 1 Then rs_datos14.Close
'    If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
'        ' fv_tes_adjudica_recibos   'ESTADO_CODIGO='APR'
'        DtGLista.Visible = True
'        If dtc_moneda22.Text = "" Then
'            queryinicial2 = "select * from fv_tes_adjudica_recibos_trp where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND IdTraspasoBancos = '0' AND estado_codigo = 'APR' AND estado_destino = 'REG' "
'        Else
'            queryinicial2 = "select * from fv_tes_adjudica_recibos_trp where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND beneficiario_codigo_prov = '" & Ado_datos.Recordset!beneficiario_codigo_prov & "' AND IdTraspasoBancos = '0' AND estado_codigo = 'APR' AND estado_destino = 'REG' "
'        End If
'        'queryinicial2 = "select * from fv_tes_adjudica_recibos_trp where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo_resp & "' AND IdTraspasoBancos = '0' AND estado_codigo = 'APR' AND estado_destino = 'REG' "
'        ', db, adOpenKeyset, adLockOptimistic
'    Else
'        ' fv_tes_adjudica_recibos   'ESTADO_DESTINO='APR'
'        DtGLista.Visible = False
'        queryinicial2 = "select * from fv_tes_adjudica_recibos_trp where estado_codigo_tes = 'APR' AND (IdTraspasoBancos is NULL or IdTraspasoBancos ='0') and estado_codigo_cont = 'REG'     "
'        'order by  doc_numero
'    End If
'    'DESDE AQUI
'    rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
'    rs_datos14.Sort = "doc_numero"
'    'HASTA AQUI
'    Set ado_datos14.Recordset = rs_datos14.DataSource
'    ado_datos14.Recordset.Requery
'    If ado_datos14.Recordset.RecordCount > 0 Then
'        deta2 = 1
'        DtGLista.Visible = True
'        Set DtGLista.DataSource = ado_datos14.Recordset
'        Call AbreOrigen
'    Else
'        deta2 = 0
'        DtGLista.Visible = False
'    End If
    
    'DESTINO - DETALLE DEL RECIBO
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    If Ado_datos.Recordset!estado_codigo = "REG" Or IsNull(Ado_datos.Recordset!estado_codigo) Then
        BtnModDetalle.Visible = True
        BtnAnlDetalle.Visible = True
        rs_datos11.Open "select * from fv_tes_adjudica_recibos_trp where IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & " AND estado_codigo = 'APR' AND estado_destino = 'APR' order by  doc_numero ", db, adOpenKeyset, adLockOptimistic
    Else
        BtnModDetalle.Visible = False
        BtnAnlDetalle.Visible = False
        rs_datos11.Open "select * from fv_tes_adjudica_recibos_trp where IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & " order by  doc_numero ", db, adOpenKeyset, adLockOptimistic
    End If
    'rs_datos11.Sort = "doc_numero "
    Set Ado_datos11.Recordset = rs_datos11.DataSource
    Ado_datos11.Recordset.Requery
    If Ado_datos11.Recordset.RecordCount > 0 Then
        deta2 = 1
        DtGLista11.Visible = True
        Set DtGLista11.DataSource = Ado_datos11.Recordset
    Else
        deta2 = 0
        DtGLista11.Visible = False
    End If
End Sub

Private Sub AbreGrupoRecibo()
    'ORIGEN RECIBOS OFICIALES DETALLE
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "select * from fv_tes_orden_pago_pendientes_agrupados WHERE (unidad_codigo_resp = '" & ado_datos14.Recordset!unidad_codigo_adm & "')   ", db, adOpenKeyset, adLockOptimistic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub BtnAddDetalle_Click()
 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_verificado = "REG" Then
        If glusuario = "ASANTIVAÑEZ" Or glusuario = "TCASTILLO" Or glusuario = "FCABRERA" Or glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "VPAREDES" Or glusuario = "MWILDE" Or glusuario = "MVALDIVIA" Or glusuario = "EVILLALOBOS" Or glusuario = "CSALINAS" Then
           FraDet3.Visible = True
           Call AbreGrupoRecibo
           FraNavega.Enabled = False
           FrmDetalle.Enabled = False
           FrmABMDet.Enabled = False
           FrmABMDet2.Visible = False
           FrmDetalle2.Enabled = False
           fraOpciones.Enabled = False
         Else
               MsgBox "El USUARIO no tiene Acceso !!. Consulte con el Administrador del Sistema. ", vbExclamation, "Atención!"
         End If
    Else
        MsgBox "El Traspaso ya fue VERIFICADO o APROBADO !!. Elija otro Traspaso o elabore uno Nuevo... ", vbExclamation, "Atención!"
    End If
 Else
    MsgBox "No existe un Traspaso habilitado!!. Elabore uno Nuevo... ", vbExclamation, "Atención!"
 End If
End Sub

Private Sub BtnAñadir_Click()
accion = "NEW"
On Error GoTo UpdateErr
  If glusuario = "TCASTILLO" Or glusuario = "FCABRERA" Or glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "MWILDE" Or glusuario = "MVALDIVIA" Or glusuario = "EVILLALOBOS" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then
    'Ado_datos.Recordset.AddNew
    VAR_R = "R-644"
    dtc_codigo3.Text = VAR_R
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
    
    'dtc_desc3.backColor = &H80000005
    'dtc_desc3.ForeColor = &H80000008
    
    'txt_campo1.Caption = "0"
    'dtc_desc3.Locked = False
    'dtc_desc3.Width = 5955
    
    'DTPfechasol.Value = Date
    swgrabar = 1
    FrmCabecera.Visible = True
    FrmDetalle.Visible = False
    FraNavega.Enabled = False
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    Fra_datos.Enabled = True

    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
''    SSTab1.TabEnabled(1) = False
    'dtc_desc4.SetFocus
  Else
        MsgBox "El USUARIO no tiene Acceso !!. Consulte con el Administrador del Sistema. ", vbExclamation, "Atención!"
  End If
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub BtnAprobar_Click()
 On Error GoTo UpdateErr
  If (Ado_datos.Recordset!estado_codigo = "REG") And (Ado_datos.Recordset!estado_verificado = "APR") And (glusuario = "RCUELA" Or glusuario = "ADMIN" Or glusuario = "CSALINAS") Then
    VAR_RECIBO = Ado_datos.Recordset!IdTraspasoBancos
    'Actualiza Totales
    db.Execute "UPDATE fo_traspaso_bancos_Egresos set fo_traspaso_bancos_Egresos.total_bs  = fv_tes_recibos_detalle_sum_egreso.adjudica_bs, fo_traspaso_bancos_Egresos.total_dol   = fv_tes_recibos_detalle_sum_egreso.adjudica_dol from fo_traspaso_bancos_Egresos inner join fv_tes_recibos_detalle_sum_egreso " & _
        " on fo_traspaso_bancos_Egresos.IdTraspasoBancos  = fv_tes_recibos_detalle_sum_egreso.IdTraspasoBancos where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & VAR_RECIBO & "  "

    'Actualiza Detalle cta_codigo_origen Y cta_codigo_destino
    db.Execute "UPDATE fo_recibos_detalle_egresos set fo_recibos_detalle_egresos.cta_codigo_origen = fo_traspaso_bancos_Egresos.cta_codigo, fo_recibos_detalle_egresos.cta_codigo_destino  = fo_traspaso_bancos_Egresos.cta_codigo_destino FROM fo_recibos_detalle_egresos INNER JOIN fo_traspaso_bancos_Egresos " & _
        " ON fo_recibos_detalle_egresos.IdTraspasoBancos = fo_traspaso_bancos_Egresos.IdTraspasoBancos where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & VAR_RECIBO & "  "

    'Actualiza Detalle estado_aprueba
    db.Execute "update fo_recibos_detalle_egresos set fo_recibos_detalle_egresos.estado_aprueba = 'APR', fecha_aprueba= '" & Date & "' WHERE fo_recibos_detalle_egresos.IdTraspasoBancos =  " & VAR_RECIBO & "  "
    
    'APRUEBA ao_ventas_cobranza_det estado_codigo_concilia
    db.Execute "update ao_ventas_cobranza_det set ao_ventas_cobranza_det.estado_codigo_concilia = 'APR' from ao_ventas_cobranza_det inner join fo_recibos_detalle_egresos on ao_ventas_cobranza_det.adjudica_codigo = fo_recibos_detalle_egresos.adjudica_codigo WHERE fo_recibos_detalle_egresos.IdTraspasoBancos =  " & VAR_RECIBO & "  "
'fecha_destino
    'APRUEBA fo_traspaso_bancos_Egresos
    db.Execute "update fo_traspaso_bancos_Egresos set correl_doc = IdRecibo, estado_codigo = 'APR', usr_codigo_aprueba = '" & glusuario & "', fecha_registro_aprueba = '" & Date & "'  where IdTraspasoBancos = " & VAR_RECIBO & " "
    
    'CONTABILIZA COBRANZAS -----------------------------------------------
    Call Contabiliza_Pago(VAR_RECIBO, glusuario)
    
    OptFilGral2_Click
    
    If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
    End If
    If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "IdTraspasoBancos = " & VAR_RECIBO & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
        ' If rs_det1.RecordCount > 0 Then
        ' rs_det1.MoveLast
        'End If
    Else
        rs_datos.MoveLast
    End If
    
  Else
    MsgBox "No se puede aprobar el registro actual"
  End If
Exit Sub
UpdateErr:
MsgBox Err.Description

End Sub

'APRUEBA fo_traspaso_bancos_Egresos
'db.Execute "update fo_traspaso_bancos_Egresos set estado_verificado = 'APR', usr_codigo_verificado = '" & glusuario & "', fecha_verificado = '" & Date & "'  where IdTraspasoBancos = " & VAR_RECIBO & " "
'Call Contabiliza_Cobranzas

Private Sub Contabiliza_Cobranzas()
    ' Contabilizacion al momento de aprobacion
    'Base de datos
    Dim db2 As New ADODB.Connection
    ' Recordset
    Dim rs_aux99 As New ADODB.Recordset  ' Data
    Dim rs_aux100 As New ADODB.Recordset ' Transaccion
    Dim rs_aux101 As New ADODB.Recordset ' Beneficiarios
    Dim rs_aux102 As New ADODB.Recordset ' Edificios
    Dim rs_aux103 As New ADODB.Recordset ' Moneda de cuenta corriente
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    Dim VAR_EMPRESA As Integer
    Dim VAR_PARTIDA As String
    Dim VAR_TIPOCOMPID As Integer
    Dim VAR_FECHA As Date
    Dim VAR_MONEDAID As Integer
    Dim VAR_TIPOCAMBIO As Double
    Dim EntregadoA As String
    Dim VAR_CONCEPTO As String
    Dim VAR_GLOSA As String
    Dim VAR_DEBEORG As Double
    Dim VAR_HABERORG As Double
    'Impuestos
    Dim VAR_PorIVA As Double
    Dim VAR_PorIT As Double
    Dim VAR_PorITF As Double
    'Otros valores
    Dim VAR_ConFac As Integer
    Dim VAR_SinFac As Integer
    Dim VAR_Automatico As Integer
    Dim VAR_TipoNotaId As Integer
    Dim VAR_NotaNro As String
    Dim VAR_EstadoId As Integer
    Dim VAR_iConcurrency_id As Integer
    Dim VAR_TipoAsientoId As Integer
    Dim VAR_CentroCostoId As Integer
    Dim VAR_TipoRetencionId As Integer
    Dim VAR_TipoId As Integer
    Dim VAR_CompDetIdOrg As Integer
    Dim VAR_PROY2 As String
    ' Data (AdoDb - fuente de datos)
    Set rs_aux99 = New ADODB.Recordset
    If rs_aux99.State = 1 Then rs_aux99.Close
    rs_aux99.Open "SELECT trans, tipoV, depto, fecha, bs2, dol2, beneficiario, edifCodCorto, edifCodigo, cuentaDestino, glosa, solicitudTipo, notaNro, codBancarizacion FROM av_contabiliza_cobranzas WHERE fecha IS NOT NULL AND IdTraspasoBancos = '" & Ado_datos.Recordset!IdTraspasoBancos & "' ORDER BY fecha ", db, adOpenKeyset, adLockOptimistic
    'Codigo tipo = "REC"
    VAR_CODTIPO = "REC"
    ' ==================
    ' Ciclo Inicio
    ' ==================
    If rs_aux99.RecordCount > 0 Then
        rs_aux99.MoveFirst
    End If
    Do While rs_aux99.EOF = False
        ' Rubro Codigo/Centro de costo
        Set rs_aux100 = New ADODB.Recordset
        If rs_aux100.State = 1 Then rs_aux100.Close
        rs_aux100.Open "SELECT trans_descripcion, rubro_codigo, CentroCostoId FROM gc_tipo_transaccion WHERE trans_codigo = '" & rs_aux99!trans & "'  ", db, adOpenKeyset, adLockOptimistic
        If rs_aux100.RecordCount > 0 Then
            VAR_PARTIDA = rs_aux100!rubro_codigo
            VAR_CentroCostoId = rs_aux100!CentroCostoId
        End If
        ' Empresa
        ' Solo si es: Ventas CGE => venta_tipo = 'G'
        If rs_aux99!tipoV = "G" Then
            VAR_EMPRESA = 2
        Else
            VAR_EMPRESA = 1
        End If
        'Departamento/Sucursal/Region
        VAR_DPTO = rs_aux99!Depto
        ' Fecha de venta
        VAR_FECHA = CDate(rs_aux99!Fecha)
        ' Tipo moneda/Debe/Haber (Adicionar)
        'VAR_MONEDAID = 1
        'VAR_DEBEORG = rs_aux99!bs2 'Boliviano
        'VAR_HABERORG = rs_aux99!bs2 'Boliviano
        Set rs_aux103 = New ADODB.Recordset
        If rs_aux103.State = 1 Then rs_aux103.Close
        rs_aux103.Open "SELECT MonedaId FROM tblCuentaCorriente WHERE TipoIE = 1 AND Cuenta = '" & rs_aux99!cuentaDestino & "' ", db, adOpenKeyset, adLockOptimistic
        If rs_aux103.RecordCount > 0 Then
            VAR_MONEDAID = rs_aux103!MonedaId
        End If
        If VAR_MONEDAID = 2 Then
            VAR_DEBEORG = rs_aux99!dol2 'Dolar
            VAR_HABERORG = rs_aux99!dol2 'Dolar
        Else
            VAR_DEBEORG = rs_aux99!bs2 'Boliviano
            VAR_HABERORG = rs_aux99!bs2 'Boliviano
        End If
        ' Tipo de cambio -> BOB - USD
        VAR_TIPOCAMBIO = Round(rs_aux99!bs2 / rs_aux99!dol2, 2)
        'MMMMMMMMMMMMMMMMMMMMMMMMMMod
        ' Entregado A:
        'Set rs_aux101 = New ADODB.Recordset
        'If rs_aux101.State = 1 Then rs_aux101.Close
        'rs_aux101.Open "SELECT beneficiario_nit, beneficiario_denominacion FROM gc_beneficiario WHERE beneficiario_codigo = '" & rs_aux99!beneficiario & "'  ", db, adOpenKeyset, adLockOptimistic
        'EntregadoA = rs_aux99!beneficiario
        'If rs_aux101.RecordCount > 0 Then
        '    EntregadoA = EntregadoA & " - " & rs_aux101!beneficiario_denominacion
        'End If
        ' Entregado A:
        EntregadoA = "Edificio " & rs_aux99!edifCodCorto
        If rs_aux102.State = 1 Then rs_aux102.Close
        rs_aux102.Open "select edif_descripcion from gc_edificaciones where edif_codigo = '" & rs_aux99!edifCodigo & "'  ", db, adOpenKeyset, adLockOptimistic
        If rs_aux102.RecordCount > 0 Then
            EntregadoA = rs_aux102!edif_descripcion & " - " & EntregadoA
        End If
        'Por concepto
        VAR_CONCEPTO = "Contabilizacion cobranza: " & rs_aux99!notaNro & " - Codigo de Bancarizacion: " & rs_aux99!codBancarizacion & " - " & rs_aux99!glosa
        'VAR_CONCEPTO = rs_aux99!glosa
        ' Glosa
        'VAR_GLOSA = "Codigo de Bancarizacion: " & rs_aux99!codBancarizacion
        VAR_GLOSA = rs_aux99!glosa
        'MMMMMMMMMMMMMMMMMMMMMMMMMMod
        VAR_PROY2 = rs_aux99!edifCodigo
        ' TipoCompId (Tipo comprobante id) Ingreso
        VAR_TIPOCOMPID = 1
        ' Impuestos
        VAR_PorIVA = 0.13
        VAR_PorIT = 0.03
        VAR_PorITF = 0.0015
        ' Otros valores
        VAR_ConFac = 0 'Con factura
        VAR_SinFac = 1 'Sin factura
        VAR_Automatico = 1 '0 Permite edicion, 1 no permite editar
        VAR_TipoNotaId = rs_aux99!solicitudTipo
        VAR_NotaNro = rs_aux99!notaNro
        VAR_EstadoId = 11 'Libro Mayor requiere que sean de EstadoId = 10 Cerrado OR EstadoId = 11 Abierto
        VAR_TipoAsientoId = 0 ' Operativo
        VAR_TipoRetencionId = 0
        VAR_TipoId = 0
        VAR_CompDetIdOrg = 0
        'Procedimiento almacenado
        ' Creamos conexion unica para CONDOBO
        db2.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CONDOBO;Data Source=SSOFIA"
        db2.Execute ("EXEC fp_contabiliza_ingresos '" & VAR_CODTIPO & "', '" & VAR_PARTIDA & "', " & VAR_EMPRESA & ", " & VAR_DPTO & ", " & VAR_TIPOCOMPID & ", '" & VAR_FECHA & "', " & VAR_MONEDAID & ", '" & VAR_TIPOCAMBIO & "', '" & VAR_DEBEORG & "', '" & VAR_HABERORG & "', '" & EntregadoA & "', '" & VAR_CONCEPTO & "', '" & VAR_PorIVA & "', '" & VAR_PorIT & "', '" & VAR_PorITF & "', " & VAR_ConFac & ", " & VAR_SinFac & ", " & VAR_Automatico & ", '" & VAR_GLOSA & "', " & VAR_TipoNotaId & ", " & VAR_NotaNro & ", " & VAR_EstadoId & ", '" & glusuario & "', " & VAR_TipoAsientoId & ", " & VAR_CentroCostoId & ", " & VAR_TipoRetencionId & ", " & VAR_TipoId & ", " & VAR_CompDetIdOrg & ", '" & VAR_PROY2 & "'")
        db2.Close
        ' Siguiente
        rs_aux99.MoveNext
    Loop
    If rs_aux99.State = 1 Then rs_aux99.Close
    If rs_aux100.State = 1 Then rs_aux100.Close
    If rs_aux101.State = 1 Then rs_aux101.Close
    If rs_aux102.State = 1 Then rs_aux102.Close
    If rs_aux103.State = 1 Then rs_aux103.Close
End Sub
Private Sub GENERA_COMPRA()
'    If rs_datos!estado_cotiza = "REG" Then
'      VAR_COD4 = Ado_datos.Recordset!unidad_codigo
'      VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'      VAR_PROY2 = Ado_datos.Recordset!edif_codigo
'      VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'        ' MANTENIMIENTO PREVENTIVO - INSUMOS y/o COMPRAS BB y SS
'                'EQUIPO
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    rs_aux2.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & parametro & "'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux2.RecordCount > 0 Then
'                       rs_aux2!correl_negocia = rs_aux2!correl_negocia + 1
'                       correldetalle = rs_aux2!correl_negocia
'                       rs_aux2.Update
'                    End If
'                    'WWWWWWWWWWWWWWW
'                    'correlv = Ado_datos.Recordset!venta_codigo
'                    'VAR_TIPOV = Ado_datos.Recordset!venta_tipo
'
'                    Set rs_aux3 = New ADODB.Recordset
'                    If rs_aux3.State = 1 Then rs_aux3.Close
'                    rs_aux3.Open "select * from ao_compra_cabecera where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo = " & VAR_SOL & " ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux3.RecordCount = 0 Then
'                    'beneficiario_codigo_resp,'doc_numero,estado_codigo_tra, estado_codigo_nac, estado_codigo_des, hora_registro, usr_codigo_aprueba,'                      fecha_registro_aprueba
'                        rs_aux3.AddNew
'                        rs_aux3!ges_gestion = glGestion     'Year(Date)
'                        'rs_aux3!compra_codigo = 0      'Autonumerico
'                        rs_aux3!unidad_codigo_adm = parametro
'                        rs_aux3!solicitud_codigo_adm = correldetalle
'                        rs_aux3!unidad_codigo = VAR_COD4
'                        rs_aux3!solicitud_codigo = VAR_SOL
'                        rs_aux3!edif_codigo = VAR_PROY2
'                        rs_aux3!beneficiario_codigo = VAR_BENEF
'                        rs_aux3!solicitud_tipo = Ado_datos.Recordset!solicitud_tipo       '"10"
'                        rs_aux3!venta_tipo = "E"
'                        rs_aux3!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant   'VAR_CITE
'                        rs_aux3!compra_fecha = Date
'                        rs_aux3!compra_descripcion = "COMPRA POR: " + lbl_titulo.Caption
'                        rs_aux3!compra_observaciones = "Edificio: " + Trim(dtc_desc3.Text)
'                        rs_aux3!compra_cantidad_total = 1   'Ado_datos.Recordset!venta_cantidad_total
'                        rs_aux3!compra_monto_bs = 0     'VAR_BS2
'                        rs_aux3!tipo_moneda = "BOB"
'                        rs_aux3!compra_monto_dol = 0        'VAR_DOL2
'                        rs_aux3!proceso_codigo = "TEC"
'                        rs_aux3!subproceso_codigo = "TEC-06"
'                        rs_aux3!etapa_codigo = "TEC-06-01"
'                        rs_aux3!clasif_codigo = "ADM"
'                        rs_aux3!doc_codigo = "R-114"
'                        rs_aux3!poa_codigo = "3.2.8"
'                        rs_aux3!estado_codigo_eqp = "REG"
'                        rs_aux3!estado_codigo = "REG"
'                        rs_aux3!usr_codigo = glusuario
'                        rs_aux3!fecha_registro = Date
'                        rs_aux3.Update
'
'                        'DETALLE Carga ao_ventas_detalle
'                        Set rstdestino = New ADODB.Recordset
'                        If rstdestino.State = 1 Then rstdestino.Close
'                        rstdestino.Open "select * from ao_compra_detalle  ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rstdestino.RecordCount > 0 Then
'                        End If
'                        Set rs_aux4 = New ADODB.Recordset
'                        If rs_aux4.State = 1 Then rs_aux4.Close
'                        'rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo= " & rs_aux3!compra_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
'                        rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo= " & VAR_SOL & "  and grupo_codigo = '30000' ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rs_aux4.RecordCount > 0 Then
'                            VAR_REG = 1
'                           rs_aux4.MoveFirst
'                           While Not rs_aux4.EOF
'                              If rs_aux4!grupo_codigo = "30000" Then
'                                db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, compra_codigo_det, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigoR , usr_usuario, fecha_registro) " & _
'                                "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", " & VAR_REG & ", '" & rs_aux4!bien_codigo & "', " & rs_aux4!bien_cantidad & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", '" & rs_aux3!compra_descripcion & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
'
'                                db.Execute "Update ao_compra_detalle SET ao_compra_detalle.compra_concepto  = ac_bienes.bien_descripcion From ao_compra_detalle INNER JOIN ac_bienes ON ao_compra_detalle.bien_codigo = ac_bienes.bien_codigo where ao_compra_detalle.compra_codigo = " & rs_aux3!compra_codigo & " and ao_compra_detalle.bien_codigo = '" & rs_aux4!bien_codigo & "' "
'                                VAR_REG = VAR_REG + 1
'                              End If
'                               rs_aux4.MoveNext
'                           Wend
'                        End If
'                        If rstdestino.State = 1 Then rstdestino.Close
'                    End If
'                    'WWWWWWWWWW
'        Set rs_aux2 = New ADODB.Recordset
'        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9 & "'  "
'        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'            Txt_campo1.Caption = rs_aux2!correl_doc
'            rs_aux2.Update
'        End If
'        rs_datos!doc_numero = Txt_campo1.Caption
'        'REVISAR !!! JQA 2014_07_08
'        'VAR_ARCH = RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
'        VAR_ARCH = "COM_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(Txt_campo1.Caption)))
'        rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
'        rs_datos!archivo_respaldo_cargado = "N"
'        rs_datos!estado_cotiza = "APR"
'        rs_datos!fecha_aprueba = Date
'        rs_datos!usr_codigo_aprueba = glusuario
'        rs_datos.UpdateBatch adAffectAll
'      End If
'
'  Else
'      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
End Sub

Private Sub BtnAprobar1_Click()
 On Error GoTo UpdateErr
  If (Ado_datos.Recordset!estado_verificado = "REG") And (glusuario = "ADMIN" Or glusuario = "TCASTILLO" Or glusuario = "FCABRERA" Or glusuario = "VPAREDES" Or glusuario = "MWILDE" Or glusuario = "MVALDIVIA" Or glusuario = "EVILLALOBOS" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS") Then
    VAR_RECIBO = Ado_datos.Recordset!IdTraspasoBancos
    'Actualiza Totales
    db.Execute "UPDATE fo_traspaso_bancos_Egresos set fo_traspaso_bancos_Egresos.total_bs  = fv_tes_recibos_detalle_sum_egreso.adjudica_bs, fo_traspaso_bancos_Egresos.total_dol   = fv_tes_recibos_detalle_sum_egreso.adjudica_dol from fo_traspaso_bancos_Egresos inner join fv_tes_recibos_detalle_sum_egreso " & _
        " on fo_traspaso_bancos_Egresos.IdTraspasoBancos  = fv_tes_recibos_detalle_sum_egreso.IdTraspasoBancos where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & VAR_RECIBO & "  "

    'Actualiza Detalle cta_codigo_origen Y cta_codigo_destino
    db.Execute "UPDATE fo_recibos_detalle_egresos set fo_recibos_detalle_egresos.cta_codigo_origen = fo_traspaso_bancos_Egresos.cta_codigo, fo_recibos_detalle_egresos.cta_codigo_destino  = fo_traspaso_bancos_Egresos.cta_codigo_destino FROM fo_recibos_detalle_egresos INNER JOIN fo_traspaso_bancos_Egresos " & _
        " ON fo_recibos_detalle_egresos.IdTraspasoBancos = fo_traspaso_bancos_Egresos.IdTraspasoBancos where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & VAR_RECIBO & "  "

    'Actualiza Detalle estado_aprueba wwwwwwwwwwwwwwwwwwwwwww
    db.Execute "update fo_recibos_detalle_egresos set fo_recibos_detalle_egresos.estado_aprueba = 'APR', fecha_aprueba= '" & Date & "' WHERE fo_recibos_detalle_egresos.IdTraspasoBancos =  " & VAR_RECIBO & "  "
    
    'APRUEBA ao_ventas_cobranza_det estado_codigo_concilia wwwwwwwwwwwwwwwwwwwww
    'db.Execute "update ao_ventas_cobranza_det set ao_ventas_cobranza_det.estado_codigo_concilia = 'APR' from ao_ventas_cobranza_det inner join fo_recibos_detalle_egresos on ao_ventas_cobranza_det.adjudica_codigo = fo_recibos_detalle_egresos.adjudica_codigo WHERE fo_recibos_detalle_egresos.IdTraspasoBancos =  " & VAR_RECIBO & "  "
    'fecha_destino
    
    'APRUEBA fo_traspaso_bancos_Egresos
    db.Execute "update fo_traspaso_bancos_Egresos set estado_verificado = 'APR', usr_codigo_verificado = '" & glusuario & "', fecha_verificado = '" & Date & "'  where IdTraspasoBancos = " & VAR_RECIBO & " "
    
    OptFilGral2_Click
    
    If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
    End If
    If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "IdTraspasoBancos = " & VAR_RECIBO & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
        ' If rs_det1.RecordCount > 0 Then
        ' rs_det1.MoveLast
        'End If
    Else
        rs_datos.MoveLast
    End If
    
  Else
    MsgBox "El Usuario NO tiene Permiso o, el registro actual ya fue Verificado o Anulado, verifique el estado !!"
  End If
Exit Sub
UpdateErr:
MsgBox Err.Description
End Sub

Private Sub BtnBuscar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      buscados = 1
      PosibleApliqueFiltro = False
      
      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexión = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = dg_datos
      ClBuscaGrid.QueryUtilizado = queryinicial
      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
    OptFilGral1.Visible = True
    OptFilGral2.Visible = True
  End If
End Sub

Private Sub BtnBuscar1_Click()
  If ado_datos14.Recordset.RecordCount > 0 Then
    'JQA
      buscados = 1
      PosibleApliqueFiltro = False

      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexión = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = DtGLista
      ClBuscaGrid.QueryUtilizado = queryinicial2
      Set ClBuscaGrid.RecordsetTrabajo = ado_datos14.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
    'OptFilGral1.Visible = True
    'OptFilGral2.Visible = True
  End If
    
'    SWFILTRO = 1
'    FraDet3.Visible = True
End Sub

Private Sub BtnCancelar_Click()
On Error GoTo UpdateErr
  If swgrabar = 2 Then
    var_cod5 = Ado_datos.Recordset!venta_codigo
  End If
  'Ado_datos.Refresh
  fraOpciones.Visible = True
  FraGrabarCancelar.Visible = False
  'marca1 = Ado_datos.Recordset.Bookmark
  FraNavega.Enabled = True
  FrmCabecera.Visible = False
  Fra_datos.Enabled = True
  FrmDetalle.Visible = True

'  Fra_Total.Visible = True
  dg_datos.Visible = True
  FrmABMDet.Visible = True
  FrmABMDet2.Visible = True
  dtc_desc3.backColor = &H80000008
  dtc_desc3.ForeColor = &H80000005
  
  'Refrescar Grid
  If OptFilGral1.Value = True Then
       Call OptFilGral1_Click        'Pendientes
  Else
       Call OptFilGral2_Click        'TODOS
  End If
  If (dg_datos.SelBookmarks.Count <> 0) Then
       dg_datos.SelBookmarks.Remove 0
  End If
  If Ado_datos.Recordset.RecordCount > 0 And swgrabar = 2 Then
       rs_datos.Find "venta_codigo = " & var_cod5 & "   ", , , 1
       dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
  Else
       rs_datos.MoveLast
  End If
  swgrabar = 0
'  SSTab1.Tab = 0
'  SSTab1.TabEnabled(0) = True
'  SSTab1.TabEnabled(1) = True
  accion = ""
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub BtnCancelar1_Click()
    FraDet3.Visible = False
    
    FraNavega.Enabled = True
    FrmDetalle.Enabled = True
    FrmABMDet.Enabled = True
    FrmABMDet2.Visible = True
    FrmDetalle2.Enabled = True
    fraOpciones.Enabled = True
    SWFILTRO = 0
    VARFILTRO = 0
    Call AbrirDetalle
End Sub

Private Sub BtnCancelar2_Click()
On Error GoTo UpdateErr
If glusuario = "TCASTILLO" Or glusuario = "FCABRERA" Or glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "MWILDE" Or glusuario = "MVALDIVIA" Or glusuario = "EVILLALOBOS" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then
 'If Ado_datos.Recordset.RecordCount > 0 Then
    'If Ado_datos.Recordset!estado_codigo = "REG" Then
        If ado_datos14.Recordset.RecordCount > 0 Then         '<> "" Then
            Set rs_datos7 = New ADODB.Recordset
            If rs_datos7.State = 1 Then rs_datos7.Close
            rs_datos7.Open "select * from fv_tes_adjudica_recibos_trp WHERE (IdRecibo = " & dtc_codigo6.Text & " AND estado_destino = 'REG')  ", db, adOpenKeyset, adLockOptimistic
            If rs_datos7.RecordCount > 0 Then
                rs_datos7.MoveFirst
                While Not rs_datos7.EOF
                    If (IsNull(rs_datos7!adjudica_fecha) Or (rs_datos7!adjudica_fecha = "01/01/1900")) Then         '(rs_datos7!trans_codigo <> "E") And
                        MsgBox "No se puede enviar a DESTINO, verifique la fecha de Factura o Recibo y vuelva a intentar ...", , "Atención"
                        FraNavega.Enabled = True
                        FrmDetalle.Enabled = True
                        FrmABMDet.Enabled = True
                        FrmABMDet2.Visible = True
                        FrmDetalle2.Enabled = True
                        fraOpciones.Enabled = True
                        Exit Sub
                    End If
                    'GRABA RECIBO DETALLE
                    db.Execute "update fo_recibos_detalle_egresos set IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & "  where adjudica_codigo = " & rs_datos7!adjudica_codigo & " "
                    db.Execute "update fo_recibos_detalle_egresos set estado_destino = 'APR'  where adjudica_codigo = " & rs_datos7!adjudica_codigo & " "
                    
                    'ACTUALIZA APRUEBA ao_compra_adjudica
                    db.Execute "UPDATE ao_compra_adjudica SET estado_pagado = 'APR'  WHERE adjudica_codigo = " & rs_datos7!adjudica_codigo & " "
                    
                    ' ACTUALIZA TOTALES fo_traspaso_bancos_Egresos
                    db.Execute "update fo_traspaso_bancos_Egresos set total_bs = (select sum(fo_recibos_detalle_egresos.adjudica_bs) from fo_recibos_detalle_egresos where fo_recibos_detalle_egresos.IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & ")   " & _
                    " from fo_traspaso_bancos_Egresos inner join fo_recibos_detalle_egresos on  fo_traspaso_bancos_Egresos.IdTraspasoBancos = fo_recibos_detalle_egresos.IdTraspasoBancos where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & " "

                    db.Execute "update fo_traspaso_bancos_Egresos set total_dol = (select sum(fo_recibos_detalle_egresos.adjudica_dol) from fo_recibos_detalle_egresos where fo_recibos_detalle_egresos.IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & ")   " & _
                    " from fo_traspaso_bancos_Egresos inner join fo_recibos_detalle_egresos on  fo_traspaso_bancos_Egresos.IdTraspasoBancos = fo_recibos_detalle_egresos.IdTraspasoBancos where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & " "
                    
                    rs_datos7.MoveNext
                Wend
            End If
                FraNavega.Enabled = True
                FraNavega.Enabled = True
                FrmDetalle.Enabled = True
                FrmABMDet.Enabled = True
                FrmABMDet2.Visible = True
                FrmDetalle2.Enabled = True
                FrmDetalle2.Enabled = True
                fraOpciones.Enabled = True
            Call AbrirOrigen
            Call AbrirDetalle
        Else
            MsgBox "Debe elegir un registro cobrado,  vuelva a intentar ...", , "Atención"
        End If
    'Else
    '    MsgBox "El registro ya se encuentra procesado, vuelva a intentar ...", , "Atención"
    'End If
' Else
'    MsgBox "Debe elegir un registro para procesarlo,  vuelva a intentar ...", , "Atención"
' End If
Else
    MsgBox "El USUARIO no tiene Acceso !!. Consulte con el Administrador del Sistema. ", vbExclamation, "Atención!"
End If
  FraDet3.Visible = False
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub BtnCancelar3_Click()
        Fra_reporte.Visible = False
End Sub

Private Sub BtnCancelarBen_Click()
    frm_benef.Visible = False
    FraGrabarCancelar.Enabled = True
End Sub

Private Sub BtnEliminar_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    If Ado_datos.Recordset("estado_almacen") = "REG" Then
'      sino = MsgBox("Esta seguro de ANULAR la venta registrada ?", vbYesNo, "Confirmando")
'      If sino = vbYes Then
'          db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo = 'ANL' Where ao_ventas_cabecera.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  "
'          'Dim rstdestino As New ADODB.Recordset
'          'Set rstdestino = New ADODB.Recordset
'          'If rstdestino.State = 1 Then rstdestino.Close
'          'rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correl_venta = " & Ado_datos.Recordset("correl_venta") & " and venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  ", db, adOpenDynamic, adLockOptimistic
'          'If Not rstdestino.BOF Then rstdestino.MoveFirst
'          'If Not rstdestino.BOF And Not rstdestino.EOF Then
'          '    rstdestino("estado_codigo") = "E"
'          '    rstdestino.Update
'          'End If
'          'If rstdestino.State = 1 Then rstdestino.Close
'          marca1 = Ado_datos.Recordset.Bookmark
'          'Ado_datos.Recordset.Requery
'          'Ado_datos.Refresh
'          Call OptFilGral1_Click
'          Ado_datos.Recordset.Move marca1 - 1
'      End If
'    Else
'      MsgBox "NO se puede ANULAR el registro que ya fue Aprobado o previamente Anulado.", , "Atencion"
'    End If
'  Else
'    MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
     If rs_datos!estado_almacen = "REG" Then
       sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If sino = vbYes Then
'     If ExisteReg(Ado_datos.Recordset!unidad_codigo_sol, Ado_datos.Recordset!solicitud_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
          rs_datos!estado_almacen = "ANL"
'          rs_datos!fecha_registro = Date
'          rs_datos!usr_codigo = glusuario
'           Ado_datos.Recordset.Requery
'           Ado_datos.Refresh
           db.Execute "ap_ventas_grla 1 ,'" & glGestion & "', " & Ado_datos.Recordset!almacen_codigo & ", '" & Ado_datos.Recordset!doc_codigo_alm & "', " & Ado_datos.Recordset!doc_numero_alm & ", '" & ado_datos14.Recordset!bien_codigo & "', '" & Ado_datos.Recordset!EDIF_CODIGO & "'," & Ado_datos.Recordset!venta_codigo & ",'" & Ado_datos.Recordset!beneficiario_codigo_alm & "','" & Ado_datos.Recordset!fecha_verif & "'," & ado_datos14.Recordset!bien_cantidad_por_empaque & "," & precio_tot & ", " & IIf(IsNull(ado_datos14.Recordset!venta_precio_total_dol), 0, ado_datos14.Recordset!venta_precio_total_dol) & ", 'REG', '" & glusuario & "','" & Ado_datos.Recordset!venta_descripcion & "'," & precio_uni & ""
           Call AbrirDetalle
          rs_datos.UpdateBatch adAffectAll
       End If
    Else
       MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
    End If
  Else
      MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  Exit Sub
  
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnFiltro1_Click()
    SWFILTRO = 1
    VARFILTRO = 1
    Call AbrirDetalle
End Sub

Private Sub BtnFiltro2_Click()
    SWFILTRO = 1
    VARFILTRO = 2
    Call AbrirDetalle
End Sub

Private Sub BtnFiltro3_Click()
    SWFILTRO = 1
    VARFILTRO = 3
    Call AbrirDetalle
End Sub

Private Sub BtnFiltro4_Click()
    SWFILTRO = 1
    VARFILTRO = 4
    Call AbrirDetalle
End Sub

Private Sub BtnGrabar_Click()
On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If swgrabar = 2 Then
        'var_cod5 = Ado_datos.Recordset!venta_codigo
        'FInicio = IIf(IsNull(Ado_datos.Recordset!venta_fecha_inicio), Date, Ado_datos.Recordset!venta_fecha_inicio)
        'CANTOT = IIf(IsNull(Ado_datos.Recordset!venta_cantidad_total), 1, Ado_datos.Recordset!venta_cantidad_total)
        'gestion0 = IIf(IsNull(Ado_datos.Recordset!ges_gestion), glGestion, Ado_datos.Recordset!ges_gestion)
        VAR_BENEF = IIf(IsNull(Ado_datos.Recordset!beneficiario_codigo_resp), "0", Ado_datos.Recordset!beneficiario_codigo_resp)
        corrprog = Ado_datos.Recordset!correl_doc
        'VAR_MED = Ado_datos.Recordset!unimed_codigo
        'VAR_UNI = Ado_datos.Recordset!unidad_codigo
        'FControl = IIf(IsNull(Ado_datos.Recordset!fecha_verif), Date, Ado_datos.Recordset!fecha_verif)
        'Ado_datos.Recordset("fecha_verif") = DTPfechasol.Value
        '        rs_datos!fecha_verif = Date
        var_cod5 = Ado_datos.Recordset!IdTraspasoBancos
    End If
    'FrmCabecera.Visible = False
    Call grabar
    '
    'db.Execute "update ao_almacen_salidas set concepto = '" & TxtConcepto.Text & "' WHERE venta_codigo = " & var_cod5
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FraNavega.Enabled = True
    
    Fra_datos.Enabled = True
    dg_datos.Visible = True
    FrmDetalle.Visible = True
    'dtc_desc3.backColor = &H80000008
    'dtc_desc3.ForeColor = &H80000005
'    Fra_Total.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
    FrmCabecera.Visible = False
    'Refrescar Grid
    If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 And swgrabar = 2 Then
        rs_datos.Find "IdTraspasoBancos = " & var_cod5 & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
        rs_datos.MoveLast
     End If
     swgrabar = 0
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
  End If
    accion = ""
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub valida_campos()

  If dtc_codigo22 = "" Then
    MsgBox "Debe Elejir Cuenta Destino, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If dtc_codigo4 = "" Then
    MsgBox "Debe Elejir Responsable de la entrega ORIGEN, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If dtc_codigo11 = "" Then
'    MsgBox "Debe Elejir el Almacen!! , Vuelva a Intentar ...", vbExclamation, "Atención"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If dtc_codigo5 = "" Then
    MsgBox "Debe Elejir ... Entregado a:, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo3 = "" Then
    MsgBox "Debe Registrar el Documento ISO, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo21 = "" Then
    MsgBox "Debe Elejir Cuenta ORIGEN, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnGrabar2_Click()
On Error GoTo UpdateErr
If glusuario = "ASANTIVAÑEZ" Or glusuario = "TCASTILLO" Or glusuario = "FCABRERA" Or glusuario = "ADMIN" Or glusuario = "SPAREDES" Or glusuario = "VPAREDES" Or glusuario = "MWILDE" Or glusuario = "MVALDIVIA" Or glusuario = "EVILLALOBOS" Or glusuario = "CSALINAS" Then
 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        If ado_datos14.Recordset.RecordCount > 0 Then         '<> "" Then
            If (IsNull(ado_datos14.Recordset!adjudica_fecha) Or (ado_datos14.Recordset!adjudica_fecha = "01/01/1900")) Then         ' (ado_datos14.Recordset!trans_codigo <> "E") And
                MsgBox "No se puede ACEPTAR, verifique la fecha de Fectura o Recibo y vuelva a intentar ...", , "Atención"
                FraNavega.Enabled = True
                FraNavega.Enabled = True
                FrmDetalle.Enabled = True
                FrmABMDet.Enabled = True
                FrmABMDet2.Visible = True
                FrmDetalle2.Enabled = True
                FrmDetalle2.Enabled = True
                fraOpciones.Enabled = True
                Exit Sub
            End If
            'GRABA DESTINO DETALLE
            db.Execute "update fo_recibos_detalle_egresos set IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & "  where adjudica_codigo = " & ado_datos14.Recordset!adjudica_codigo & " "
            db.Execute "update fo_recibos_detalle_egresos set estado_destino = 'APR'  where adjudica_codigo = " & ado_datos14.Recordset!adjudica_codigo & " "
            
            'ACTUALIZA APRUEBA ao_compra_adjudica
            db.Execute "UPDATE ao_compra_adjudica SET estado_pagado = 'APR'  WHERE adjudica_codigo = " & ado_datos14.Recordset!adjudica_codigo & " "
            
            ' ACTUALIZA TOTALES fo_traspaso_bancos_Egresos
            db.Execute "update fo_traspaso_bancos_Egresos set total_bs = (select sum(fo_recibos_detalle_egresos.adjudica_bs) from fo_recibos_detalle_egresos where fo_recibos_detalle_egresos.IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & ")   " & _
            " from fo_traspaso_bancos_Egresos inner join fo_recibos_detalle_egresos on  fo_traspaso_bancos_Egresos.IdTraspasoBancos = fo_recibos_detalle_egresos.IdTraspasoBancos where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & " "
        
            db.Execute "update fo_traspaso_bancos_Egresos set total_dol = (select sum(fo_recibos_detalle_egresos.adjudica_dol) from fo_recibos_detalle_egresos where fo_recibos_detalle_egresos.IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & ")   " & _
            " from fo_traspaso_bancos_Egresos inner join fo_recibos_detalle_egresos on  fo_traspaso_bancos_Egresos.IdTraspasoBancos = fo_recibos_detalle_egresos.IdTraspasoBancos where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & " "
            
                FraNavega.Enabled = True
                FraNavega.Enabled = True
                FrmDetalle.Enabled = True
                FrmABMDet.Enabled = True
                FrmABMDet2.Visible = True
                FrmDetalle2.Enabled = True
                FrmDetalle2.Enabled = True
                fraOpciones.Enabled = True
            Call AbrirOrigen
            Call AbrirDetalle
        Else
            MsgBox "Debe elegir un registro cobrado,  vuelva a intentar ...", , "Atención"
        End If
    Else
        MsgBox "El registro ya se encuentra procesado, vuelva a intentar ...", , "Atención"
    End If
 Else
    MsgBox "Debe elegir un registro para procesarlo,  vuelva a intentar ...", , "Atención"
 End If
Else
        MsgBox "El USUARIO no tiene Acceso !!. Consulte con el Administrador del Sistema. ", vbExclamation, "Atención!"
End If
  FraDet3.Visible = False
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub BtnGrabarBen_Click()
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from gc_beneficiario_vs_cta_banco where cta_codigo = '" & dtc_cta8.Text & "' and beneficiario_codigo = '" & dtc_codigo8.Text & "'  ", db, adOpenStatic
    If rs_datos10.RecordCount = 0 Then
        'abrir gc_beneficiario_vs_cta_banco
        db.Execute "INSERT INTO gc_beneficiario_vs_cta_banco (beneficiario_codigo, cta_codigo, cta_tipo, bco_codigo, cta_descripcion, codigo_empresa, estado_codigo, usr_codigo, fecha_registro) " & _
        " VALUES ('" & dtc_codigo8.Text & "', '" & dtc_cta8.Text & "', 'C', 'NN', '" & dtc_nom8.Text & "', '1', 'APR', '" & glusuario & "', '" & Date & "')"
        'Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
        Set rs_datos22 = New ADODB.Recordset
        If rs_datos22.State = 1 Then rs_datos22.Close
        rs_datos22.Open "select * from fv_beneficiario_vs_cta_banco ", db, adOpenStatic
        Set Ado_datos22.Recordset = rs_datos22
        dtc_desc22.BoundText = dtc_moneda22.BoundText
        dtc_codigo22.BoundText = dtc_moneda22.BoundText
        'FraGrabarCancelar.Enabled = True
    Else
        MsgBox "Ya existe el Beneficiario relacionado, Verifique y Vuelva a intentar ...", , "Atención"
    End If
    FraGrabarCancelar.Enabled = True
    frm_benef.Visible = False

End Sub

Private Sub BtnImprimir_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        VAR_IDTRP = Ado_datos.Recordset!IdTraspasoBancos
        
'        db.Execute "UPDATE fo_recibos_detalle_egresos set fo_recibos_detalle_egresos.trans_codigo  = ao_ventas_cobranza_det.trans_codigo FROM fo_recibos_detalle_egresos INNER JOIN ao_ventas_cobranza_det ON fo_recibos_detalle_egresos.adjudica_codigo  = ao_ventas_cobranza_det.adjudica_codigo where fo_recibos_detalle_egresos.trans_codigo Is Null"
'
'        db.Execute "UPDATE fo_traspaso_bancos_Egresos set fo_traspaso_bancos_Egresos.total_bs  = fv_tes_recibos_detalle_sum_egreso.adjudica_bs, fo_traspaso_bancos_Egresos.total_dol   = fv_tes_recibos_detalle_sum_egreso.adjudica_dol from fo_traspaso_bancos_Egresos inner join fv_recibos_detalle_sum " & _
'        " on fo_traspaso_bancos_Egresos.IdTraspasoBancos  = fv_tes_recibos_detalle_sum_egreso.IdTraspasoBancos where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & VAR_IDTRP & "  "
'
'        db.Execute "UPDATE fo_recibos_detalle_egresos set fo_recibos_detalle_egresos.cta_codigo_origen = fo_traspaso_bancos_Egresos.cta_codigo, fo_recibos_detalle_egresos.cta_codigo_destino  = fo_traspaso_bancos_Egresos.cta_codigo_destino FROM fo_recibos_detalle_egresos INNER JOIN fo_traspaso_bancos_Egresos " & _
'        " ON fo_recibos_detalle_egresos.IdTraspasoBancos = fo_traspaso_bancos_Egresos.IdTraspasoBancos where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & VAR_IDTRP & "  "
'
'        db.Execute "UPDATE fo_traspaso_bancos_Egresos set fo_traspaso_bancos_Egresos.literal = (Select dbo.CantidadConLetra(dbo.fo_traspaso_bancos_Egresos.total_bs) From fo_traspaso_bancos_Egresos Where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & VAR_IDTRP & ") where IdTraspasoBancos = " & VAR_IDTRP & "  "
'
'        db.Execute "UPDATE fo_traspaso_bancos_Egresos set fo_traspaso_bancos_Egresos.literalDol=  (Select dbo.CantidadConLetra(dbo.fo_traspaso_bancos_Egresos.total_dol) From fo_traspaso_bancos_Egresos Where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & VAR_IDTRP & ") where IdTraspasoBancos = " & VAR_IDTRP & "  "
        
        Set rs_datos1 = New ADODB.Recordset
        If rs_datos1.State = 1 Then rs_datos1.Close
        rs_datos1.Open "Select * from fo_traspaso_bancos_Egresos WHERE IdTraspasoBancos = " & VAR_IDTRP & " ", db, adOpenStatic
        If rs_datos1.RecordCount > 0 Then
            VAR_LITERAL1 = rs_datos1!Literal + "BOLIVIANOS"
            VAR_LITERAL2 = IIf(IsNull(rs_datos1!LiteralDol), "0", rs_datos1!LiteralDol) + "DOLARES AMERICANOS"
        Else
            VAR_LITERAL1 = ""
            VAR_LITERAL2 = ""
        End If
        
        CryV01.Reset
        CryV01.WindowState = crptMaximized
        CryV01.WindowShowSearchBtn = True
        CryV01.WindowShowRefreshBtn = True
        CryV01.WindowShowPrintSetupBtn = True
        
        Dim iResult As Integer
        If GlBaseDatos = "ADMIN_EMPRESA" Then
            CryV01.ReportFileName = App.Path & "\Reportes\Tesoreria\fr_traspasos_tesoreria_egresos.rpt"
        Else
            CryV01.ReportFileName = App.Path & "\Reportes\Tesoreria\fr_traspasos_tesoreria_egresosPrueba.rpt"
        End If
            var_titulo = "TRASPASO BANCOS EGRESOS"
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True
        CryV01.StoredProcParam(0) = Ado_datos.Recordset!IdTraspasoBancos
        'CryV01.StoredProcParam(1) = Ado_datos.Recordset!ges_gestion
        CryV01.Formulas(0) = "titulo = '" & var_titulo & "' "
        CryV01.Formulas(1) = "subtitulo = 'DETALLE DEL ARQUEO' "
        CryV01.Formulas(2) = "Literal1 = '" & VAR_LITERAL1 & "' "
        CryV01.Formulas(3) = "Literal2 = '" & VAR_LITERAL2 & "' "
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
        CryV01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
    End If
 
End Sub


Private Sub BtnImprimir1_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        CryV01.Reset
        CryV01.WindowState = crptMaximized
        CryV01.WindowShowSearchBtn = True
        CryV01.WindowShowRefreshBtn = True
        CryV01.WindowShowPrintSetupBtn = True
        
        Dim iResult As Integer
            CryV01.ReportFileName = App.Path & "\Reportes\Tesoreria\fr_recibos_oficiales_tesoreria.rpt"
            var_titulo = "RECIBO DE TESORERIA"
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True
        CryV01.StoredProcParam(0) = Ado_datos.Recordset!IdRecibo
        'CryV01.StoredProcParam(1) = Ado_datos.Recordset!ges_gestion
        CryV01.Formulas(0) = "titulo = '" & var_titulo & "' "
        CryV01.Formulas(1) = "subtitulo = 'DETALLE DE COBRNZAS' "
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
        CryV01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
    End If
End Sub

Private Sub BtnModificar_Click()
On Error GoTo UpdateErr
If glusuario = "TCASTILLO" Or glusuario = "FCABRERA" Or glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "MWILDE" Or glusuario = "MVALDIVIA" Or glusuario = "EVILLALOBOS" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset("estado_codigo") = "REG" Then
        accion = "MOD"
        FrmCabecera.Visible = True
        FrmDetalle.Visible = False
        FraNavega.Enabled = False
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
'        If dtc_desc4.Text = "" Or dtc_desc11.Text = "" Or dtc_desc21.Text = "" Then
'            Fra_datos.Enabled = True
'        Else
'            Fra_datos.Enabled = False
'        End If
'        Fra_Total.Visible = False
        FrmABMDet.Visible = False
        FrmABMDet2.Visible = False
        swgrabar = 2
'        SSTab1.Tab = 0
'        SSTab1.TabEnabled(0) = True
'        SSTab1.TabEnabled(1) = False
        'If Ado_datos.Recordset!unidad_codigo = "UALMI" Or Ado_datos.Recordset!unidad_codigo = "UALMR" Or Ado_datos.Recordset!unidad_codigo = "UALMH" Or Ado_datos.Recordset!unidad_codigo = "DADM" Then
        'If Ado_datos.Recordset!unidad_codigo = VAR_ORIGEN Then
'        If VAR_ORIGEN = "UALMR" Then
'            dtc_desc3.Locked = False
'            dtc_desc3.Width = 5955
'            'TxtConcepto.Locked = False
'        Else
'            dtc_desc3.Width = 6315
'            dtc_desc3.Locked = True
'            'TxtConcepto.Locked = True
'        End If
    Else
      MsgBox "NO se puede MODIFICAR, porque el registro ya fue Aprobado, Anulado o Cerrado.", , "Atencion"
    End If
  Else
        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
Else
        MsgBox "El USUARIO no tiene Acceso !!. Consulte con el Administrador del Sistema. ", vbExclamation, "Atención!"
End If
    
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub BtnRefrescar_Click()
    SWFILTRO = 1
    VARFILTRO = 0
    Call AbrirDetalle
End Sub

Private Sub BtnNuevaCta_Click()
    Set rs_datos8 = New ADODB.Recordset     'Beneficiario Personas Nat. y Juridicas
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_beneficiario where tipoben_codigo <> '0' and tipoben_codigo <> '1' and estado_codigo = 'APR' ORDER BY beneficiario_denominacion", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    If Ado_datos8.Recordset.RecordCount > 0 Then
        dtc_desc8.BoundText = dtc_codigo8.BoundText
        FraGrabarCancelar.Enabled = False
        frm_benef.Visible = True
    End If

End Sub

Private Sub BtnSalir_Click()
    sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
    If sino = vbYes Then
'        Ado_datos.Recordset.Close
'        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        If rs_Ventas.State = 1 Then rs_Ventas.Close
        Unload Me
    End If
End Sub

'Private Sub Cmd_Cliente_Click()
'    glPersNew = "P"
'    frmBeneficiario.Show 'vbModal
'End Sub

Private Sub CmdCancelaCobro_Click()
  FrmCobros.Enabled = False
  'swgrabar = 0
  'Call cerea
  swnuevo = 0
  If Ado_datos.Recordset("estado_codigo") = "REG" Then
    Call OptFilGral1_Click
  Else
    Call OptFilGral2_Click
  End If
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
''    SSTab1.TabEnabled(2) = False
    FraNavega.Enabled = True
    fraOpciones.Enabled = True
    FrmDetalle.Visible = True
'    FrmCobranza.Visible = True
    TxtCobrador.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
End Sub

Private Sub BtnDesAprobar_Click()
On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
     If rs_datos!estado_codigo = "APR" Or rs_datos!estado_verificado = "APR" Then
       sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If sino = vbYes Then
           rs_datos!estado_codigo = "REG"
           rs_datos!estado_verificado = "REG"
           Call AbrirDetalle
          rs_datos.UpdateBatch adAffectAll
       End If
    Else
       MsgBox "No se puede DESPROBAR un registro Aulado(ANL) o Registrado (REG) ...", vbExclamation, "Validación de Registro"
    End If
  Else
      MsgBox "NO se puede DESAPROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  Exit Sub
  
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub CmdNOunidad_Click()
    swunidad = 0
    Frmunidad.Visible = False
End Sub

Private Sub CmdOKunidad_Click()
    swunidad = 1
        If swunidad = 1 Then
            Dim rstpagos As New ADODB.Recordset
            Set rstpagos = New ADODB.Recordset
            If rstpagos.State = 1 Then rstpagos.Close
            rstpagos.Open "select * from pagos where GES_gestion = '5000'", db, adOpenKeyset, adLockOptimistic
            rstpagos.AddNew
                rstpagos("ges_gestion") = glGestion     'Ado_datos.Recordset("ges_gestion")
                rstpagos("org_codigo") = DataCombo1.Text   'Ado_datos.Recordset("formulario")
                rstpagos("codigo_pago") = "" 'genera jorge
                rstpagos("codigo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
                rstpagos("formulario") = Ado_datos.Recordset("formulario")
                rstpagos("codigo_unidad") = Ado_datos.Recordset("codigo_unidad")
                rstpagos("monto_bolivianos") = Ado_datos.Recordset("monto_bolivianos")
                rstpagos("estado_compromiso") = "N"
                rstpagos("justificacion") = Ado_datos.Recordset("justificacion_solicitud")
            rstpagos.Update
        End If
End Sub

Private Sub BtnImprimir2_Click()
    If ado_datos14.Recordset.RecordCount > 0 Then
         Dim iResult As Integer
        'Dim co As New ADODB.Command
        'CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_almacen_kardex.rpt"
        CryR01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_kardex_almacen_acumulado.rpt" '
        CryR01.WindowShowPrintSetupBtn = True
        CryR01.WindowShowRefreshBtn = True
        'CryR01.StoredProcParam(0) = Ado_datos.Recordset!bien_codigo
        CryR01.StoredProcParam(0) = ado_datos14.Recordset!bien_codigo
        CryR01.StoredProcParam(1) = Trim(Str(ado_datos14.Recordset!almacen_codigo))            'dtc_codigo1.Text
        CryR01.StoredProcParam(2) = Format(DTP_Finicio.Value, "dd/mm/yyyy")
        CryR01.StoredProcParam(3) = Format(DTP_Ffin.Value, "dd/mm/yyyy")
        CryR01.Formulas(0) = "almace = '" & dtc_desc1.Text & "' "
        'CryR01.Formulas(2) = "DEL_AL = '' "
        'CryR01.Formulas(3) = "fechafin = '" & DTP_Ffin.Value & "' "
        
        iResult = CryR01.PrintReport
        If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
        CryR01.WindowState = crptMaximized
        Fra_reporte.Visible = False
    Else
        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
    End If
    Fra_reporte.Visible = True
End Sub

Private Sub BtnAnlDetalle_Click()
On Error GoTo UpdateErr
If glusuario = "TCASTILLO" Or glusuario = "FCABRERA" Or glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "MWILDE" Or glusuario = "MVALDIVIA" Or glusuario = "EVILLALOBOS" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then           '
 If Ado_datos11.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo = "REG" Then
          sino = MsgBox("Está Seguro de ANULAR el Registro Activo --> ", vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
            'BORRA RECIBO DETALLE
            db.Execute "update fo_recibos_detalle_egresos set IdTraspasoBancos = '0'  where adjudica_codigo = " & Ado_datos11.Recordset!adjudica_codigo & " "
            db.Execute "update fo_recibos_detalle_egresos set estado_destino = 'REG'  where adjudica_codigo = " & Ado_datos11.Recordset!adjudica_codigo & " "
            
            'ACTUALIZA APRUEBA ao_compra_adjudica
            db.Execute "UPDATE ao_compra_adjudica SET estado_pagado = 'REG'  WHERE adjudica_codigo = " & Ado_datos11.Recordset!adjudica_codigo & " "
            
            ' ACTUALIZA TOTALES fo_traspaso_bancos_Egresos
            db.Execute "update fo_traspaso_bancos_Egresos set total_bs = (select sum(fo_recibos_detalle_egresos.adjudica_bs) from fo_recibos_detalle_egresos where fo_recibos_detalle_egresos.IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & ")   " & _
            " from fo_traspaso_bancos_Egresos inner join fo_recibos_detalle_egresos on  fo_traspaso_bancos_Egresos.IdTraspasoBancos = fo_recibos_detalle_egresos.IdTraspasoBancos where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & " "
        
            db.Execute "update fo_traspaso_bancos_Egresos set total_dol = (select sum(fo_recibos_detalle_egresos.adjudica_dol) from fo_recibos_detalle_egresos where fo_recibos_detalle_egresos.IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & ")   " & _
            " from fo_traspaso_bancos_Egresos inner join fo_recibos_detalle_egresos on  fo_traspaso_bancos_Egresos.IdTraspasoBancos = fo_recibos_detalle_egresos.IdTraspasoBancos where fo_traspaso_bancos_Egresos.IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & " "
            Call AbrirOrigen
            Call AbrirDetalle
          End If
       Else
          MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
 Else
     MsgBox "No se puede ANULAR, el registro ya fue APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
 End If
Else
        MsgBox "El USUARIO no tiene Acceso !!. Consulte con el Administrador del Sistema. ", vbExclamation, "Atención!"
End If
 'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub BtnModDetalle_Click()
If glusuario = "TCASTILLO" Or glusuario = "RGIL" Or glusuario = "FCABRERA" Or glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "MWILDE" Or glusuario = "MVALDIVIA" Or glusuario = "EVILLALOBOS" Or glusuario = "ASANTIVAÑEZ" Then
 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo = "REG" And Ado_datos.Recordset!estado_verificado = "REG" Then
        If Ado_datos11.Recordset.RecordCount > 0 Then         '<> "" Then
            'MODIFICA TRASPASO DETALLE
            Text11.Text = IIf(IsNull(Ado_datos11.Recordset!cmpbte_deposito_bco), 0, Ado_datos11.Recordset!cmpbte_deposito_bco)
            DTP_Finicio.Value = IIf(IsNull(Ado_datos11.Recordset!fecha_registro_bco), Date, Ado_datos11.Recordset!fecha_registro_bco)
            'Label6.Caption = Ado_datos11.Recordset!trans_descripcion
            Call Extracto
            'DtGLista.Enabled = True
        Else
            MsgBox "Debe elegir un registro cobrado para modificar, verifique y vuelva a intentar ...", , "Atención"
        End If
    Else
        If Ado_datos.Recordset!estado_codigo = "REG" And Ado_datos.Recordset!estado_verificado = "APR" And (glusuario = "MVALDIVIA" Or glusuario = "ADMIN") Then
            If Ado_datos11.Recordset.RecordCount > 0 Then
                Text11.Text = IIf(IsNull(Ado_datos11.Recordset!cmpbte_deposito_bco), 0, Ado_datos11.Recordset!cmpbte_deposito_bco)
                DTP_Finicio.Value = IIf(IsNull(Ado_datos11.Recordset!fecha_registro_bco), Date, Ado_datos11.Recordset!fecha_registro_bco)
                
                Label6.Caption = Ado_datos11.Recordset!trans_descripcion
                Call Extracto
            Else
                MsgBox "Debe elegir un registro cobrado para modificar, verifique y vuelva a intentar ...", , "Atención"
            End If
        Else
            MsgBox "El registro ya se encuentra APROBADO, Verifique y vuelva a intentar ...", , "Atención"
        End If
    End If
 Else
    MsgBox "Debe elegir un registro para procesarlo,  vuelva a intentar ...", , "Atención"
 End If
Else
        MsgBox "El USUARIO no tiene Acceso !!. Consulte con el Administrador del Sistema. ", vbExclamation, "Atención!"
End If
End Sub

Private Sub Extracto()
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = False
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
    FraNavega.Enabled = False
    'FraExtracto.Visible = True
    Fra_reporte.Visible = True
    '-- ACTUALIZA ESTADO TRASPASOS EN fo_extracto_egreso_GRAL
    db.Execute "UPDATE fo_extracto_egreso_GRAL SET estado_conciliado = 'REG' "
    '-- EN BOLIVIANOS
        db.Execute "UPDATE fo_extracto_egreso_GRAL SET fo_extracto_egreso_GRAL.estado_conciliado = 'APR' FROM fo_extracto_egreso_GRAL INNER JOIN fv_recibos_detEgreso_sum_cmpbte_APR ON fv_recibos_detEgreso_sum_cmpbte_APR.cmpbte_deposito_bco = fo_extracto_egreso_GRAL.cod_bancarizacion AND fv_recibos_detEgreso_sum_cmpbte_APR.fecha_registro_bco = fo_extracto_egreso_GRAL.fecha_transaccion AND fv_recibos_detEgreso_sum_cmpbte_APR.cta_codigo_destino = fo_extracto_egreso_GRAL.cuenta " & _
        " AND fv_recibos_detEgreso_sum_cmpbte_APR.adjudicaBs  = fo_extracto_egreso_GRAL.monto_bs WHERE (fo_extracto_egreso_GRAL.cuenta ='2015046557-03-054' OR fo_extracto_egreso_GRAL.cuenta ='4010439742' OR fo_extracto_egreso_GRAL.cuenta ='4010620792' OR fo_extracto_egreso_GRAL.cuenta ='4010644195' OR fo_extracto_egreso_GRAL.cuenta ='4010772049' OR fo_extracto_egreso_GRAL.cuenta ='4011005599' " & _
        " OR fo_extracto_egreso_GRAL.cuenta ='4011048967' OR fo_extracto_egreso_GRAL.cuenta ='4011048981' OR fo_extracto_egreso_GRAL.cuenta ='4069626219' OR fo_extracto_egreso_GRAL.cuenta ='4069626233' OR fo_extracto_egreso_GRAL.cuenta ='10000019133060')  "
    '-- EN DOLARES
        db.Execute "UPDATE fo_extracto_egreso_GRAL SET estado_conciliado = 'APR' FROM fo_extracto_egreso_GRAL INNER JOIN fv_recibos_detEgreso_sum_cmpbte_APR ON fv_recibos_detEgreso_sum_cmpbte_APR.cmpbte_deposito_bco = fo_extracto_egreso_GRAL.cod_bancarizacion AND fv_recibos_detEgreso_sum_cmpbte_APR.fecha_registro_bco = fo_extracto_egreso_GRAL.fecha_transaccion AND fv_recibos_detEgreso_sum_cmpbte_APR.cta_codigo_destino = fo_extracto_egreso_GRAL.cuenta AND fv_recibos_detEgreso_sum_cmpbte_APR.adjudicaDol = fo_extracto_egreso_GRAL.monto_dol " & _
        " WHERE (fo_extracto_egreso_GRAL.cuenta ='201-5041743-2-18' OR fo_extracto_egreso_GRAL.cuenta ='096359-201-9' OR fo_extracto_egreso_GRAL.cuenta ='4010038393' OR fo_extracto_egreso_GRAL.cuenta ='4010620785' OR fo_extracto_egreso_GRAL.cuenta ='4010780124' OR fo_extracto_egreso_GRAL.cuenta ='4011005601' OR fo_extracto_egreso_GRAL.cuenta ='4011048974' OR fo_extracto_egreso_GRAL.cuenta ='4069626242' OR fo_extracto_egreso_GRAL.cuenta ='4069626265' ) "
    '---APROBAR: VARIOS EN SOFIA VS. UNO EN EXTRACTO
        'db.Execute "UPDATE fo_recibos_oficiales_egresos SET fo_recibos_oficiales_egresos.estado_conciliado = 'APR' FROM fo_recibos_oficiales_egresos INNER JOIN fo_extracto_egreso_GRAL ON fo_recibos_oficiales_egresos.cmpbte_deposito_bco = fo_extracto_egreso_GRAL.cod_bancarizacion AND fo_recibos_oficiales_egresos.fecha_registro_bco = fo_extracto_egreso_GRAL.fecha_transaccion AND fo_recibos_oficiales_egresos.cta_codigo_destino = fo_extracto_egreso_GRAL.cuenta AND fo_recibos_oficiales_egresos.total_bs = fo_extracto_egreso_GRAL.monto_bs WHERE (fo_extracto_egreso_GRAL.estado_conciliado = 'APR') AND (fo_recibos_oficiales_egresos.estado_conciliado  ='REG') "
    '---ACTUALIZA estado_conciliado (Anterior)
    'db.Execute "update fo_extracto_egreso_GRAL SET estado_conciliado = 'REG' "
    'db.Execute "update fo_extracto_egreso_GRAL SET estado_conciliado = 'APR' FROM fo_extracto_egreso_GRAL INNER JOIN fo_recibos_oficiales_egresos ON fo_extracto_egreso_GRAL.cod_bancarizacion = fo_recibos_oficiales_egresos.cmpbte_deposito_bco AND fo_extracto_egreso_GRAL.cuenta  = fo_recibos_oficiales_egresos.cta_codigo_destino AND fo_extracto_egreso_GRAL.fecha_transaccion = fo_recibos_oficiales_egresos.fecha_registro_bco AND fo_extracto_egreso_GRAL.monto_bs = fo_recibos_oficiales_egresos.cobranza_bs "

    Set rs_datos18 = New ADODB.Recordset
    If rs_datos18.State = 1 Then rs_datos18.Close
    rs_datos18.Open "Select * from fv_extracto_egresos_NO_conciliados order by cod_bancarizacion", db, adOpenStatic        'fecha_transaccion, hora_transaccion
    Set ado_datos18.Recordset = rs_datos18
    If ado_datos18.Recordset.RecordCount > 0 Then
        DctFecha18.BoundText = DctCod18.BoundText
        DctMonto18.BoundText = DctCod18.BoundText
        DctCliente18.BoundText = DctCod18.BoundText
        DctDeposita18.BoundText = DctCod18.BoundText
        DctOrigina18.BoundText = DctCod18.BoundText
    Else
        MsgBox "No Existen registros de Extractos Pendientes, Debe Migrar los Extactos de esta Cuenta y vuelva a intentar ...", , "Atención"
    End If
End Sub

Private Sub DctCliente18_Click(Area As Integer)
    DctCod18.BoundText = DctCliente18.BoundText
    DctFecha18.BoundText = DctCliente18.BoundText
    DctMonto18.BoundText = DctCliente18.BoundText
    DctDeposita18.BoundText = DctCliente18.BoundText
    DctOrigina18.BoundText = DctCliente18.BoundText
    DctMontoDol18.BoundText = DctCliente18.BoundText
    DctCuenta18.BoundText = DctCliente18.BoundText
End Sub

Private Sub DctCod18_Click(Area As Integer)
    DctFecha18.BoundText = DctCod18.BoundText
    DctMonto18.BoundText = DctCod18.BoundText
    DctCliente18.BoundText = DctCod18.BoundText
    DctDeposita18.BoundText = DctCod18.BoundText
    DctOrigina18.BoundText = DctCod18.BoundText
    DctMontoDol18.BoundText = DctCod18.BoundText
    DctCuenta18.BoundText = DctCod18.BoundText
End Sub

Private Sub DctCuenta18_Click(Area As Integer)
    DctCod18.BoundText = DctCuenta18.BoundText
    DctFecha18.BoundText = DctCuenta18.BoundText
    DctCliente18.BoundText = DctCuenta18.BoundText
    DctDeposita18.BoundText = DctCuenta18.BoundText
    DctOrigina18.BoundText = DctCuenta18.BoundText
    DctMonto18.BoundText = DctCuenta18.BoundText
    DctMontoDol18.BoundText = DctCuenta18.BoundText
End Sub

Private Sub DctDeposita18_Click(Area As Integer)
    DctCod18.BoundText = DctDeposita18.BoundText
    DctFecha18.BoundText = DctDeposita18.BoundText
    DctMonto18.BoundText = DctDeposita18.BoundText
    DctCliente18.BoundText = DctDeposita18.BoundText
    DctOrigina18.BoundText = DctDeposita18.BoundText
    DctMontoDol18.BoundText = DctDeposita18.BoundText
    DctCuenta18.BoundText = DctDeposita18.BoundText
End Sub

Private Sub DctFecha18_Click(Area As Integer)
    DctCod18.BoundText = DctFecha18.BoundText
    DctMonto18.BoundText = DctFecha18.BoundText
    DctCliente18.BoundText = DctFecha18.BoundText
    DctDeposita18.BoundText = DctFecha18.BoundText
    DctOrigina18.BoundText = DctFecha18.BoundText
    DctMontoDol18.BoundText = DctFecha18.BoundText
    DctCuenta18.BoundText = DctFecha18.BoundText
End Sub

Private Sub DctMonto18_Click(Area As Integer)
    DctCod18.BoundText = DctMonto18.BoundText
    DctFecha18.BoundText = DctMonto18.BoundText
    DctCliente18.BoundText = DctMonto18.BoundText
    DctDeposita18.BoundText = DctMonto18.BoundText
    DctOrigina18.BoundText = DctMonto18.BoundText
    DctMontoDol18.BoundText = DctMonto18.BoundText
    DctCuenta18.BoundText = DctMonto18.BoundText
End Sub

Private Sub DctMontoDol18_Click(Area As Integer)
    DctCod18.BoundText = DctMontoDol18.BoundText
    DctFecha18.BoundText = DctMontoDol18.BoundText
    DctCliente18.BoundText = DctMontoDol18.BoundText
    DctDeposita18.BoundText = DctMontoDol18.BoundText
    DctOrigina18.BoundText = DctMontoDol18.BoundText
    DctMonto18.BoundText = DctMontoDol18.BoundText
    DctCuenta18.BoundText = DctMontoDol18.BoundText
End Sub

Private Sub DctOrigina18_Click(Area As Integer)
    DctCod18.BoundText = DctOrigina18.BoundText
    DctFecha18.BoundText = DctOrigina18.BoundText
    DctMonto18.BoundText = DctOrigina18.BoundText
    DctCliente18.BoundText = DctOrigina18.BoundText
    DctDeposita18.BoundText = DctOrigina18.BoundText
    DctMontoDol18.BoundText = DctOrigina18.BoundText
    DctCuenta18.BoundText = DctOrigina18.BoundText
End Sub


Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_aux8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_aux8.BoundText
    dtc_codigo8.BoundText = dtc_aux8.BoundText
End Sub

Private Sub dtc_codigo21_Click(Area As Integer)
    dtc_desc21.BoundText = dtc_codigo21.BoundText
    VAR_ALMX = dtc_codigo21.BoundText
End Sub

Private Sub dtc_codigo21_LostFocus()
    dtc_codigo21.BoundText = VAR_ALMX
    dtc_desc21.BoundText = dtc_codigo21.BoundText
End Sub

Private Sub dtc_codigo22_Click(Area As Integer)
    dtc_desc22.BoundText = dtc_codigo22.BoundText
    dtc_moneda22.BoundText = dtc_codigo22.BoundText
    VAR_ALMT = dtc_codigo22.BoundText
End Sub

Private Sub dtc_codigo22_LostFocus()
    dtc_codigo22.BoundText = VAR_ALMT
    dtc_desc22.BoundText = dtc_codigo22.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    'dtc_recibo6.BoundText = dtc_codigo6.BoundText
    'dtc_fecha6.BoundText = dtc_codigo6.BoundText
    dtc_reciboCobr6.BoundText = dtc_codigo6.BoundText
    'dtc_edificio6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    dtc_aux8.BoundText = dtc_codigo8.BoundText
End Sub

'Private Sub dtc_desc15_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        TxtDescuento.SetFocus
'    End If
'End Sub

Private Sub dtc_desc21_Click(Area As Integer)
  dtc_codigo21.BoundText = dtc_desc21.BoundText
End Sub

Private Sub dtc_desc22_Click(Area As Integer)
    dtc_codigo22.BoundText = dtc_desc22.BoundText
    dtc_moneda22.BoundText = dtc_desc22.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
    VAR_BEN2 = dtc_codigo4.Text
    Call pCta1(dtc_codigo4.Text)
    dtc_desc21.Enabled = True
End Sub

Private Sub pCta1(CodigoA As String)
   Dim strConsultaF As String

   strConsultaF = "select * from fc_cuenta_bancaria where (cta_es_CUT = 'E') or (beneficiario_codigo = '" & CodigoA & "') or (hora_registro = '" & CodigoA & "') "

   Set dtc_codigo21.RowSource = Nothing
   Set dtc_codigo21.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo21.ReFill
   dtc_codigo21.BoundText = Empty

   Set dtc_desc21.RowSource = Nothing
   Set dtc_desc21.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc21.ReFill
   dtc_desc21.BoundText = Empty

End Sub

Private Sub dtc_desc4_LostFocus()
    dtc_codigo4.Text = VAR_BEN2
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    VAR_BEN3 = dtc_codigo5.Text
End Sub

'Private Sub pCta1(CodigoA As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from ac_almacenes where beneficiario_codigo = '" & CodigoA & "'"
'
'   Set dtc_codigo20.RowSource = Nothing
'   Set dtc_codigo20.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo20.ReFill
'   dtc_codigo20.BoundText = Empty
'
'   Set dtc_desc20.RowSource = Nothing
'   Set dtc_desc20.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc20.ReFill
'   dtc_desc20.BoundText = Empty
'
'End Sub

'Private Sub dtc_codigo13_Click(Area As Integer)
'    dtc_desc13.BoundText = dtc_codigo13.BoundText
'    Dtc_Stock13.BoundText = dtc_codigo13.BoundText
'End Sub

Private Sub dtc_codigo2A_Click(Area As Integer)
    dtc_desc2A.BoundText = dtc_codigo2A.BoundText
End Sub

Private Sub dtc_codigo4A_Click(Area As Integer)
    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
End Sub

Private Sub DataCombo1_Click(Area As Integer)
'    DataCombo2.Text = DataCombo1.BoundText
End Sub

Private Sub DataCombo2_Click(Area As Integer)
    DataCombo1.Text = DataCombo2.BoundText
End Sub

Private Sub cmdVerifica_existencia_Click()
' verifica existencia  del almacen
Cant_Alm = 0
AlFrmExistencia_Almacen.Show

DE.dbo_albSacaDetalleMaterial Mid(TxtCodigo, 3, 12), descri_bien, Cant_Alm
Txtcant_alm = Cant_Alm
If Cant_Alm >= TxtCantPedi Then
        optSi = True
    Else
        optNo = True
    End If
End Sub

'Private Sub dtc_codigo11_Click(Area As Integer)
'    dtc_desc11.BoundText = dtc_codigo11.BoundText
'    dtc_Aux11.BoundText = dtc_codigo11.BoundText
'End Sub

'Private Sub dtc_desc11_Click(Area As Integer)
'    dtc_codigo11.BoundText = dtc_desc11.BoundText
'    dtc_Aux11.BoundText = dtc_desc11.BoundText
'    Call pDepto(dtc_Aux11.Text)
'    dtc_desc21.Enabled = True
'End Sub

'Private Sub pDepto(CodigoA As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_departamento where depto_codigo  = '" & CodigoA & "'"
'
'   Set dtc_codigo21.RowSource = Nothing
'   Set dtc_codigo21.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo21.ReFill
'   dtc_codigo21.BoundText = Empty
'
'   Set dtc_desc21.RowSource = Nothing
'   Set dtc_desc21.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc21.ReFill
'   'dtc_desc21.BoundText = Empty
'End Sub

Private Sub dtccodmanejo_Click(Area As Integer)
    DtCCodigo.BoundText = dtccodmanejo.BoundText
    DtCDescripcion.BoundText = dtccodmanejo.BoundText
    dtcunidadmedida.BoundText = dtccodmanejo.BoundText
    dtccodpeso.BoundText = dtccodmanejo.BoundText
End Sub

Private Sub dtccodpeso_Click(Area As Integer)
    DtCCodigo.BoundText = dtccodpeso.BoundText
    DtCDescripcion.BoundText = dtccodpeso.BoundText
    dtcunidadmedida.BoundText = dtccodpeso.BoundText
    dtccodmanejo.BoundText = dtccodpeso.BoundText
End Sub


Private Sub dtccodpar_Click(Area As Integer)
    dtcdescripar.Text = dtccodpar.BoundText
End Sub

Private Sub dtccodpoa_Click(Area As Integer)
    dtcdespoa.Text = dtccodpoa.BoundText
End Sub

Private Sub dtccodpuesto_Click(Area As Integer)
    dtcdenopuesto.Text = dtccodpuesto.BoundText
End Sub

Private Sub dtccodtipoid_Click(Area As Integer)
    dtcdescrtipoid.BoundText = dtccodtipoid.BoundText
End Sub

Private Sub dtccoduni_Click(Area As Integer)
    dtcdescripuni.Text = dtccoduni.BoundText
End Sub

Private Sub dtccorrcompromiso_Click(Area As Integer)
    dtcfechacompromiso.BoundText = dtccorrcompromiso.BoundText
End Sub

Private Sub dtccorrsol_Click(Area As Integer)
 dtcfechasol.BoundText = dtccorrsol.BoundText
End Sub

Private Sub dtcdenominacionruc_Click(Area As Integer)
    dtcnroruc.BoundText = dtcdenominacionruc.BoundText
End Sub

Private Sub dtcdenopuesto_Click(Area As Integer)
    dtccodpuesto.Text = dtcdenopuesto.BoundText
End Sub

Private Sub DtCDescripcion_Click(Area As Integer)
    DtCCodigo.BoundText = DtCDescripcion.BoundText
    dtcunidadmedida.BoundText = DtCDescripcion.BoundText
    dtccodmanejo.BoundText = DtCDescripcion.BoundText
    dtccodpeso.BoundText = DtCDescripcion.BoundText
End Sub


'Private Sub dtc_partida15_Click(Area As Integer)
'    dtc_desc15.BoundText = Dtc_partida15.BoundText
'    dtc_unimed15.BoundText = Dtc_partida15.BoundText
'    dtc_stocktotal15.BoundText = Dtc_partida15.BoundText
'    dtc_grupo15.BoundText = Dtc_partida15.BoundText
'    dtc_subgrupo15.BoundText = Dtc_partida15.BoundText
'    dtc_codigo15.BoundText = Dtc_partida15.BoundText
''    dtc_precioventafinal15.BoundText = Dtc_partida15.BoundText
''    dtc_precioventabase15.BoundText = Dtc_partida15.BoundText
''    dtc_preciocompra15.BoundText = Dtc_partida15.BoundText
'End Sub

Private Sub dtcdescripar_Click(Area As Integer)
    dtccodpar.Text = dtcdescripar.BoundText
End Sub

Private Sub dtcdescripuni_Click(Area As Integer)
    dtccoduni.Text = dtcdescripuni.BoundText
End Sub

Private Sub dtcdescrtipoid_Click(Area As Integer)
    dtccodtipoid.BoundText = dtcdescrtipoid.BoundText
End Sub

Private Sub dtcfechacompromiso_Click(Area As Integer)
    dtccorrcompromiso.BoundText = dtcfechacompromiso.BoundText
End Sub

Private Sub dtcfechasol_Click(Area As Integer)
    dtccorrsol.BoundText = dtcfechasol.BoundText
End Sub

Private Sub dtcnroruc_Click(Area As Integer)
    dtcdenominacionruc.Text = dtcnroruc.BoundText
End Sub

'Private Sub dtc_desc2_LostFocus()
'    'If AdoBeneficiario.Recordset!beneficiario_deudor = "SI" Then
'    If Dtc_deudor2.Text = "SI" Then
'        Dtc_deudor2.backColor = &HFF&
'    Else
'        Dtc_deudor2.backColor = &H80000010
'    End If
'
'End Sub

Private Sub dtc_desc4A_Click(Area As Integer)
    dtc_codigo4A.BoundText = dtc_desc4A.BoundText
End Sub

Private Sub dtctipodoc_Click(Area As Integer)
    dtcdenodoc.Text = dtctipodoc.BoundText
End Sub

Private Sub dtcunidadmedida_Click(Area As Integer)
    DtCCodigo.BoundText = dtcunidadmedida.BoundText
    DtCDescripcion.BoundText = dtcunidadmedida.BoundText
    dtccodmanejo.BoundText = dtcunidadmedida.BoundText
    dtccodpeso.BoundText = dtcunidadmedida.BoundText
End Sub

Private Sub dtcdespoa_Click(Area As Integer)
    dtccodpoa.Text = dtcdespoa.BoundText
End Sub

Private Sub dtc_desc2A_Click(Area As Integer)
    dtc_codigo2A.BoundText = dtc_desc2A.BoundText
End Sub

Private Sub dtc_desc5_LostFocus()
    dtc_codigo5.Text = VAR_BEN3
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    'dtc_edificio6.BoundText = dtc_desc6.BoundText
    'dtc_reciboCobr6.BoundText = dtc_desc6.BoundText
    dtc_recibo6.BoundText = dtc_desc6.BoundText
    dtc_codigo6.BoundText = dtc_desc6.BoundText
    'dtc_fecha6.BoundText = dtc_desc6.BoundText
End Sub

'Private Sub dtc_edificio6_Click(Area As Integer)
'    dtc_desc6.BoundText = dtc_edificio6.BoundText
'    dtc_reciboCobr6.BoundText = dtc_edificio6.BoundText
'    dtc_recibo6.BoundText = dtc_edificio6.BoundText
'    dtc_codigo6.BoundText = dtc_edificio6.BoundText
'    dtc_fecha6.BoundText = dtc_edificio6.BoundText
'End Sub

'Private Sub dtc_fecha6_Click(Area As Integer)
'    dtc_desc6.BoundText = dtc_fecha6.BoundText
'    dtc_reciboCobr6.BoundText = dtc_fecha6.BoundText
'    dtc_recibo6.BoundText = dtc_fecha6.BoundText
'    dtc_codigo6.BoundText = dtc_fecha6.BoundText
'    dtc_edificio6.BoundText = dtc_fecha6.BoundText
'End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
    dtc_aux8.BoundText = dtc_desc8.BoundText
End Sub

Private Sub dtc_moneda22_Click(Area As Integer)
    dtc_desc22.BoundText = dtc_moneda22.BoundText
    dtc_codigo22.BoundText = dtc_moneda22.BoundText
End Sub

Private Sub dtc_recibo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_recibo6.BoundText
    'dtc_reciboCobr6.BoundText = dtc_recibo6.BoundText
    'dtc_fecha6.BoundText = dtc_recibo6.BoundText
    dtc_codigo6.BoundText = dtc_recibo6.BoundText
    'dtc_edificio6.BoundText = dtc_recibo6.BoundText
End Sub

'Private Sub dtc_reciboCobr6_Click(Area As Integer)
'    dtc_desc6.BoundText = dtc_reciboCobr6.BoundText
'    dtc_recibo6.BoundText = dtc_reciboCobr6.BoundText
'    dtc_fecha6.BoundText = dtc_reciboCobr6.BoundText
'    dtc_codigo6.BoundText = dtc_reciboCobr6.BoundText
'    dtc_edificio6.BoundText = dtc_reciboCobr6.BoundText
'End Sub

Private Sub Form_Load()
    frmMain.ProgressBar1.Visible = False
    buscados = 0
    swnuevo = 0
    accion = ""
    VAR_SW = ""
    lbl_cerrado = ""
    SWFILTRO = 0
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        usuario2 = rs_aux3!beneficiario_codigo
        VAR_BENEF = rs_aux3!beneficiario_codigo
        VAR_DA = rs_aux3!da_codigo
    Else
        usuario2 = "4908774"
        VAR_BENEF = "4908774"
        VAR_DA = "1.3"
    End If
    VAR_ORIGEN = Aux
     Select Case VAR_DA
        Case "1.8"    'Cochabamba
            VAR_DPTO = "3"
            parametro = "DPAGB"
        Case "1.7"    'Santa Cruz
            VAR_DPTO = "7"
            parametro = "DPAGS"
        Case "1.3", "1.4", "1.5"    'La Paz
            VAR_DPTO = "2"
            parametro = "DTESO"
        Case "1.9"    ' Chuquisaca
            VAR_DPTO = "1"
            parametro = "DPAGC"
        Case Else    ' OTRO
            VAR_DPTO = "2"
            parametro = "DTESO"
     End Select
    'REVISAR PARA TODOS LOS DOCS................
'    VAR_R = Aux     '"R-644"
    
    'Call CARGAPARAM
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    Call AbrirOrigen
    'Usuario
    lbl_cerrado.Caption = ""
    FrmDetalle.Caption = "ORIGEN - Ordenes de Pago a Proveedores " '+ Str((IIf(IsNull(Ado_datos.Recordset!correl_doc), 0, Ado_datos.Recordset!correl_doc)))
    'FrmDetalle.Caption = "DETALLE DE PAGOS - ORDEN DE PAGO NRO. 0"         '+ VAR_BIEN
    'aw_almacen_salida.Caption = "" + VAR_BIEN
    
    mbDataChanged = False
    FrmCabecera.Visible = False
    dg_datos.Enabled = True
    GlNombFor = "F04"

    marca1 = 1
    deta2 = 0
    swgrabar = 0
    swnuevo = 0
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False

    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption

  
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
'    'UNIDAD EJECUTORA
'    Set rs_datos1 = New ADODB.Recordset
'    If rs_datos1.State = 1 Then rs_datos1.Close
'    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
'    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
'    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    'Documentos de Respaldo                 OK
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_documentos_respaldo ", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText

    'Beneficiario Funcionario - Quien Entrega       OK
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' AND estado_codigo_resp = 'APR' order by beneficiario_denominacion", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText

    'Beneficiario Funcionario - Quien Recibe        OK
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from rv_unidad_vs_responsable where unidad_codigo = 'DCONT' AND estado_codigo = 'APR' order by beneficiario_denominacion", db, adOpenStatic
    'rs_datos5.Open "Select * from gc_beneficiario where tipoben_codigo = '1' and estado_codigo = 'APR' order by beneficiario_denominacion", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText

    'fc_cuenta_bancaria - Origen
    Set rs_datos21 = New ADODB.Recordset
    If rs_datos21.State = 1 Then rs_datos21.Close
    rs_datos21.Open "select * from fc_cuenta_bancaria   ", db, adOpenStatic
    Set Ado_datos21.Recordset = rs_datos21
    dtc_desc21.BoundText = dtc_codigo21.BoundText
    
    'gc_beneficiario_vs_cta_banco - Destino
    Set rs_datos22 = New ADODB.Recordset
    If rs_datos22.State = 1 Then rs_datos22.Close
    'rs_datos22.Open "select * from gc_beneficiario_vs_cta_banco ", db, adOpenStatic
    rs_datos22.Open "select * from fv_beneficiario_vs_cta_banco ", db, adOpenStatic
    Set Ado_datos22.Recordset = rs_datos22
    dtc_desc22.BoundText = dtc_codigo22.BoundText
    
    'Beneficiario Personas Nat. y Juridicas
     Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "select * from gc_beneficiario where tipoben_codigo > 1 AND estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    'dtc_desc2.BoundText = dtc_codigo2.BoundText
    
End Sub

Private Sub grabar()
  'db.BeginTrans
    If swgrabar = 1 Then
'        Set rs_aux4 = New ADODB.Recordset
'        SQL_FOR = "Select max(correl_doc) as Codigo from fo_traspaso_bancos_Egresos where doc_codigo = '" & VAR_ORIGEN & "' "
'        rs_aux4.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If Not rs_aux4.EOF Then
'            var_cod = IIf(IsNull(rs_aux4!Codigo), 1, rs_aux4!Codigo + 1)
'            db.Execute "Update gc_documentos_respaldo Set correl_doc = " & var_cod & " Where doc_codigo = '" & VAR_ORIGEN & "'   "
'        Else
'            var_cod = 1
'        End If
        var_cod = 0
        'CREA CABECERA
       VAR_R = Aux  '"R-644"
       'IdTraspasoBancos, clasif_codigo, doc_codigo, correl_doc, beneficiario_codigo_resp, beneficiario_codigo, unidad_codigo_resp, unidad_codigo, total_bs, total_dol ,
        'fecha_traspaso, cta_codigo, cta_codigo_destino, estado_conciliado, estado_codigo, usr_codigo, fecha_registro, hora_registro
        db.Execute "INSERT INTO fo_traspaso_bancos_Egresos (clasif_codigo, doc_codigo, correl_doc, beneficiario_codigo_resp, beneficiario_codigo, unidad_codigo_resp, unidad_codigo, total_bs, total_dol, " & _
            " fecha_traspaso, cta_codigo, cta_codigo_destino, estado_conciliado, estado_codigo, usr_codigo, fecha_registro, hora_registro) " & _
            " values ('" & dtc_aux3.Text & "', '" & dtc_codigo3.Text & "', " & var_cod & ", '" & dtc_codigo4 & "', '" & dtc_codigo5 & "', '" & parametro & "', '" & parametro & "', '0', '0',  " & _
            " '" & DTPfechasol & "', '" & dtc_codigo21.Text & "', '" & dtc_codigo22.Text & "', 'REG', 'REG', '" & glusuario & "', '" & Date & "', ''  ) "
    End If
    If swgrabar = 2 Then
        If Ado_datos.Recordset.RecordCount > 0 Then
            'INI ACTUALIZA
            db.Execute "UPDATE fo_traspaso_bancos_Egresos SET beneficiario_codigo_resp = '" & dtc_codigo4 & "', usr_codigo = '" & glusuario & "', fecha_traspaso = '" & DTPfechasol & "', beneficiario_codigo = '" & dtc_codigo5.Text & "', cta_codigo = '" & dtc_codigo21.Text & "', cta_codigo_destino= '" & dtc_codigo22.Text & "' WHERE IdTraspasoBancos = " & Ado_datos.Recordset!IdTraspasoBancos & " "
        End If

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'  If glPersNew = "P" Then
'    frmmo_formulario_M1.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre = rs_Personal!pers_nombres
'    frmmo_formulario_M1.Dtc_Pers_Cargo = rs_Personal!cargo_codigo
'  End If
'  glPersNew = "N"

End Sub


Private Sub OptFilGral1_Click()
   '===== Proceso para filtrado general de datos(registros NO aprobados)
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_datos6.RecordCount > 0 Then
        VAR_BENI = rs_datos6!beneficiario_codigo
    End If
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case glusuario
        Case "ADMIN", "NPAREDES", "RCUELA"              '"SQUISPE", "ASANTIVAÑEZ",
            queryinicial = "select * From fv_traspaso_bancos_Egresos WHERE (estado_codigo = 'REG')  "
        Case "VPAREDES", "MWILDE", "MVALDIVIA"
            'queryinicial = "select * From fv_traspaso_bancos_Egresos WHERE (estado_codigo = 'REG' AND (beneficiario_codigo_resp ='" & VAR_BENI & "' OR beneficiario_codigo_resp ='6962804')) "
            queryinicial = "select * From fv_traspaso_bancos_Egresos WHERE (estado_codigo = 'REG')  "
        Case "FCABRERA", "FDELGADILLO", "ASANTIVAÑEZ"
            queryinicial = "select * From fv_traspaso_bancos_Egresos WHERE (estado_codigo = 'REG' AND  beneficiario_codigo_resp ='" & VAR_BENI & "') "
        Case "TCASTILLO", "RVALDIVIEZO"
            queryinicial = "select * From fv_traspaso_bancos_Egresos WHERE (estado_codigo = 'REG' AND  beneficiario_codigo_resp ='" & VAR_BENI & "') "
        Case "EVILLALOBOS"
            queryinicial = "select * From fv_traspaso_bancos_Egresos WHERE (estado_codigo = 'REG' AND  beneficiario_codigo_resp ='" & VAR_BENI & "') "
'        Case "PRODAS"
'            queryinicial = "select * From fv_traspaso_bancos_Egresos WHERE (estado_codigo = 'REG' AND  beneficiario_codigo_resp ='" & VAR_BENI & "') "
        Case Else
            queryinicial = "select * From fv_traspaso_bancos_Egresos WHERE (estado_codigo = 'REG' AND (beneficiario_codigo_resp ='" & VAR_BENI & "' OR beneficiario_codigo_resp ='6962804')) "
    End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "IdTraspasoBancos"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset

End Sub

Private Sub OptFilGral2_Click()
 '===== Proceso para filtrado general de datos (todos los registros )
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_datos6.RecordCount > 0 Then
        VAR_BENI = rs_datos6!beneficiario_codigo
    End If
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case glusuario
        Case "ADMIN", "NPAREDES", "RCUELA"              '"SQUISPE", "ASANTIVAÑEZ",
            queryinicial = "select * From fv_traspaso_bancos_Egresos   "
        Case "VPAREDES", "MWILDE", "MVALDIVIA"
            queryinicial = "select * From fv_traspaso_bancos_Egresos   "
        Case "FCABRERA", "FDELGADILLO", "ASANTIVAÑEZ"
            queryinicial = "select * From fv_traspaso_bancos_Egresos WHERE (beneficiario_codigo_resp ='" & VAR_BENI & "') "
        Case "TCASTILLO", "RVALDIVIEZO"
            queryinicial = "select * From fv_traspaso_bancos_Egresos WHERE (beneficiario_codigo_resp ='" & VAR_BENI & "') "
        Case "EVILLALOBOS"
            queryinicial = "select * From fv_traspaso_bancos_Egresos WHERE (beneficiario_codigo_resp ='" & VAR_BENI & "') "
'        Case "PRODAS"
'            queryinicial = "select * From fv_traspaso_bancos_Egresos WHERE (beneficiario_codigo_resp ='" & VAR_BENI & "') "
        Case Else
            queryinicial = "select * From fv_traspaso_bancos_Egresos WHERE (beneficiario_codigo_resp ='" & VAR_BENI & "' OR beneficiario_codigo_resp ='6962804') "
    End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "IdTraspasoBancos"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset

End Sub

'Private Sub Option1_Click()
'    Fra_Total.Visible = True
'End Sub
'
'Private Sub Option2_Click()
'    FrmCobranza.Visible = True
'End Sub

Private Sub TxtCantPedi_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtcaracteristicas_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtMonto_bolivianos_contra_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub TxtMonto_bolivianos_contra_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If (Len(Trim(TxtMonto_bolivianos_contra.Text)) > 0) Then
       Txtmonto_dolares_contra.Text = IIf(TxtMonto_bolivianos_contra.Text > 0, TxtMonto_bolivianos_contra.Text / TxtTipo_cambio, 0)
    Else
       Txtmonto_dolares_contra.Text = 0
    End If
  End If
End Sub

Private Sub TxtMonto_bolivianos_KeyPress(KeyAscii As Integer)
'solo numeros y , .
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
    Else
      KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub txtjustifica_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtMonto_bolivianos_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If (Len(Trim(TxtMonto_bolivianos.Text)) > 0) Then
       Txtmonto_dolares.Text = IIf(TxtMonto_bolivianos.Text > 0, TxtMonto_bolivianos.Text / TxtTipo_cambio, 0)
    Else
       Txtmonto_dolares.Text = 0
    End If
  End If

End Sub

Private Sub Txtmonto_dolares_contra_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtmonto_dolares_contra_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If Len(Trim(Txtmonto_dolares_contra.Text)) > 0 Then
      TxtMonto_bolivianos_contra.Text = IIf(Txtmonto_dolares_contra.Text > 0, Txtmonto_dolares_contra * TxtTipo_cambio, 0)
    Else
      TxtMonto_bolivianos_contra.Text = 0
    End If
  End If
End Sub

Private Sub Txtmonto_dolares_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtmonto_dolares_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If Len(Trim(Txtmonto_dolares.Text)) > 0 Then
      TxtMonto_bolivianos.Text = IIf(Txtmonto_dolares.Text > 0, Txtmonto_dolares * TxtTipo_cambio, 0)
    Else
      TxtMonto_bolivianos.Text = 0
    End If
  End If
End Sub

Private Sub Txtobservaciones_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtsolpeso_KeyPress(KeyAscii As Integer)
'solo numeros y , .
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then

    Else
      KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

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

Private Sub cerea()
  txt_venta = " "
  dtc_codigo4.Text = " "
  Dtcpaternosol.Text = " "  'dtc_codigo4.BoundText
'  dtcmaternosol.Text = " "
'  dtcnombresol.Text = " "
  txtCantTotal = "0"
  TxtMontoBs = "0"
  TxtMontoUs = "0"
  TxtConcepto = ""
  dtc_codigo2 = ""
  dtc_desc2 = ""
  txtTDC.Text = GlTipoCambioOficial

'  DtCDenominacion_moneda = ""
'  TxtMonto_bolivianos = 0
'  Txtmonto_dolares = 0
'  TxtMonto_bolivianos_contra = 0
'  Txtmonto_dolares_contra = 0
'  DtCOrg_descripcion = ""
'  txtjustifica = ""
'  txt_venta = ""
'  txtterref = ""
End Sub
'Private Sub fbuscaunidad()
'  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
'  rstFc_unidad_ejecutora.Open "select * from Fc_unidad_ejecutora where uni_codigo = '" & Trim(adopuestosol.Recordset("codigo_unidad")) & "'", db, adOpenKeyset, adLockReadOnly
'  If rstFc_unidad_ejecutora.RecordCount > 0 Then
'    LblUni_descripcion_larga.Caption = rstFc_unidad_ejecutora("Uni_descripcion_larga")
'  Else
'    LblUni_descripcion_larga.Caption = ""
'  End If
'  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
'End Sub

Sub creaVista()
db.Execute "drop view vwF04"

db.Execute "create view vwF04 as " & _
            "select  ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.tipoben_codigo, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.grado_instruccion, ao_solicitud_lista.profesion, ao_solicitud_lista.paterno, ao_solicitud_lista.materno, ao_solicitud_lista.nombres, ao_solicitud_lista.telefono, ao_solicitud_lista.razon_s, ao_solicitud.codigo_solicitud, ao_solicitud.justificacion_solicitud, ao_solicitud.duracion_estimada_numero, ao_solicitud.por_tiempo, ao_solicitud.fecha_estimada_inicio, ao_solicitud.caracteristicas, ao_solicitud.duracion_estimada_tiempo, " & _
            "ao_solicitud.tr_adjuntos AS docAdjunta, " & _
            "ao_solicitud.codigo_bien, ac_bienes.bie_descripcion , ao_solicitud.observaciones, fc_unidad_ejecutora.uni_descripcion_larga, ao_solicitud.fecha_solicitud, " & _
            "(rc_personal.paterno) + ' ' + (rc_personal.materno) + ' ' +(rc_personal.nombres) + ' [' + ao_solicitud.ci + ']' AS pmn " & _
            "from ao_solicitud_lista  ,     " & _
                 "ao_solicitud       ,     " & _
                 "fc_unidad_ejecutora,     " & _
                 "rc_personal,             " & _
                 "ac_bienes                " & _
            "where  ao_solicitud_lista.ges_Gestion       = '" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
                    "ao_solicitud_lista.codigo_unidad    = '" & Me.Ado_datos.Recordset!codigo_unidad & "' and " & _
                    "ao_solicitud_lista.codigo_solicitud =  " & Me.Ado_datos.Recordset!codigo_solicitud & " and " & _
                    "ao_solicitud_lista.ges_Gestion      = ao_solicitud.ges_gestion            and " & _
                    "ao_solicitud_lista.codigo_unidad    = ao_solicitud.codigo_unidad          and " & _
                    "ao_solicitud_lista.codigo_solicitud = ao_solicitud.codigo_solicitud       and " & _
                    "ao_solicitud.codigo_unidad          = fc_unidad_ejecutora.codigo_unidad   and " & _
                    "ao_solicitud.codigo_bien            = ac_bienes.codigo_bien               and " & _
                    "ao_solicitud.ci                     = rc_personal.ci                      " & _
            "GROUP BY ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.tipoben_codigo, " & _
            "ao_solicitud.codigo_solicitud, ao_solicitud_lista.grado_instruccion, ao_solicitud_lista.profesion, ao_solicitud_lista.razon_s, ao_solicitud_lista.paterno, ao_solicitud_lista.materno, ao_solicitud_lista.nombres, " & _
            "ao_solicitud_lista.telefono, ao_solicitud.justificacion_solicitud, ao_solicitud.duracion_estimada_tiempo, ao_solicitud.nacional_extranjero, ao_solicitud.por_tiempo, ao_solicitud.codigo_bien, ac_bienes.bie_descripcion, ao_solicitud.duracion_estimada_numero, ao_solicitud.duracion_estimada_tiempo, ao_solicitud.fecha_estimada_inicio, ao_solicitud.esparaRH, ao_solicitud.tr_adjuntos, ao_solicitud.observaciones, ao_solicitud.caracteristicas, fc_unidad_ejecutora.Uni_descripcion_larga, ao_solicitud.fecha_solicitud, (rc_personal.paterno)+' '+(rc_personal.materno)+' '+(rc_personal.nombres)+' ['+ao_solicitud.ci+']', ao_solicitud_lista.id_beneficiario "

'            "trim$(rc_personal.paterno) + ' ' + trim$(rc_personal.materno) + ' ' +trim$(rc_personal.nombres) + ' [' + ao_solicitud.ci + ']' AS pmn " & _

'''db.Execute "create view vwF05 as " & _
'''            "select  ao_solicitud_lista.* " & _
'''            "from ao_solicitud_lista"
End Sub

Sub CREAVISTAF11()
db.Execute "drop view VWF11"
db.Execute "create view VWF11 as " & _
    "SELECT ao_Solicitud.Ges_Gestion, ao_Solicitud.codigo_unidad, " & _
    "ao_Solicitud.codigo_solicitud, ao_Solicitud.formulario, " & _
    "ao_Solicitud.justificacion_solicitud, ao_Solicitud.CI, " & _
    "ao_Solicitud.fecha_solicitud, ao_Solicitud.codigo_bien, " & _
    "ac_bienes_grupo.DescGrupo, RC_Personal.paterno, RC_Personal.materno, RC_Personal.nombres, " & _
    "ao_Solicitud.observaciones, ao_Solicitud.caracteristicas, " & _
    "ao_Solicitud.tr_adjuntos, ao_Solicitud.estatus, ao_Solicitud.estado_aprobacion, " & _
    "ao_Solicitud.duracion_estimada_numero, ao_Solicitud.duracion_estimada_tiempo, " & _
    "ao_solicitud_lista.codDetalle AS ci_material,  ao_solicitud_lista.profesion, ao_solicitud_lista.Aplanilla, " & _
    "ao_solicitud_lista.razon_s, ao_solicitud_lista.Nro_pagos, ao_solicitud_lista.Monto_solicitud_dl, ao_solicitud_lista.AUnidad " & _
"FROM ao_Solicitud, ao_Solicitud_detalle, ac_bienes_grupo, RC_Personal, ao_solicitud_lista " & _
"WHERE (ao_Solicitud.Ges_Gestion) = '" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    "(ao_Solicitud.codigo_unidad) = '" & Me.Ado_datos.Recordset!codigo_unidad & "' and " & _
    "(ao_Solicitud.codigo_solicitud) =  " & Me.Ado_datos.Recordset!codigo_solicitud & " and " & _
    "ao_Solicitud.Ges_Gestion = ao_Solicitud_detalle.Ges_Gestion AND " & _
    "ao_Solicitud.codigo_unidad = ao_Solicitud_detalle.codigo_unidad AND " & _
    "ao_Solicitud.codigo_solicitud = ao_Solicitud_detalle.codigo_solicitud AND " & _
    "ao_Solicitud.codigo_unidad = ao_Solicitud_lista.codigo_unidad AND " & _
    "ao_Solicitud.codigo_solicitud = ao_Solicitud_lista.codigo_solicitud AND " & _
    "ao_Solicitud.CodGrupo = ac_bienes_grupo.CodGrupo AND " & _
    "ao_Solicitud.ci = RC_Personal.ci"
End Sub

Private Sub acumulaMont(ges, Nro)
  Set rstacumdet = New ADODB.Recordset
  If rstacumdet.State = 1 Then rstacumdet.Close
  Set rs_datos19 = New ADODB.Recordset
  If rs_datos19.State = 1 Then rs_datos19.Close
'  LblGestion
'  lblcorrelVenta
'  lblNroVenta
  'rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as VAR_COBR2 from ao_ventas_detalle where ges_gestion = '" & ges & "' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
  rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot0 from ao_ventas_detalle where venta_codigo = " & Nro & " and par_codigo = '43340'", db, adOpenKeyset, adLockOptimistic
  If IsNull(rstacumdet!totbs) Then
    VAR_AUX = 0
    VAR_AUX2 = 0
    VAR_CANT = 1
  Else
    VAR_AUX = Round(rstacumdet!totbs, 2)
    VAR_AUX2 = Round(rstacumdet!totdl, 2)
    VAR_CANT = rstacumdet!cantot0
  End If

  'rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza_prog where ges_gestion = '" & ges & "' and estado_codigo = 'APR' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
  rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza where estado_codigo = 'APR' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  If IsNull(rs_datos19!totbs2) Then
    Cobrobs = 0
    VAR_COBR = 0
  Else
    Cobrobs = Round(rs_datos19!totbs2, 2)
    VAR_COBR = Round(rs_datos19!totdl2, 2)
  End If

  VAR_Bs = VAR_AUX - Cobrobs
  VAR_Dol = VAR_AUX2 - VAR_COBR
  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & VAR_AUX & " , ao_ventas_cabecera.venta_monto_total_dol = " & VAR_AUX2 & ", ao_ventas_cabecera.venta_cantidad_total = " & VAR_CANT & ", ao_ventas_cabecera.venta_monto_cobrado_bs = " & Cobrobs & ", ao_ventas_cabecera.venta_monto_cobrado_dol = " & VAR_COBR & ",  ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_Bs & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_Dol & "  Where ao_ventas_cabecera.venta_codigo = " & Nro & " "

  TxtMontoBs.Text = VAR_AUX
  TxtCobrado.Text = Cobrobs
  TxtBstotal.Text = VAR_Bs

'  If IsNull(Ado_datos.Recordset!venta_monto_cobrado_bs) Then
'    Ado_datos.Recordset!venta_monto_cobrado_bs = 0
'    VAR_AUX = Ado_datos.Recordset!venta_monto_total_bs
'  Else
'    VAR_AUX = Ado_datos.Recordset!venta_monto_total_bs - Ado_datos.Recordset!venta_monto_cobrado_bs
'  End If
'  If VAR_AUX > 0 Then
'        VAR_AUX2 = VAR_AUX / Ado_datos.Recordset!venta_tipo_cambio
'  Else
'        VAR_AUX2 = 0
'  End If
'  'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.monto_total_Bs = " & rstacumdet!totbs & " , ao_ventas_cabecera.monto_cobrado = " & rstacumdet!totbs & ", ao_ventas_cabecera.monto_total_Us = " & rstacumdet!totdl & ", ao_ventas_cabecera.cantidad_total_vendida = " & rstacumdet!cantot & ", ao_ventas_cabecera.saldo_p_cobrar = ao_ventas_cabecera.monto_total_Bs - ao_ventas_cabecera.deuda_cobrada Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & nro & " "
'  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & rstacumdet!totbs & " , ao_ventas_cabecera.venta_monto_total_dol = " & rstacumdet!totdl & ", ao_ventas_cabecera.venta_cantidad_total = " & rstacumdet!cantot & ", ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_AUX & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_AUX2 & "  Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & nro & " "
'
'  TxtMontoBs = rstacumdet!totbs
'  TxtCobrado = rs_datos19!totbs2    'IIf(IsNull(Ado_datos.Recordset("venta_monto_cobrado_bs")), 0, Ado_datos.Recordset("venta_monto_cobrado_bs"))
'  If IsNull(Ado_datos.Recordset("venta_saldo_p_cobrar_bs")) Then
'    Text2 = VAR_AUX 'Ado_datos.Recordset("venta_monto_total_bs") - Ado_datos.Recordset("venta_monto_cobrado_bs")
'    Ado_datos.Recordset("venta_saldo_p_cobrar_bs") = VAR_AUX
'  Else
'    Text2 = Ado_datos.Recordset("venta_saldo_p_cobrar_bs")
'  End If

  If rstacumdet.State = 1 Then rstacumdet.Close

  'Print ado_datos14.Recordset!ges_gestion
  'Print ado_datos14.Recordset!correl_venta
  'Print ado_datos14.Recordset!venta_codigo
  'ado_datos14.Recordset!monto_Bolivianos = rstacumdet!totbs
  'ado_datos14.Recordset!monto_dolares = rstacumdet!totdl
  'ado_datos14.Recordset.Update
'  Set rstdestino = New ADODB.Recordset
'  If rstdestino.State = 1 Then rstdestino.Close
'  rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & ges & "' and correl_venta = '" & corr & "' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
'  If rstdestino.RecordCount > 0 Then
'    rstdestino!monto_total_Bs = rstacumdet!totbs
'    rstdestino!monto_cobrado = rstacumdet!totbs
'    rstdestino!monto_total_Us = rstacumdet!totdl
'    rstdestino!cantidad_total_vendida = rstacumdet!cantot
'    rstdestino!saldo_p_cobrar = 0
'    rstdestino.Update
'  End If
'  'Set Ado_datos.Recordset = rstdestino
'  If rstdestino.State = 1 Then rstdestino.Close
'  If rstacumdet.State = 1 Then rstacumdet.Close
End Sub

Private Sub Picture2_Click()
    If DctMonto18.Text <> Ado_datos11.Recordset!adjudica_bs Then
        sino = MsgBox("El Importe del Extracto es DIFERENTE al pago solicitado, esta seguro de Aceptar el registro ?", vbYesNo, "Confirmando")
        If sino = vbYes Then
        Else
            Exit Sub
        End If
    End If
    Text11.Text = DctCod18.Text
    DTP_Finicio.Value = Format(CDate(DctFecha18.Text), "DD/MM/YYYY")
    Text12.Text = Trim(DctDeposita18.Text) + " " + Trim(DctOrigina18.Text)
    
    db.Execute "update fo_recibos_detalle_egresos set cmpbte_deposito_bco = '" & Text11.Text & "', fecha_registro_bco = '" & DTP_Finicio & "', fecha_destino = '" & Date & "', observaciones = '" & Text12.Text & "'  where adjudica_codigo = " & Ado_datos11.Recordset!adjudica_codigo & " "
    Fra_reporte.Visible = False
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
    FraNavega.Enabled = True
    Call AbrirDetalle
    
'    db.Execute "update fo_recibos_detalle_egresos set cmpbte_deposito_bco = '" & Text11.Text & "', fecha_registro_bco = '" & DTP_Finicio & "', fecha_destino = '" & Date & "', observaciones = '" & Text12.Text & "'  where adjudica_codigo = " & Ado_datos11.Recordset!adjudica_codigo & " "
'    Fra_reporte.Visible = False
'    Call AbrirOrigen
'    Call AbrirDetalle
End Sub

'Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
'    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
'End Sub

Private Sub TxtCobrado_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Then      '(KeyAscii = 8) Or '(0..9)
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub TxtDsctoTot_LostFocus()
    If TxtDsctoTot.Text = "" Or TxtDsctoTot.Text = "0" Or TxtDsctoTot.Text = "0.00" Then
        TxtMonto.Text = "0"
    Else
        TxtMonto.Text = Round(CDbl(TxtDsctoTot.Text) * GlTipoCambioMercado, 2)
    End If
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
  '? . , 09
  ',.01234856789
End Sub

Private Sub TxtMonto_LostFocus()
    If TxtMonto.Text = "" Or TxtMonto.Text = "0" Or TxtMonto.Text = "0.00" Then
        TxtDsctoTot.Text = "0"
    Else
        TxtDsctoTot.Text = Round(CDbl(TxtMonto.Text) / GlTipoCambioMercado, 2)
    End If
End Sub

Private Sub TxtPlazo_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]" Or KeyAscii = 8, KeyAscii, 0)
End Sub


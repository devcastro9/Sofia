VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fw_nota_credito_debito 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Procesos Financieros - Tesorer�a - Notas de Credito-Debito"
   ClientHeight    =   10410
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   17115
   Icon            =   "fw_nota_credito_debito.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   1.57266e6
   ScaleMode       =   0  'User
   ScaleWidth      =   1.79561e7
   WindowState     =   2  'Maximized
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      Picture         =   "fw_nota_credito_debito.frx":0A02
      ScaleHeight     =   1755
      ScaleMode       =   0  'User
      ScaleWidth      =   1875
      TabIndex        =   120
      Top             =   7800
      Width           =   1935
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1395
      Left            =   120
      Picture         =   "fw_nota_credito_debito.frx":6CA34
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   119
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Frame FrmDetalle 
      BackColor       =   &H00000000&
      Caption         =   "DATOS DE LA VENTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   1575
      Left            =   2160
      TabIndex        =   88
      Top             =   6105
      Width           =   14775
      Begin MSDataGridLib.DataGrid dg_datos16 
         Bindings        =   "fw_nota_credito_debito.frx":D8A66
         Height          =   1170
         Left            =   120
         TabIndex        =   89
         Top             =   240
         Width           =   14520
         _ExtentX        =   25612
         _ExtentY        =   2064
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483633
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
         ColumnCount     =   19
         BeginProperty Column00 
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cite.Tramite"
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
            DataField       =   "edif_codigo"
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
         BeginProperty Column02 
            DataField       =   "edif_descripcion"
            Caption         =   "Denominacion del Edificio"
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
            DataField       =   "zona_denominacion"
            Caption         =   "Zona"
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
            DataField       =   "calle_tipo"
            Caption         =   "Via.Acceso"
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
            DataField       =   "calle_denominacion"
            Caption         =   "Nombre de Calle, Av u otro"
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
            DataField       =   "edif_nro"
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
         BeginProperty Column07 
            DataField       =   "venta_codigo"
            Caption         =   "Nro.Venta"
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
            DataField       =   "beneficiario_denominacion"
            Caption         =   "Cliente/Representante.Legal"
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
            DataField       =   "venta_fecha_inicio"
            Caption         =   "F.Inicio.Contrato"
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
            DataField       =   "venta_fecha_fin"
            Caption         =   "F.Fin.Contrato"
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
            DataField       =   "venta_cantidad_total"
            Caption         =   "Cantidad"
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
            DataField       =   "unimed_codigo"
            Caption         =   "Periodicidad"
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
            DataField       =   "venta_monto_total_bs"
            Caption         =   "Total,Contrato.Bs"
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
            DataField       =   "venta_monto_cobrado_bs"
            Caption         =   "Cobrado.Bs"
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
         BeginProperty Column15 
            DataField       =   "venta_saldo_p_cobrar_bs"
            Caption         =   "Saldo.P/Cobar"
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
            DataField       =   "unidad_codigo"
            Caption         =   "Unidad.E."
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
            DataField       =   "solicitud_codigo"
            Caption         =   "No.Tramite"
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
            Caption         =   "Estado"
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
               Alignment       =   2
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   3495.118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2789.858
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2564.788
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   3060.284
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   675.213
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5910
      Left            =   120
      TabIndex        =   10
      Top             =   45
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   10425
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   0
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "REGISTROS FACTURADOS"
      TabPicture(0)   =   "fw_nota_credito_debito.frx":D8A80
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraNavega1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrmCobros1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraGrabarCancelar1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraOpciones1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "NOTAS CREDITO-DEBITO"
      TabPicture(1)   =   "fw_nota_credito_debito.frx":D8A9C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmCobros"
      Tab(1).Control(1)=   "FraNavega"
      Tab(1).Control(2)=   "frm_benef"
      Tab(1).Control(3)=   "FraGrabarCancelar"
      Tab(1).Control(4)=   "fraOpciones"
      Tab(1).ControlCount=   5
      Begin VB.PictureBox fraOpciones1 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   16635
         TabIndex        =   112
         Top             =   360
         Width           =   16635
         Begin VB.PictureBox BtnSalir1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   14040
            Picture         =   "fw_nota_credito_debito.frx":D8AB8
            ScaleHeight     =   615
            ScaleWidth      =   1245
            TabIndex        =   118
            ToolTipText     =   "Cierra la Ventana Activa"
            Top             =   0
            Width           =   1245
         End
         Begin VB.PictureBox BtnImprimir5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4200
            Picture         =   "fw_nota_credito_debito.frx":D927A
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   117
            ToolTipText     =   "Re-Imprimir Factura"
            Top             =   0
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.PictureBox BtnBuscar1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2880
            Picture         =   "fw_nota_credito_debito.frx":D9B47
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   116
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox BtnAprobar1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1440
            Picture         =   "fw_nota_credito_debito.frx":DA2FC
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   115
            ToolTipText     =   "Envia a Nota Credito-Debito"
            Top             =   0
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   7080
            Picture         =   "fw_nota_credito_debito.frx":DAB2F
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   114
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5625
            Picture         =   "fw_nota_credito_debito.frx":DB27B
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   113
            Top             =   0
            Visible         =   0   'False
            Width           =   1430
         End
         Begin VB.Label lbl_titulo1 
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
            Left            =   9495
            TabIndex        =   94
            Top             =   195
            Width           =   1815
         End
      End
      Begin VB.PictureBox FraGrabarCancelar1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   60
         ScaleHeight     =   675
         ScaleWidth      =   15360
         TabIndex        =   108
         Top             =   360
         Visible         =   0   'False
         Width           =   15360
         Begin VB.PictureBox BtnGrabar1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2880
            Picture         =   "fw_nota_credito_debito.frx":DBB90
            ScaleHeight     =   615
            ScaleWidth      =   1305
            TabIndex        =   110
            Top             =   0
            Width           =   1300
         End
         Begin VB.PictureBox BtnCancelar1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4275
            Picture         =   "fw_nota_credito_debito.frx":DC366
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   109
            Top             =   0
            Width           =   1400
         End
         Begin VB.Label lbl_titulo3 
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
            Left            =   10695
            TabIndex        =   111
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.PictureBox fraOpciones 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   -74940
         ScaleHeight     =   660
         ScaleWidth      =   16680
         TabIndex        =   95
         Top             =   360
         Width           =   16680
         Begin VB.CommandButton CmdFoto 
            BackColor       =   &H00808000&
            Caption         =   "&Reporte"
            Height          =   720
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   107
            ToolTipText     =   "Carga Imagen QR"
            Top             =   0
            Visible         =   0   'False
            Width           =   740
         End
         Begin VB.CommandButton BtnImprimir2 
            BackColor       =   &H00C0C000&
            Caption         =   "Recibo"
            Height          =   720
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   106
            ToolTipText     =   "Imprime Recibo"
            Top             =   0
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton BtnDesAprobar 
            BackColor       =   &H00808080&
            Height          =   600
            Left            =   8400
            Picture         =   "fw_nota_credito_debito.frx":DCC52
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   0
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CommandButton BtnVer 
            BackColor       =   &H00808000&
            Caption         =   "Digitaliza"
            Height          =   600
            Left            =   7320
            Picture         =   "fw_nota_credito_debito.frx":DCE5C
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Guarda en Archivo Digital"
            Top             =   0
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.PictureBox BtnA�adir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   9600
            Picture         =   "fw_nota_credito_debito.frx":DD29E
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   102
            Top             =   0
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.PictureBox BtnModificar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   105
            Picture         =   "fw_nota_credito_debito.frx":DDA5D
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   101
            Top             =   0
            Visible         =   0   'False
            Width           =   1430
         End
         Begin VB.PictureBox BtnEliminar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            Picture         =   "fw_nota_credito_debito.frx":DE372
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   100
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.PictureBox BtnAprobar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5760
            Picture         =   "fw_nota_credito_debito.frx":DEABE
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   99
            Top             =   0
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.PictureBox BtnBuscar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2880
            Picture         =   "fw_nota_credito_debito.frx":DF2F1
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   98
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox BtnImprimir3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4320
            Picture         =   "fw_nota_credito_debito.frx":DFAA6
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   97
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox BtnSalir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   14040
            Picture         =   "fw_nota_credito_debito.frx":E0373
            ScaleHeight     =   615
            ScaleWidth      =   1245
            TabIndex        =   96
            ToolTipText     =   "Cierra la Ventana Activa"
            Top             =   0
            Width           =   1245
         End
         Begin VB.Label lbl_titulo 
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
            Left            =   10575
            TabIndex        =   105
            Top             =   195
            Width           =   1815
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
         Left            =   -74940
         ScaleHeight     =   675
         ScaleWidth      =   15360
         TabIndex        =   90
         Top             =   360
         Visible         =   0   'False
         Width           =   15360
         Begin VB.PictureBox BtnCancelar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4275
            Picture         =   "fw_nota_credito_debito.frx":E0B35
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   92
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox BtnGrabar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2880
            Picture         =   "fw_nota_credito_debito.frx":E1421
            ScaleHeight     =   615
            ScaleWidth      =   1305
            TabIndex        =   91
            Top             =   0
            Width           =   1300
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
            Left            =   10695
            TabIndex        =   93
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.Frame frm_benef 
         BackColor       =   &H00404040&
         Caption         =   "Elije un Nuevo Beneficiario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   2295
         Left            =   -67800
         TabIndex        =   76
         Top             =   3480
         Visible         =   0   'False
         Width           =   9495
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8520
            TabIndex        =   77
            Top             =   735
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton BtnCancelarBen 
            BackColor       =   &H00E0E0E0&
            Height          =   675
            Left            =   4440
            MaskColor       =   &H00000000&
            Picture         =   "fw_nota_credito_debito.frx":E1BF7
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Cancelar"
            Top             =   1200
            Width           =   1485
         End
         Begin VB.CommandButton BtnGrabarBen 
            BackColor       =   &H00E0E0E0&
            Height          =   675
            Left            =   3000
            Picture         =   "fw_nota_credito_debito.frx":E25D1
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1200
            Width           =   1365
         End
         Begin MSDataListLib.DataCombo dtc_codigo8 
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6720
            TabIndex        =   78
            Top             =   1080
            Visible         =   0   'False
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "00000000000004"
         End
         Begin MSDataListLib.DataCombo dtc_desc8 
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   480
            TabIndex        =   6
            Top             =   720
            Width           =   6240
            _ExtentX        =   11007
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   14737632
            ForeColor       =   0
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux8 
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6720
            TabIndex        =   84
            Top             =   720
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_nit"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "00000000000004"
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Factura a Nombre de:"
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
            Left            =   480
            TabIndex        =   79
            Top             =   465
            Width           =   2025
         End
      End
      Begin VB.Frame FrmCobros1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4740
         Left            =   7365
         TabIndex        =   48
         Top             =   1080
         Width           =   9375
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   9045
            TabIndex        =   121
            Top             =   3000
            Width           =   255
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6855
            TabIndex        =   82
            Top             =   2335
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_codigo4A1 
            Bindings        =   "fw_nota_credito_debito.frx":E2E91
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   5655
            TabIndex        =   81
            Top             =   2325
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "12345678901234"
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            CausesValidation=   0   'False
            DataField       =   "cobranza_observaciones"
            DataSource      =   "Ado_datos01"
            ForeColor       =   &H00000000&
            Height          =   585
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   1580
            Width           =   6015
         End
         Begin VB.TextBox TxtMonto1 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
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
            Left            =   7005
            TabIndex        =   1
            Text            =   "0"
            Top             =   3495
            Width           =   1275
         End
         Begin VB.TextBox TxtMontoDol1 
            Alignment       =   2  'Center
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
            Left            =   7005
            TabIndex        =   9
            Text            =   "0"
            Top             =   3915
            Width           =   1275
         End
         Begin MSComCtl2.DTPicker DTPFechaSol 
            DataField       =   "cobranza_fecha_sol"
            DataSource      =   "Ado_datos01"
            Height          =   300
            Left            =   3540
            TabIndex        =   0
            Top             =   4395
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   119668737
            CurrentDate     =   41678
         End
         Begin MSDataListLib.DataCombo dtc_desc4A1 
            Bindings        =   "fw_nota_credito_debito.frx":E2EAB
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   1080
            TabIndex        =   80
            Top             =   2325
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo DataCombo14 
            Bindings        =   "fw_nota_credito_debito.frx":E2EC5
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   7365
            TabIndex        =   122
            Top             =   2985
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "12345678901234"
         End
         Begin MSDataListLib.DataCombo DataCombo13 
            Bindings        =   "fw_nota_credito_debito.frx":E2EDE
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   2280
            TabIndex        =   123
            Top             =   2985
            Width           =   5400
            _ExtentX        =   9525
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Lbl_nombre_fac3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Factura a Nombre de..."
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
            TabIndex        =   124
            Top             =   3000
            Width           =   2040
         End
         Begin VB.Line Line11 
            BorderColor     =   &H00C00000&
            X1              =   0
            X2              =   7365
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line9 
            BorderColor     =   &H00C00000&
            X1              =   7365
            X2              =   7365
            Y1              =   0
            Y2              =   2810
         End
         Begin VB.Label cmd_fac 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "FACTURA"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
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
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   7845
            TabIndex        =   75
            Top             =   4275
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label TxtMontoDol0 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_total_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos01"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1485
            TabIndex        =   74
            Top             =   3915
            Width           =   1515
         End
         Begin VB.Label Label51 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_total_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos01"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   7620
            TabIndex        =   73
            Top             =   1380
            Width           =   1710
         End
         Begin VB.Label Label50 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Facturado USD"
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
            Left            =   7560
            TabIndex        =   72
            Top             =   1095
            Width           =   1785
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Facturado Dol:                                                          Monto Credito-Debito Dol (13%):"
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
            TabIndex        =   71
            Top             =   3945
            Width           =   6795
         End
         Begin VB.Line Line10 
            BorderColor     =   &H00C00000&
            X1              =   0
            X2              =   9615
            Y1              =   2805
            Y2              =   2805
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Facturado Bs.:                                                          Monto Credito-Debito Bs (13%):"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   70
            Top             =   3525
            Width           =   6690
         End
         Begin VB.Label TxtMonto0 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_total_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos01"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1485
            TabIndex        =   69
            Top             =   3495
            Width           =   1515
         End
         Begin VB.Label TxtDsctoTot1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_total_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos01"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   7620
            TabIndex        =   68
            Top             =   480
            Width           =   1710
         End
         Begin VB.Label DTPFechaProg1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_fecha_fac"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            DataSource      =   "Ado_datos01"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   7620
            TabIndex        =   67
            Top             =   2235
            Width           =   1710
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "venta_codigo"
            DataSource      =   "Ado_datos01"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4080
            TabIndex        =   66
            Top             =   255
            Width           =   1125
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Facturaci�n"
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
            Left            =   7560
            TabIndex        =   65
            Top             =   1965
            Width           =   1785
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "doc_codigo_fac"
            DataSource      =   "Ado_datos01"
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
            Left            =   180
            TabIndex        =   64
            Top             =   1125
            Width           =   1245
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_prog_codigo"
            DataSource      =   "Ado_datos01"
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
            Left            =   6480
            TabIndex        =   63
            Top             =   255
            Width           =   735
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Cobranza:                      Nro.Venta:                     Nro.Cuota:"
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
            Left            =   120
            TabIndex        =   62
            Top             =   270
            Width           =   6270
         End
         Begin VB.Label Label24 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto:"
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
            Top             =   1635
            Width           =   960
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Facturado BOB"
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
            Left            =   7560
            TabIndex        =   60
            Top             =   195
            Width           =   1785
         End
         Begin VB.Label Label20 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Cobrador de CGI:"
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
            Height          =   480
            Left            =   120
            TabIndex        =   59
            Top             =   2240
            Width           =   960
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "       Fecha Solicitud.Nota.Credito-Debito"
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
            Left            =   -240
            TabIndex        =   58
            Top             =   4410
            Width           =   3585
         End
         Begin VB.Label lbl_factura1 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.de Factura"
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
            Left            =   2460
            TabIndex        =   57
            Top             =   825
            Width           =   1605
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de Autorizaci�n"
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
            Left            =   5040
            TabIndex        =   56
            Top             =   825
            Width           =   1875
         End
         Begin VB.Label lbl_doc01 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "doc_codigo"
            DataSource      =   "Ado_datos01"
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
            Left            =   1635
            TabIndex        =   55
            Top             =   660
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label lbl_docnro1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "doc_numero"
            DataSource      =   "Ado_datos01"
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
            Left            =   5085
            TabIndex        =   54
            Top             =   660
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "C�digo Registro"
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
            Left            =   120
            TabIndex        =   53
            Top             =   825
            Width           =   1470
         End
         Begin VB.Label Txt_cod_cobro1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_codigo"
            DataSource      =   "Ado_datos01"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1635
            TabIndex        =   52
            Top             =   255
            Width           =   1125
         End
         Begin VB.Label TxtCmpbte1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_nro_factura"
            DataSource      =   "Ado_datos01"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   320
            Left            =   2580
            TabIndex        =   51
            Top             =   1125
            Width           =   1320
         End
         Begin VB.Label TxtAutorizacion1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_nro_autorizacion"
            DataSource      =   "Ado_datos01"
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
            Left            =   4845
            TabIndex        =   50
            Top             =   1125
            Width           =   2235
         End
      End
      Begin VB.Frame FraNavega 
         BackColor       =   &H00000000&
         Caption         =   "LISTA"
         ForeColor       =   &H00FFFFC0&
         Height          =   4750
         Left            =   -74960
         TabIndex        =   44
         Top             =   1080
         Width           =   7050
         Begin VB.OptionButton OptFilGral2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Facturados y No Cobrados"
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
            Left            =   2400
            TabIndex        =   46
            Top             =   4395
            Width           =   2595
         End
         Begin VB.OptionButton OptFilGral1 
            BackColor       =   &H00FFFFC0&
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
            Left            =   840
            TabIndex        =   45
            Top             =   4395
            Value           =   -1  'True
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid dg_datos 
            Height          =   4020
            Left            =   75
            TabIndex        =   47
            Top             =   240
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   7091
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   -2147483633
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   13
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   12
            BeginProperty Column00 
               DataField       =   "dc_fecha"
               Caption         =   "F.Nota.C-D"
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
               DataField       =   "edif_codigo"
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
            BeginProperty Column02 
               DataField       =   "cobranza_codigo"
               Caption         =   "No.Cobranza"
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
               DataField       =   "beneficiario_codigo_resp"
               Caption         =   "Cobrador"
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
               DataField       =   "cobranza_fecha_fac"
               Caption         =   "F.Facturacion"
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
               DataField       =   "cobranza_total_bs"
               Caption         =   "Facturad.Bs."
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
               DataField       =   "cobranza_total_dol"
               Caption         =   "Nota.D-C.Bs"
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
            BeginProperty Column07 
               DataField       =   "doc_numero"
               Caption         =   "Nro.Doc.Respaldo"
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
               DataField       =   "cobranza_nro_factura"
               Caption         =   "Nro. Factura"
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
               DataField       =   "estado_codigo_fac1"
               Caption         =   "Estado_Fact1"
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
            BeginProperty Column10 
               DataField       =   "estado_codigo_fac"
               Caption         =   "Estado"
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
               DataField       =   "beneficiario_codigo"
               Caption         =   "NIT/CI del Cliente"
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
                  ColumnWidth     =   1110.047
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1154.835
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column08 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   1335.118
               EndProperty
               BeginProperty Column10 
                  Alignment       =   2
                  ColumnWidth     =   645.165
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Ado_datos 
            Height          =   330
            Left            =   75
            Top             =   4320
            Width           =   6900
            _ExtentX        =   12171
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
            BackColor       =   -2147483633
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
      Begin VB.Frame FrmCobros 
         BackColor       =   &H00C0C0C0&
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
         Height          =   4740
         Left            =   -67875
         TabIndex        =   17
         Top             =   1080
         Width           =   9615
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "monto_dc_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   8040
            TabIndex        =   132
            Text            =   "0"
            Top             =   3960
            Width           =   1275
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "cobranza_tdc"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   6960
            TabIndex        =   131
            Text            =   "6.96"
            Top             =   3960
            Width           =   795
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            DataField       =   "monto_dc_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
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
            Height          =   285
            Left            =   5400
            TabIndex        =   130
            Text            =   "0"
            Top             =   3960
            Width           =   1275
         End
         Begin VB.TextBox TxtMontoDol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "cobranza_total_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2685
            Locked          =   -1  'True
            TabIndex        =   128
            Text            =   "0"
            Top             =   3960
            Width           =   1275
         End
         Begin VB.TextBox txt_tdc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "cobranza_tdc"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   127
            Text            =   "6.96"
            Top             =   3960
            Width           =   795
         End
         Begin VB.TextBox TxtMonto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "cobranza_total_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   165
            Locked          =   -1  'True
            TabIndex        =   126
            Text            =   "0"
            Top             =   3960
            Width           =   1275
         End
         Begin VB.TextBox TxtCmpbte 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "cobranza_nro_factura"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   2640
            TabIndex        =   87
            Text            =   "0"
            Top             =   1060
            Width           =   1455
         End
         Begin VB.PictureBox Picture2 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   1500
            Left            =   5280
            ScaleHeight     =   1500
            ScaleMode       =   0  'User
            ScaleWidth      =   1411.765
            TabIndex        =   85
            Top             =   960
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7080
            TabIndex        =   19
            Top             =   2485
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_codigo4A 
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5640
            TabIndex        =   20
            Top             =   2475
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "12345678901234"
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8820
            TabIndex        =   18
            Top             =   3045
            Width           =   255
         End
         Begin VB.TextBox TxtObs 
            CausesValidation=   0   'False
            DataField       =   "cobranza_observaciones"
            DataSource      =   "Ado_datos"
            Height          =   585
            Left            =   1080
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   1620
            Width           =   6255
         End
         Begin VB.CommandButton cmd_benef 
            BackColor       =   &H00808000&
            Height          =   320
            Left            =   9105
            Picture         =   "fw_nota_credito_debito.frx":E2EF7
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Buscar Beneficiario"
            Top             =   3030
            Width           =   375
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2085
            TabIndex        =   4
            Top             =   3030
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux5 
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7275
            TabIndex        =   21
            Top             =   3030
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "beneficiario_nit"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "00000000000004"
         End
         Begin MSDataListLib.DataCombo dtc_desc4A 
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1455
            TabIndex        =   22
            Top             =   2475
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4920
            TabIndex        =   23
            Top             =   2760
            Visible         =   0   'False
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "00000000000004"
         End
         Begin MSComCtl2.DTPicker DTPFechaCobro 
            DataField       =   "cobranza_fecha_fac"
            DataSource      =   "Ado_datos"
            Height          =   300
            Left            =   7620
            TabIndex        =   3
            Top             =   2415
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   119668737
            CurrentDate     =   41678
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   $"fw_nota_credito_debito.frx":E38F9
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
            TabIndex        =   129
            Top             =   3720
            Width           =   9225
         End
         Begin VB.Label lbl_nit 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "beneficiario_codigo"
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
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1080
            TabIndex        =   125
            Top             =   3120
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label lbl_doc1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "doc_codigo"
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
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1635
            TabIndex        =   38
            Top             =   660
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Cobranza:                      Nro.Venta:                      Nro.Cuota:"
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
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   6330
         End
         Begin VB.Label TxtAutorizacion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_nro_autorizacion"
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   5160
            TabIndex        =   43
            Top             =   1080
            Width           =   2115
         End
         Begin VB.Label Txt_cod_cobro 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_codigo"
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   1635
            TabIndex        =   42
            Top             =   255
            Width           =   1245
         End
         Begin VB.Label Lbl_nombre_fac 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Factura a Nombre de:                                                                                                      NIT/CI"
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
            TabIndex        =   41
            Top             =   3045
            Width           =   7110
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   7455
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "C�digo Registro"
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
            Index           =   4
            Left            =   120
            TabIndex        =   40
            Top             =   800
            Width           =   1470
         End
         Begin VB.Label lbl_docnro 
            Alignment       =   2  'Center
            BackColor       =   &H80000015&
            Caption         =   "0"
            DataField       =   "doc_numero"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   2685
            TabIndex        =   39
            Top             =   1065
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de Autorizaci�n"
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
            Left            =   5280
            TabIndex        =   37
            Top             =   795
            Width           =   1875
         End
         Begin VB.Label lbl_factura 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.de Factura"
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
            Left            =   2700
            TabIndex        =   36
            Top             =   795
            Width           =   1485
         End
         Begin VB.Label lbl_fechas 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Facturaci�n"
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
            Left            =   7560
            TabIndex        =   35
            Top             =   2160
            Width           =   1785
         End
         Begin VB.Label Lbl_Cobrador 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Cobrador CGI:"
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
            TabIndex        =   34
            Top             =   2490
            Width           =   1275
         End
         Begin VB.Label Label46 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Solicitado BOB"
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
            Left            =   7560
            TabIndex        =   33
            Top             =   75
            Width           =   1785
         End
         Begin VB.Label lbl_obs 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto:"
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
            TabIndex        =   32
            Top             =   1755
            Width           =   960
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_prog_codigo"
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   6480
            TabIndex        =   31
            Top             =   255
            Width           =   855
         End
         Begin VB.Label lbl_fac 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "doc_codigo_fac"
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   420
            TabIndex        =   30
            Top             =   1060
            Width           =   1005
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Solicitud Fac"
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
            Left            =   7560
            TabIndex        =   29
            Top             =   1410
            Width           =   1785
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFF80&
            X1              =   7455
            X2              =   7455
            Y1              =   0
            Y2              =   2855
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   7455
            Y1              =   660
            Y2              =   660
         End
         Begin VB.Label TxtNroVentaC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "venta_codigo"
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4080
            TabIndex        =   28
            Top             =   255
            Width           =   1125
         End
         Begin VB.Label DTPFechaProg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_fecha_sol"
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
            Height          =   300
            Left            =   7620
            TabIndex        =   27
            Top             =   1680
            Width           =   1710
         End
         Begin VB.Label TxtDsctoTot 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_solicitado_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7620
            TabIndex        =   26
            Top             =   345
            Width           =   1710
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   9615
            Y1              =   3465
            Y2              =   3465
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00FFFF80&
            X1              =   7455
            X2              =   9615
            Y1              =   2085
            Y2              =   2085
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Solicitado USD"
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
            Left            =   7560
            TabIndex        =   25
            Top             =   720
            Width           =   1785
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_solicitado_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7620
            TabIndex        =   24
            Top             =   975
            Width           =   1710
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00FFFF80&
            X1              =   7455
            X2              =   9615
            Y1              =   2850
            Y2              =   2850
         End
      End
      Begin VB.Frame FraNavega1 
         BackColor       =   &H00000000&
         Caption         =   "LISTA"
         ForeColor       =   &H00FFFFC0&
         Height          =   4755
         Left            =   40
         TabIndex        =   13
         Top             =   1080
         Width           =   7290
         Begin VB.OptionButton OptFilGral01 
            Caption         =   "Facturados"
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
            Left            =   720
            TabIndex        =   16
            Top             =   4395
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.OptionButton OptFilGral02 
            Caption         =   "Enviados a Nota Credito-Debito"
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
            Left            =   2160
            TabIndex        =   15
            Top             =   4395
            Visible         =   0   'False
            Width           =   2835
         End
         Begin MSDataGridLib.DataGrid dg_datos1 
            Height          =   4020
            Left            =   75
            TabIndex        =   14
            Top             =   240
            Width           =   7140
            _ExtentX        =   12594
            _ExtentY        =   7091
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   -2147483633
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   13
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   12
            BeginProperty Column00 
               DataField       =   "cobranza_fecha_fac"
               Caption         =   "Fecha.Fac."
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
               DataField       =   "edif_codigo"
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
            BeginProperty Column02 
               DataField       =   "cobranza_codigo"
               Caption         =   "Cod.Registro"
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
               DataField       =   "beneficiario_codigo_resp"
               Caption         =   "Cobrador"
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
               DataField       =   "cobranza_fecha_sol"
               Caption         =   "F.Solicit.Fac."
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
               DataField       =   "cobranza_total_bs"
               Caption         =   "Facturado.Bs."
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
               DataField       =   "cobranza_total_dol"
               Caption         =   "Cobrado en Dol."
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
            BeginProperty Column07 
               DataField       =   "doc_numero"
               Caption         =   "Nro.Doc.Respaldo"
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
               DataField       =   "cobranza_nro_factura"
               Caption         =   "Nro.Factura"
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
               DataField       =   "estado_codigo_fac1"
               Caption         =   "Estado_Fac1"
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
            BeginProperty Column10 
               DataField       =   "estado_codigo_sol"
               Caption         =   "Estado"
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
               DataField       =   "beneficiario_codigo"
               Caption         =   "NIT/CI del Cliente"
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
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1065.26
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1275.024
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column08 
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1244.976
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   1244.976
               EndProperty
               BeginProperty Column10 
                  Alignment       =   2
                  ColumnWidth     =   780.095
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Ado_datos01 
            Height          =   330
            Left            =   75
            Top             =   4320
            Width           =   7140
            _ExtentX        =   12594
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
            BackColor       =   -2147483633
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
   End
   Begin VB.Frame FrmCobranza 
      BackColor       =   &H00000000&
      Caption         =   "DETALLE DE BIENES / SERVICIOS VENDIDOS"
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
      Height          =   1965
      Left            =   2160
      TabIndex        =   11
      Top             =   7725
      Width           =   14775
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "fw_nota_credito_debito.frx":E3982
         Height          =   1620
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   2858
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483633
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "venta_codigo"
            Caption         =   "Nro.Venta"
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
            DataField       =   "bien_codigo"
            Caption         =   "Codigo.Bien"
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
            DataField       =   "concepto_venta"
            Caption         =   "Descripcion y Caracter�sticas del Bien"
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
            DataField       =   "venta_det_cantidad"
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
         BeginProperty Column04 
            DataField       =   "venta_precio_unitario_bs"
            Caption         =   "Prec.Unitario"
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
         BeginProperty Column05 
            DataField       =   "venta_descuento_bs"
            Caption         =   "Descuento"
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
            DataField       =   "venta_precio_total_bs"
            Caption         =   "Precio Total"
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
         BeginProperty Column07 
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo.Vendido"
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
            DataField       =   "almacen_codigo"
            Caption         =   "Almacen"
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
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   4334.74
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   1874.835
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   750.047
            EndProperty
         EndProperty
      End
   End
   Begin Crystal.CrystalReport CryV01 
      Left            =   240
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6840
      Top             =   10080
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2280
      Top             =   10080
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
   Begin MSAdodcLib.Adodc ado_datos14 
      Height          =   330
      Left            =   0
      Top             =   10800
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
      Left            =   9120
      Top             =   10440
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
      Top             =   10440
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
      Left            =   2280
      Top             =   10800
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
      Left            =   6840
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
   Begin MSAdodcLib.Adodc AdoDsctos 
      Height          =   330
      Left            =   11400
      Top             =   10080
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
      Caption         =   "AdoDsctos"
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
      Left            =   2280
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
      Left            =   4560
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
      Left            =   13680
      Top             =   10080
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
      Left            =   4560
      Top             =   10080
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
      Left            =   0
      Top             =   10080
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
   Begin MSAdodcLib.Adodc ado_datos4A 
      Height          =   330
      Left            =   9120
      Top             =   10080
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
      Caption         =   "ado_datos4A"
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
      Left            =   720
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
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
   Begin MSAdodcLib.Adodc Ado_datos20 
      Height          =   330
      Left            =   4560
      Top             =   10800
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
   Begin Crystal.CrystalReport CryF01 
      Left            =   1200
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   6840
      Top             =   10800
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   9120
      Top             =   10800
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
   Begin MSAdodcLib.Adodc Ado_datos7 
      Height          =   330
      Left            =   11400
      Top             =   10440
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
      Caption         =   "Ado_datos7"
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
      Left            =   13080
      Top             =   9720
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
   Begin Crystal.CrystalReport CryF02 
      Left            =   1680
      Top             =   9840
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
   Begin Crystal.CrystalReport CryQ01 
      Left            =   2160
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
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
   Begin VB.PictureBox Picture1 
      Height          =   1560
      Left            =   13680
      ScaleHeight     =   1500
      ScaleWidth      =   1695
      TabIndex        =   86
      Top             =   4080
      Visible         =   0   'False
      Width           =   1755
   End
End
Attribute VB_Name = "fw_nota_credito_debito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ventas
'INI QR
'Enum TQRCodeEncoding
'ceALPHA
'ceBYTE
'ceNUMERIC
'ceKANJI
'ceAUTO
'End Enum
'Enum TQRCodeECLevel
'LEVEL_L
'LEVEL_M
'LEVEL_Q
'LEVEL_H
'End Enum
'Private Declare Sub FullQRCode Lib "QRCodeLib.dll" _
'(ByVal autoConfigurate As Boolean, _
' ByVal AutoFit As Boolean, _
' ByVal backColor As Long, _
' ByVal barColor As Long, _
' ByVal Texto As String, _
' ByVal correctionLevel As TQRCodeECLevel, _
' ByVal encoding As TQRCodeEncoding, _
' ByVal marginpixels As Integer, _
' ByVal moduleWidth As Integer, _
' ByVal Height As Integer, _
' ByVal Width As Integer, _
' ByVal FileName As String)
'Private Declare Sub FastQRCode Lib "QRCodeLib.dll" _
'(ByVal Texto As String, _
' ByVal FileName As String)
'Private Declare Function QRCodeLibVer Lib "QRCodeLib.dll" () As String
'Dim sFile As String
'Dim CadenaQ As String
'FIN QR
Dim rs_datos As New ADODB.Recordset     'FACTURACION
Dim rs_datos01 As New ADODB.Recordset     'INICIO COBRANZAS
Dim rs_datos02 As New ADODB.Recordset     'REG. COBRANZAS
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos4A As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
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
Dim rs_datos20 As New ADODB.Recordset   'Cta Bancaria

Dim rs_Ventas_lista As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rstdestino As New ADODB.Recordset
Dim rstcorrel_ing As New ADODB.Recordset

'CLASIFICADORES
Dim rstdetsalalm As New ADODB.Recordset
Dim RS_BENEF As New ADODB.Recordset
Dim rs_TipoCambio As New ADODB.Recordset
Dim rs_almacen2 As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim rsAuxDetalle As New ADODB.Recordset
'IMAGENES
Dim m_stream    As ADODB.Stream
'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir As String
'Dim queryinicial As String
Dim queryinicial1 As String
Dim queryinicial2 As String

'Dim descri_bien As String
'VARIABLES
Dim iResult As Variant  ', i%, y%
Dim marca1 As Variant

Dim VAR_CANT As Integer         'Cant_Alm,
Dim correlativo1 As Integer
Dim swgrabar, swnuevo, deta2 As Integer
Dim nroventa, correlv, NRO_COBR As Integer
Dim VAR_PARTIDA, VAR_PROY, correldetalle As Integer
Dim VAR_CODANT, Var_Comp, VAR_SW, VAR_TSOL As Integer
Dim VAR_SOL As Integer
Dim i As Integer

Dim Cobrobs, VAR_COBR, VAR_AUX, VAR_AUX2 As Double
Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2, COBR_BS As Double
Dim VAR_CONTAB As Double

Dim gestion0, var_literal, VAR_PROY2, VAR_CTA, VAR_PROY3 As String
Dim VAR_CODTIPO, VAR_ORG, VAR_FTE, VAR_BENEF, VAR_GLOSA, VAR_MONEDA As String
Dim VAR_COD1, VAR_COD2, VAR_COD3 As String
Dim VAR_ANIO, VAR_MES, VAR_DIA, VAR_FECHA As String
Dim VAR_COD4, VAR_TIPOV, VAR_CITE  As String
Dim DESAUX, VARAUX, VARCODIG As String
Dim VAR_EST, VAR_FAC As String
Dim codigo_doc As String
Dim Numero As String
Dim Autorizacion As String
Dim NroFactura As String
Dim NitCi As String
Dim Fecha As String
Dim Monto As String
Dim Llave As String
Dim CodigoContro As String
'Dim Exel As New Excel.Application
Dim fs As FileSystemObject      'Variable de tipo file System Object

Private Sub CmdDetalle_Click()
    FrmCobranza.Visible = True
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  Dim descri_bien As String
'  Dim Cant_Alm As Integer
  If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then   'EOF
     If Not IsNull(Ado_datos.Recordset("venta_codigo")) Then            'venta_codigo
        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
            BtnModificar.Visible = True
            If Ado_datos.Recordset!doc_codigo_fac = "R-101" Then
               BtnImprimir3.Visible = True
'               BtnImprimir3.Caption = "Facturar"
               lbl_factura.Caption = "Nro.de Factura"
               TxtCmpbte.Visible = True
               TxtCmpbte.Locked = True
               lbl_docnro.Visible = False
               'TxtCmpbte.backColor = &H404040
               'TxtCmpbte.ForeColor = &HFFFFFF
               Lbl_nombre_fac.Caption = "Factura a Nombre de:                                                                                                 NIT/CI"
               lbl_fechas.Caption = "Fecha Facturaci�n"
            Else
               BtnImprimir3.Visible = False
'               BtnImprimir3.Caption = "Recibo"
               lbl_factura.Caption = "Nro.de Recibo"
               lbl_docnro.Visible = True
               TxtCmpbte.Visible = False
               'TxtCmpbte.Locked = False     ' CAMBIAR DE Objeto
               'TxtCmpbte.backColor = &H80000005
               'TxtCmpbte.ForeColor = &H80000008
               Lbl_nombre_fac.Caption = "Recibo a Nombre de:                                                                                                  NIT/CI"
               lbl_fechas.Caption = "Fecha de Recibo"
            End If
            If (Ado_datos.Recordset("cobranza_fecha_sol") <= Date - 16) Then
                TxtDsctoTot.backColor = &HFF&             'ROJO
                DTPFechaProg.backColor = &HFF&             'ROJO
            Else
                If (Ado_datos.Recordset("cobranza_fecha_sol") > Date - 16) And (Ado_datos.Recordset("cobranza_fecha_sol") <= Date - 1) Then
                    TxtDsctoTot.backColor = &H80FF&           'NARANJA
                    DTPFechaProg.backColor = &H80FF&           'NARANJA
                Else
                    TxtDsctoTot.backColor = &H404040        '&H80000013      'Fondo Oscuro
                    DTPFechaProg.backColor = &H404040       '&H80000013      'Fondo Oscuro
                End If
            End If
        Else
            BtnModificar.Visible = False
'            BtnEliminar.Visible = False
'            BtnAprobar.Visible = False
'            BtnVer.Visible = True
'            FrmABMDet.Visible = False
'            FrmABMDet2.Visible = True
'            FrmCobranza.Visible = True
            TxtDsctoTot.backColor = &H404040        '&H80000013      'Fondo Oscuro
            DTPFechaProg.backColor = &H404040       '&H80000013      'Fondo Oscuro
            BtnImprimir3.Visible = False
        End If

        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        'queryinicial2 = "select * from ao_ventas_detalle where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " and correl_venta = " & Ado_datos.Recordset!correl_venta & " "
        'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
        Set ado_datos14.Recordset = rs_datos14
        ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
        Else
            deta2 = 0
        End If
        
        Set rs_datos16 = New ADODB.Recordset
        If rs_datos16.State = 1 Then rs_datos16.Close
        rs_datos16.Open "select * from av_ventas_cabecera where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        Set Ado_datos16.Recordset = rs_datos16
        Ado_datos16.Recordset.Requery
        If Ado_datos16.Recordset.RecordCount > 0 Then
            VAR_PROY3 = Ado_datos16.Recordset!edif_codigo
            FrmCobranza.Visible = True
            'BtnImprimir2.Visible = True
            'BtnImprimir3.Visible = True
        Else
            FrmCobranza.Visible = False
            'BtnImprimir2.Visible = False
            'BtnImprimir3.Visible = False
        End If
        
        ''Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
        Set rs_datos5 = New ADODB.Recordset
        If rs_datos5.State = 1 Then rs_datos5.Close
        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
        Set Ado_datos5.Recordset = rs_datos5
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        dtc_aux5.BoundText = dtc_codigo5.BoundText
        
        FrmDetalle.Caption = "VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
        
        FrmCobranza.Caption = "DETALLE DE BIENES DE LA VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
        
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_ventas_cobranza Where cobranza_codigo = '" & Ado_datos.Recordset!cobranza_codigo & "' ", "Foto")
'        Image2 = Img_Foto
'        'If adoLista.Recordset!estado_codigo = "APR" Then
'        CmdFoto.Visible = True
     End If                         'venta_codigo
     FrmDetalle.Enabled = True
     FrmCobranza.Visible = True
  Else
    BtnImprimir3.Visible = False
'                BtnDesAprobar.Visible = True
    BtnModificar.Visible = False
'    BtnEliminar.Visible = False
'    BtnVer.Visible = False
    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
  End If                            'EOF
End Sub

Private Sub Ado_datos01_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If (Not Ado_datos01.Recordset.BOF) And (Not Ado_datos01.Recordset.EOF) Then   'EOF
     If Not IsNull(Ado_datos01.Recordset("venta_codigo")) Then            'venta_codigo
        If (Ado_datos01.Recordset("estado_codigo_fac") = "APR") Then          'FACTURADOS
            TxtMonto1.Text = Round(CDbl(TxtMonto0.Caption) * 0.13, 2)
            TxtMontoDol1.Text = Round(CDbl(TxtMontoDol0.Caption) * 0.13, 2)
'            If (Ado_datos01.Recordset("cobranza_fecha_prog") <= Date - 16) Then
'                TxtDsctoTot1.backColor = &HFF&             'ROJO
'                DTPFechaProg1.backColor = &HFF&             'ROJO
'            Else
'                If (Ado_datos01.Recordset("cobranza_fecha_prog") > Date - 16) And (Ado_datos01.Recordset("cobranza_fecha_prog") <= Date - 1) Then
'                    TxtDsctoTot1.backColor = &H80FF&           'NARANJA
'                    DTPFechaProg1.backColor = &H80FF&           'NARANJA
'                Else
'                    TxtDsctoTot1.backColor = &H404040        '&H80000013      'Fondo Oscuro
'                    DTPFechaProg1.backColor = &H404040       '&H80000013      'Fondo Oscuro
'                End If
'            End If
'            BtnModificar1.Visible = True
'            BtnAprobar1.Visible = True
'            If Ado_datos01.Recordset!doc_codigo_fac = "R-103" Then
'                cmd_fac = "RECIBO"
'                lbl_factura1 = "Nro. de Recibo"
'            Else
'                cmd_fac = "FACTURA"
'                lbl_factura1 = "Nro.de Factura"
'            End If
            If glusuario = "ADMIN" Or glusuario = "CSALINAS" Then
                BtnImprimir5.Visible = True
            Else
                BtnImprimir5.Visible = False
            End If
        Else
'            BtnModificar1.Visible = False
            BtnAprobar1.Visible = False
            TxtDsctoTot1.backColor = &H404040        '&H80000013      'Fondo Oscuro
            DTPFechaProg1.backColor = &H404040       '&H80000013      'Fondo Oscuro
        End If
'        If Ado_datos01.Recordset("beneficiario_codigo") <> "" Then
'            Set RS_BENEF = New ADODB.Recordset
'            If RS_BENEF.State = 1 Then RS_BENEF.Close
'            RS_BENEF.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos01.Recordset!beneficiario_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'            'RS_BENEF.Recordset.Requery
'            If RS_BENEF.RecordCount > 0 Then
'                If RS_BENEF!beneficiario_deudor = "SI" Then
'                    Dtc_deudor2.BackColor = &HFF&
'                Else
'                    Dtc_deudor2.BackColor = &H80000010
'                End If
'            End If
'
'        End If
        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos01.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        'queryinicial2 = "select * from ao_ventas_detalle where venta_codigo = " & Ado_datos01.Recordset!venta_codigo & " and correl_venta = " & Ado_datos01.Recordset!correl_venta & " "
        'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
        Set ado_datos14.Recordset = rs_datos14
        ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
            'TxtMontoBs.Text = Ado_datos01.Recordset!monto_total_bS
            'TxtMontoUs.Text = Ado_datos01.Recordset!deuda_cobrada
            'Text2.Text = Ado_datos01.Recordset!saldo_p_cobrar
            'Call AbreAlmacen
        Else
            deta2 = 0
'            'TxtMontoBs.Text = 0
'            'TxtMontoUs.Text = 0
'            'Text2.Text = 0
'            FrmABMDet2.Visible = False
'            FrmCobranza.Visible = False
        End If
        
        Set rs_datos16 = New ADODB.Recordset
        If rs_datos16.State = 1 Then rs_datos16.Close
        rs_datos16.Open "select * from av_ventas_cabecera where venta_codigo = '" & Ado_datos01.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        Set Ado_datos16.Recordset = rs_datos16
        Ado_datos16.Recordset.Requery
        If Ado_datos16.Recordset.RecordCount > 0 Then
            VAR_PROY3 = Ado_datos16.Recordset!edif_codigo
            FrmCobranza.Visible = True
            'BtnImprimir2.Visible = True
            'BtnImprimir3.Visible = True
        Else
            FrmCobranza.Visible = False
            'BtnImprimir2.Visible = False
            'BtnImprimir3.Visible = False
        End If
        
        ''Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
        Set rs_datos5 = New ADODB.Recordset
        If rs_datos5.State = 1 Then rs_datos5.Close
        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
        Set Ado_datos5.Recordset = rs_datos5
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        dtc_aux5.BoundText = dtc_codigo5.BoundText
        
        FrmDetalle.Caption = "VENTA NRO. " + Str((Ado_datos01.Recordset("venta_codigo")))
        
        FrmCobranza.Caption = "DETALLE DE BIENES DE LA VENTA NRO. " + Str((Ado_datos01.Recordset("venta_codigo")))
        
'        TxtCobrador1 = Trim(dtc_desc4A.Text)
        
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_ventas_cobranza Where cobranza_codigo = '" & Ado_datos01.Recordset!cobranza_codigo & "' ", "Foto")
'        Image2 = Img_Foto
'        'If adoLista.Recordset!estado_codigo = "APR" Then
'        CmdFoto.Visible = True
     End If                         'venta_codigo
     FrmDetalle.Enabled = True
     FrmCobranza.Visible = True
  Else
    BtnAprobar1.Visible = False
    BtnModificar1.Visible = False
    'BtnEliminar1.Visible = False

    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
  End If                            'EOF
End Sub


Private Sub AbreAlmacen()
'    Set rs_datos13 = New ADODB.Recordset
'    If rs_datos13.State = 1 Then rs_datos13.Close
'    'rs_datos13.Open "select * from Av_DestinoDet where coddetalle= '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    rs_datos13.Open "select * from Av_almacen_detalle where bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_datos13.Recordset = rs_datos13
'    Ado_datos13.Refresh

End Sub

Private Sub Ado_datos02_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If (Not Ado_datos02.Recordset.BOF) And (Not Ado_datos02.Recordset.EOF) Then   'EOF
     If Not IsNull(Ado_datos02.Recordset("venta_codigo")) Then            'venta_codigo
        If (Ado_datos02.Recordset("estado_codigo_bco") = "REG") Then          'REG
            If (Ado_datos02.Recordset("cobranza_fecha_prog") <= Date - 16) Then
                TxtDsctoTot1.backColor = &HFF&             'ROJO
                DTPFechaProg1.backColor = &HFF&             'ROJO
            Else
                If (Ado_datos02.Recordset("cobranza_fecha_prog") > Date - 16) And (Ado_datos02.Recordset("cobranza_fecha_prog") <= Date - 1) Then
                    TxtDsctoTot2.backColor = &H80FF&           'NARANJA
                    DTPFechaProg2.backColor = &H80FF&           'NARANJA
                Else
                    TxtDsctoTot2.backColor = &H404040        '&H80000013      'Fondo Oscuro
                    DTPFechaProg2.backColor = &H404040       '&H80000013      'Fondo Oscuro
                End If
            End If
            BtnModificar2.Visible = True
            BtnAprobar2.Visible = True
            OptFilGral05.Visible = False
            If (glusuario = "ADMIN" Or glusuario = "MVALDIVIA" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS") Then
                OptFilGral05.Visible = True
            Else
                OptFilGral05.Visible = False
            End If
            If Ado_datos02.Recordset!doc_codigo_fac = "R-103" Then
                lbl_factura3 = "Nro. de Recibo"
                Lbl_nombre_fac3.Caption = "Factura a Nombre de:                                                                                                 NIT/CI"
                lbl_fechas3.Caption = "Fecha de Recibo"
            Else
                lbl_factura3 = "Nro.de Factura"
            End If
        Else
            TxtDsctoTot2.backColor = &H404040        '&H80000013      'Fondo Oscuro
            DTPFechaProg2.backColor = &H404040       '&H80000013      'Fondo Oscuro
            If Ado_datos02.Recordset!estado_codigo = "APR" Then
                BtnAprobar.Visible = False
                BtnAprobar2.Visible = False
                BtnModificar2.Visible = False
                OptFilGral05.Visible = False
            Else
                If (glusuario = "ADMIN" Or glusuario = "HBUSTILLOS" Or glusuario = "RCUELA" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS") Then
                    BtnAprobar.Visible = True
                    BtnAprobar2.Visible = False
                    BtnModificar2.Visible = True
                    OptFilGral05.Visible = True
                Else
                    BtnAprobar.Visible = False
                    BtnAprobar2.Visible = False
                    BtnModificar2.Visible = False
                    OptFilGral05.Visible = False
                End If
            End If
        End If

'        If Ado_datos02.Recordset("beneficiario_codigo") <> "" Then
'            Set RS_BENEF = New ADODB.Recordset
'            If RS_BENEF.State = 1 Then RS_BENEF.Close
'            RS_BENEF.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos02.Recordset!beneficiario_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'            'RS_BENEF.Recordset.Requery
'            If RS_BENEF.RecordCount > 0 Then
'                If RS_BENEF!beneficiario_deudor = "SI" Then
'                    Dtc_deudor2.BackColor = &HFF&
'                Else
'                    Dtc_deudor2.BackColor = &H80000010
'                End If
'            End If
'
'        End If
        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos02.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        'queryinicial2 = "select * from ao_ventas_detalle where venta_codigo = " & Ado_datos02.Recordset!venta_codigo & " and correl_venta = " & Ado_datos02.Recordset!correl_venta & " "
        'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
        Set ado_datos14.Recordset = rs_datos14
        ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
            'TxtMontoBs.Text = Ado_datos02.Recordset!monto_total_bS
            'TxtMontoUs.Text = Ado_datos02.Recordset!deuda_cobrada
            'Text2.Text = Ado_datos02.Recordset!saldo_p_cobrar
            'Call AbreAlmacen
        Else
            deta2 = 0
'            'TxtMontoBs.Text = 0
'            'TxtMontoUs.Text = 0
'            'Text2.Text = 0
'            FrmABMDet2.Visible = False
'            FrmCobranza.Visible = False
        End If
        
        Set rs_datos16 = New ADODB.Recordset
        If rs_datos16.State = 1 Then rs_datos16.Close
        rs_datos16.Open "select * from av_ventas_cabecera where venta_codigo = '" & Ado_datos02.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        Set Ado_datos16.Recordset = rs_datos16
        Ado_datos16.Recordset.Requery
        If Ado_datos16.Recordset.RecordCount > 0 Then
            VAR_PROY3 = Ado_datos16.Recordset!edif_codigo
            FrmCobranza.Visible = True
            'BtnImprimir2.Visible = True
            'BtnImprimir3.Visible = True
        Else
            FrmCobranza.Visible = False
            'BtnImprimir2.Visible = False
            'BtnImprimir3.Visible = False
        End If
        
        ''Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
        Set rs_datos5 = New ADODB.Recordset
        If rs_datos5.State = 1 Then rs_datos5.Close
        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
        Set Ado_datos5.Recordset = rs_datos5
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        dtc_aux5.BoundText = dtc_codigo5.BoundText
        
        FrmDetalle.Caption = "VENTA NRO. " + Str((Ado_datos02.Recordset("venta_codigo")))
        
        FrmCobranza.Caption = "DETALLE DE BIENES DE LA VENTA NRO. " + Str((Ado_datos02.Recordset("venta_codigo")))
        
'        TxtCobrador1 = Trim(dtc_desc4A.Text)
        
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_ventas_cobranza Where cobranza_codigo = '" & Ado_datos02.Recordset!cobranza_codigo & "' ", "Foto")
'        Image2 = Img_Foto
'        'If adoLista.Recordset!estado_codigo = "APR" Then
'        CmdFoto.Visible = True
     End If                         'venta_codigo
     FrmDetalle.Enabled = True
     FrmCobranza.Visible = True
  Else
    BtnAprobar2.Visible = False
    BtnModificar2.Visible = False
    'BtnEliminar2.Visible = False

    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
  End If                            'EOF

End Sub

Private Sub Ado_datos16_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 If (Not Ado_datos16.Recordset.BOF) And (Not Ado_datos16.Recordset.EOF) Then
    If Not IsNull(Ado_datos16.Recordset("venta_codigo")) Then
'        BtnModDetalle.Visible = True
'        BtnImprimir4.Visible = True


    Else
        'BtnAprobar2.Visible = False
        'BtnImprimir2.Visible = False
'        BtnImprimir4.Visible = False
        'BtnAnlDetalle2.Visible = False
'        BtnModDetalle.Visible = False
    End If
 Else
    'BtnAprobar2.Visible = False
    'BtnImprimir2.Visible = False
    BtnImprimir4.Visible = False
    'BtnAnlDetalle2.Visible = False
'    BtnModDetalle.Visible = False
 End If
End Sub

Private Sub BntImprimir2_Click()
    'If Ado_datos.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
        CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_cobranzas_facturadas.rpt"
            'CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_diaria_facturas.rpt"
        CryF02.WindowShowRefreshBtn = True
            'CryF02.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            'CryF02.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            'CryF02.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            'CryF02.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
        CryF02.Formulas(1) = "titulo = 'MODULO DE COBRANZAS' "
        CryF02.Formulas(2) = "subtitulo = 'ESTADO DE CUENTAS' "
            iResult = CryF02.PrintReport
            If iResult <> 0 Then MsgBox CryF02.LastErrorNumber & " : " & CryF02.LastErrorString, vbCritical, "Error de impresi�n"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
     'End If
End Sub

Private Sub BtnA�adir_Click()
marca1 = Ado_datos.Recordset.Bookmark
  'If Ado_datos.Recordset!venta_tipo = "C" And Ado_datos.Recordset!estado_codigo = "APR" Then
  If Ado_datos.Recordset!venta_tipo = "C" Or Ado_datos.Recordset!venta_tipo = "V" Then
    If Ado_datos.Recordset!venta_saldo_p_cobrar_bs > 0 Then
    'If Ado_datos.Recordset!venta_monto_total_bs - Ado_datos.Recordset!venta_monto_cobrado_bs > 0 Then
        swnuevo = 1
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        FrmCobros.Visible = True
        FrmCobros.Enabled = True
'        fraOpciones.Enabled = False
        FraNavega.Enabled = False
        FrmDetalle.Visible = False
        FrmCobranza.Visible = False
        FrmABMDet.Visible = False
        FrmABMDet2.Visible = False
'        TxtCobrador.Visible = False
        Ado_datos16.Recordset.AddNew
        dtc_codigo2A.Text = dtc_codigo2.Text
        dtc_desc2A.Text = dtc_desc2.Text
        TxtMonto.SetFocus
        DTPFechaProg.Visible = True
        DTPFechaCobro.Visible = True
        Lbl_nombre_fac.Caption = "Cliente :"
        lbl_fechas.Caption = "Fecha Programada de la Cobranza"
        Txt_parche.Visible = True
        'Ado_datos.Recordset.Move marca1 - 1
    Else
        MsgBox "Ya se cobr� el total de la deuda, Verifique por favor !! ", vbExclamation, "Atenci�n!"
    End If
  Else
    MsgBox "La Venta (al Contado o Donaci�n) NO tiene saldo para cobrar, Verifique por favor !! ", vbExclamation, "Atenci�n!"
  End If
End Sub

Private Sub BtnAprobar_Click()
' If Ado_datos02.Recordset.RecordCount > 0 Then
'     If IsNull(Ado_datos02.Recordset("cobranza_observaciones")) Or (Ado_datos02.Recordset("cobranza_deuda_bs") = 0) Then
'        MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
'        Exit Sub
'     Else
'        If Ado_datos02.Recordset("estado_codigo") = "REG" Then
'           sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
'           If sino = vbYes Then
'               'If Ado_datos02.Recordset("venta_tipo") = "C" Or Ado_datos02.Recordset("venta_tipo") = "V" Then
'               '     db.Execute "update gc_beneficiario set beneficiario_deudor = 'SI' where beneficiario_codigo = '" & dtc_codigo2 & "' "
'               'End If
'               gestion0 = glGestion                 'Ado_datos02.Recordset("ges_gestion")
'               correlv = Ado_datos02.Recordset("venta_codigo")
'               nroventa = Ado_datos02.Recordset("venta_codigo")
'
'               VAR_BENEF = Ado_datos02.Recordset!beneficiario_codigo
'               VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
'               VAR_GLOSA = Trim(Ado_datos02.Recordset!cobranza_observaciones) '+ " - Nro.: " + Trim(VAR_CITE)
'               VAR_DOL2 = Round(Ado_datos02.Recordset!cobranza_deuda_dol, 2)
'               VAR_BS2 = Round(Ado_datos02.Recordset!cobranza_deuda_bs, 2)
'               VAR_CTA = IIf(Ado_datos02.Recordset!Cta_Codigo = "", "NN", Ado_datos02.Recordset!Cta_Codigo)
'
'               VAR_PROY2 = Ado_datos16.Recordset!edif_codigo
'               VAR_COD4 = Ado_datos16.Recordset!unidad_codigo
'               VAR_TIPOV = Ado_datos16.Recordset!venta_tipo
'               VAR_SOL = Ado_datos16.Recordset!solicitud_codigo
'
'
'               NRO_COBR = Me.Ado_datos02.Recordset!cobranza_codigo
'               var_literal = Ado_datos02.Recordset!Literal
'               VAR_MONEDA = Ado_datos02.Recordset!tipo_moneda
'            'Llave = Trim(rs_aux1!dosifica_llave)
'            'NitCi = Ado_datos.Recordset!beneficiario_codigo_fac     'VAR_BENEF
'            'Autorizacion = rs_aux1!dosifica_autorizacion
'
'               ' APRUEBA ao_ventas_cabecera
'               'db.Execute "update ao_ventas_cobranza set estado_codigo = 'APR' Where ges_gestion = '" & Ado_datos02.Recordset("ges_gestion") & "' And cobranza_codigo = " & Ado_datos02.Recordset("cobranza_codigo") & " "
'               db.Execute "update ao_ventas_cobranza set estado_codigo = 'APR' Where cobranza_codigo = " & NRO_COBR & " "
'
'                Set rs_aux2 = New ADODB.Recordset
'                If rs_aux2.State = 1 Then rs_aux2.Close
'                SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos02.Recordset!doc_codigo & "'  "
'                rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'                If rs_aux2.RecordCount > 0 Then
'                    rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'                    Ado_datos02.Recordset!doc_numero = rs_aux2!correl_doc
'                    'Txt_campo1.Caption = rs_aux2!correl_doc
'                    rs_aux2.Update
'                End If
'                ' GRABA Nombre de Archivo en ao_ventas_cabecera
'
'                'VAR_ARCH = RTrim(RTrim(Ado_datos02.Recordset!doc_codigo) + "-") + LTrim(Str(Ado_datos02.Recordset!doc_numero))
'                'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo = '" & VAR_ARCH & "' + '.PDF' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos02.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & Ado_datos02.Recordset("venta_codigo") & " "
'                'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo_cargado = 'N' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos02.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & Ado_datos02.Recordset("venta_codigo") & " "
'
'
'               'marca1 = Ado_datos02.Recordset.Bookmark
'               'Ado_datos02.Recordset.Requery
'        '       Ado_datos02.Refresh
'               'Ado_datos02.Recordset.Move marca1 - 1
'
'               '  Set rstacumdet = New ADODB.Recordset
'                '  If rstacumdet.State = 1 Then rstacumdet.Close
'                '  rstacumdet.Open "select sum(deuda_cobrada) as Cobrobs from ao_ventas_cobranzas where ges_gestion = '" & Ado_datos02.Recordset("ges_gestion") & "' and venta_codigo = " & Ado_datos02.Recordset("venta_codigo"), db, adOpenKeyset, adLockOptimistic
'                '
'                '  Set rstdestino = New ADODB.Recordset
'                '  If rstdestino.State = 1 Then rstdestino.Close
'                '  rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & gestion0 & "' and venta_codigo = " & nroventa, db, adOpenKeyset, adLockOptimistic
'                '  If rstdestino.RecordCount > 0 Then
'                '    rstdestino!deuda_cobrada = rstacumdet!Cobrobs
'                '    rstdestino!saldo_p_cobrar = (rstdestino!monto_total_Bs - rstdestino!monto_cobrado - rstdestino!deuda_cobrada)
'                '    rstdestino.Update
'                '  End If
'                '  If rstdestino.State = 1 Then rstdestino.Close
'                '  If rstacumdet.State = 1 Then rstacumdet.Close
'               VAR_SW = 2
'               Call Contabiliza_venta
'               Call OptFilGral1_Click
'           End If
'        End If
'     End If
' Else
'    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atenci�n!"
'  End If
End Sub

Private Sub BtnAprobar1_Click()

 'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
 If Ado_datos01.Recordset.RecordCount > 0 Then
    If Ado_datos01.Recordset!estado_codigo_fac = "APR" And Ado_datos01.Recordset!estado_codigo_fac1 = "REG" Then      'Ado_datos.Recordset("estado_codigo_anl") = "REG"
      sino = MsgBox("Esta seguro de generar la NOTA CREDITO - DEBITO ?", vbYesNo, "Confirmando")
      If sino = vbYes Then
          db.Execute "update ao_ventas_cobranza set estado_codigo_fac1 = 'APR' Where venta_codigo = " & Ado_datos01.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos01.Recordset!cobranza_codigo & "  "
          'GENERA NOTA CREDITO - DEBITO
          Set rs_datos12 = New ADODB.Recordset
          If rs_datos12.State = 1 Then rs_datos12.Close
          rs_datos12.Open "Select * from fo_nota_credito_debito where cobranza_codigo = " & Ado_datos01.Recordset!cobranza_codigo & " and cobranza_nro_factura = " & Ado_datos01.Recordset!cobranza_nro_factura & " ", db, adOpenKeyset, adLockOptimistic
          If rs_datos12.RecordCount > 0 Then
            MsgBox "NO se generar NOTA CREDITO - DEBITO, el registro que ya fue Procesado. ", , "Atencion"
          Else
            'wwwwwwwwwwwwwwwwwwwww
              ' hora_registro
            rs_datos12.AddNew
            rs_datos12!ges_gestion = glGestion
            rs_datos12!cobranza_codigo = Ado_datos01.Recordset!cobranza_codigo
            rs_datos12!venta_codigo = Ado_datos01.Recordset!venta_codigo
            rs_datos12!cobranza_nro_factura = Ado_datos01.Recordset!cobranza_nro_factura
            
            rs_datos12!nro_factura_dc = "0"
            rs_datos12!cobranza_prog_codigo = Ado_datos01.Recordset!cobranza_prog_codigo
            rs_datos12!debito_credito = "D"
            rs_datos12!beneficiario_codigo_fac = Ado_datos01.Recordset!beneficiario_codigo_fac
            
            rs_datos12!monto_dc_bs = Ado_datos01.Recordset!cobranza_total_bs * 0.13
            rs_datos12!monto_dc_dol = Ado_datos01.Recordset!cobranza_total_dol * 0.13
            
            rs_datos12!dc_fecha = Ado_datos01.Recordset!cobranza_fecha_fac      'Format(Date, "dd/mm/yyyy")
            rs_datos12!dc_fecha_fac = Ado_datos01.Recordset!cobranza_fecha_fac2
            rs_datos12!observaciones = Ado_datos01.Recordset!cobranza_observaciones
            'rs_datos12!dc_codigo_control = Ado_datos01.Recordset!cobranza_codigo_control
            rs_datos12!Literal = Ado_datos01.Recordset!Literal
        
            'rs_datos12!dc_nro_autorizacion = Ado_datos01.Recordset!cobranza_nro_autorizacion
            rs_datos12!correl_contab_dc = Ado_datos01.Recordset!correl_contab
            rs_datos12!estado_codigo = "REG"            'Ado_datos.Recordset!estado_codigo_anl
            rs_datos12!usr_codigo = glusuario           'Ado_datos.Recordset!usr_codigo_anl
            rs_datos12!fecha_registro = Ado_datos01.Recordset!fecha_registro
        
            rs_datos12!trans_codigo = Ado_datos01.Recordset!trans_codigo
            rs_datos12!cmpbte_deposito = Ado_datos01.Recordset!cmpbte_deposito
            rs_datos12!cta_codigo = Ado_datos01.Recordset!cta_codigo
            rs_datos12.Update
          End If
      End If
        '  rs_datos12!beneficiario_codigo_resp = dtc_codigo4A.Text                                                     'Codigo Cobrador
          'wwwwwwwwwwwwwwwwwwwww
          'marca1 = Ado_datos.Recordset.Bookmark
          'Call OptFilGral2_Click
          'Ado_datos.Recordset.Move marca1 - 1
    Else
      MsgBox "NO se puede ANULAR, porque el registro NO fue Facturado o ya fue Cobrado...", , "Atencion"
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atenci�n!"
  End If
End Sub

Private Sub BtnAprobar2_Click()
' If Ado_datos02.Recordset.RecordCount > 0 Then
'     COBR_BS = Ado_datos02.Recordset!cobranza_deuda_bs + Ado_datos02.Recordset!cobranza_deuda_bs2            'Monto Total Cobrado Bs
'     If IsNull(Ado_datos02.Recordset!cobranza_deuda_bs) Or (COBR_BS = 0) Then
'        MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
'        Exit Sub
'     Else
'        If COBR_BS < Ado_datos02.Recordset!cobranza_total_bs Then
'            'MsgBox "No se puede APROBAR, hasta que el Monto Cobrado sea igual al Monto Facturado. Vuelva a intentar ...", , "Atenci�n"
'            MsgBox "No se puede APROBAR hasta que el Total Monto Cobrado sea igual al Monto Facturado ...", , "Atenci�n"
'            Ado_datos02.Recordset!cobranza_fecha_cobro1 = DTPFechaCobro2.Value
'            Ado_datos02.Recordset!estado_codigo_bco1 = "APR"
'            Ado_datos02.Recordset!estado_codigo = "REG"
'            Ado_datos02.Recordset.Update
'            'Exit Sub
'        Else
'            If Ado_datos02.Recordset("estado_codigo_bco") = "REG" Then
'               sino = MsgBox("Esta seguro de Verificar la Cobranza ?", vbYesNo, "Confirmando")
'               If sino = vbYes Then
'                 If TxtDscto2.Text = "0.00" Or TxtDscto2.Text = "" Then
'                    Ado_datos02.Recordset!cobranza_fecha_cobro = DTPFechaCobro2.Value
'                 Else
'                    Ado_datos02.Recordset!cobranza_fecha_cobro = DTPFechaCobro02.Value
'                 End If
'                 Ado_datos02.Recordset!cobranza_fecha_cobro1 = DTPFechaCobro2.Value
'                 Ado_datos02.Recordset!estado_codigo_bco1 = "APR"
'                 Ado_datos02.Recordset!estado_codigo_bco = "APR"
'                 Ado_datos02.Recordset!estado_codigo = "REG"
'                 Ado_datos02.Recordset.Update
'                  'db.Execute "update ao_ventas_cobranza set estado_codigo_sol = 'APR' Where cobranza_codigo = " & Ado_datos01.Recordset("cobranza_codigo") & " "
'               End If
'            Else
'                MsgBox "No se puede APROBAR, el Registro ya fue Aprobado !! ", vbExclamation, "Atenci�n!"
'            End If
'        End If
'     End If
' Else
'    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atenci�n!"
' End If
End Sub

Private Sub BtnAprobar3_Click()
' If Ado_datos02.Recordset.RecordCount > 0 Then
'     COBR_BS = Ado_datos02.Recordset!cobranza_deuda_bs '+ Ado_datos02.Recordset!cobranza_deuda_bs2            'Monto Total Cobrado Bs
'     If IsNull(Ado_datos02.Recordset!cobranza_deuda_bs) Or (COBR_BS = 0) Then
'        MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
'        Exit Sub
'     Else
'        If COBR_BS <= Ado_datos02.Recordset!cobranza_total_bs Then
'            If Ado_datos02.Recordset("estado_codigo_bco1") = "REG" Then
'               sino = MsgBox("Esta seguro de Verificar la Cobranza 1 ?", vbYesNo, "Confirmando")
'               If sino = vbYes Then
'                  db.Execute "UPDATE ao_ventas_cobranza SET  "
'                    Ado_datos02.Recordset!cobranza_fecha_cobro = Date
'                    Ado_datos02.Recordset!estado_codigo_bco1 = "APR"
'                    Ado_datos02.Recordset!estado_codigo = "REG"
'                    Ado_datos02.Recordset.Update
'                  'db.Execute "update ao_ventas_cobranza set estado_codigo_sol = 'APR' Where cobranza_codigo = " & Ado_datos01.Recordset("cobranza_codigo") & " "
'               End If
'            Else
'                MsgBox "No se puede APROBAR, el Registro ya fue Aprobado !! ", vbExclamation, "Atenci�n!"
'            End If
'
'        Else
'            MsgBox "No se puede APROBAR, un Monto Cobrado Mayor al Monto Facturado. Vuelva a intentar ...", , "Atenci�n"
'            Exit Sub
'        End If
'     End If
' Else
'    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atenci�n!"
' End If
End Sub

Private Sub BtnBuscar_Click()
'JQA
 If Ado_datos.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      PosibleApliqueFiltro = False
      Dim rsNada As ADODB.Recordset
      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexi�n = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = dg_datos
      ClBuscaGrid.QueryUtilizado = queryinicial1
      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "No se puede Procesar el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
  End If
End Sub

Private Sub BtnBuscar1_Click()
'JQA
 If Ado_datos01.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      PosibleApliqueFiltro = False
      Dim rsNada As ADODB.Recordset
      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexi�n = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = dg_datos1
      ClBuscaGrid.QueryUtilizado = queryinicial
      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos01.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "No se puede Procesar el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
  End If

End Sub

Private Sub BtnBuscar2_Click()
 If Ado_datos02.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      PosibleApliqueFiltro = False
      Dim rsNada As ADODB.Recordset
      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexi�n = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = dg_datos2
      ClBuscaGrid.QueryUtilizado = queryinicial2
      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos02.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "No se puede Procesar el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
  End If

End Sub

Private Sub BtnCancelar_Click()
  'Ado_datos.Refresh
'  fraOpciones.Visible = True
  FraGrabarCancelar.Visible = False
  marca1 = Ado_datos.Recordset.Bookmark
  If (Ado_datos.Recordset!estado_codigo_sol = "APR" And Ado_datos.Recordset!estado_codigo_fac = "REG") Then
    Call OptFilGral1_Click
  Else
    Call OptFilGral2_Click
  End If
  FraNavega.Enabled = True
  FrmCobros.Enabled = False
  'Fra_datos.Enabled = True
  FrmDetalle.Enabled = True
  FrmCobranza.Visible = True
  'Fra_Total.Visible = True
  dg_datos.Visible = True
  FrmABMDet.Visible = True
  FrmABMDet2.Visible = True

  SSTab1.Tab = 1
  SSTab1.TabEnabled(0) = True
  SSTab1.TabEnabled(1) = True
  SSTab1.TabEnabled(2) = True
  'Ado_datos.Recordset.Move marca1 - 1
'  BtnImprimir2.Visible = True
  BtnImprimir3.Visible = True
  
  swnuevo = 0
   
End Sub

Private Sub BtnCancelar1_Click()

  FraGrabarCancelar1.Visible = False
  marca1 = Ado_datos01.Recordset.Bookmark
  If Ado_datos01.Recordset("estado_codigo_sol") = "REG" Then
    Call OptFilGral01_Click
  Else
    Call OptFilGral02_Click
  End If
  FraNavega1.Enabled = True
  FrmCobros1.Enabled = False
  FrmDetalle.Enabled = True
  FrmCobranza.Enabled = True
  
  FrmABMDet.Visible = True
  FrmABMDet2.Visible = True
    If glusuario = "RCUELA" Or glusuario = "ADMIN" Or glusuario = "CSALINAS" Then
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
    Else
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
    End If
  'Ado_datos01.Recordset.Move marca1 - 1
  swnuevo = 0

End Sub

Private Sub BtnCancelar2_Click()
  FraGrabarCancelar2.Visible = False
  marca1 = Ado_datos02.Recordset.Bookmark
  If (Ado_datos02.Recordset!estado_codigo_fac = "APR" And Ado_datos02.Recordset!estado_codigo_bco = "REG") Then
    Call OptFilGral03_Click
  Else
    Call OptFilGral04_Click
  End If
  FraNavega2.Enabled = True
  FrmCobros2.Enabled = False
  FrmDetalle.Enabled = True
  FrmCobranza.Enabled = True
  
  FrmABMDet.Visible = True
  FrmABMDet2.Visible = True
'  BtnAprobar3.Visible = True
    If glusuario = "RCUELA" Or glusuario = "ADMIN" Or glusuario = "CSALINAS" Then
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
    Else
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
    End If
  'Ado_datos01.Recordset.Move marca1 - 1
  swnuevo = 0

End Sub

Private Sub BtnCancelarBen_Click()
    frm_benef.Visible = False
    FraGrabarCancelar.Enabled = True
End Sub

Private Sub btnEliminar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo_fac = "APR" And Ado_datos.Recordset!estado_codigo_bco = "REG" Then      'Ado_datos.Recordset("estado_codigo_anl") = "REG"
      sino = MsgBox("Esta seguro de ANULAR la facturaci�n registrada ?", vbYesNo, "Confirmando")
      If sino = vbYes Then
        sino = MsgBox("Volver� a emitir otra FACTURA con este mismo registro ? (Si elige NO, se cierra el registro)", vbYesNo, "Confirmando")
        If sino = vbYes Then
          db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'REG' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set factura_impresa = 'N' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
        Else
          db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'ANL' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
        End If
          db.Execute "update ao_ventas_cobranza set cobranza_nro_factura_anl = '" & Ado_datos.Recordset!cobranza_nro_factura & "' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set cobranza_fecha_anl = '" & Format(Date, "dd/mm/yyyy") & "' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set usr_codigo_anl = '" & glusuario & "' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set estado_codigo_anl = 'APR' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set cobranza_fecha_ant = cobranza_fecha_fac Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set cobranza_codigo_control_anl = cobranza_codigo_control Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set correl_contab_anl = correl_contab Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion_anl = cobranza_nro_autorizacion Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          
          Set rs_datos12 = New ADODB.Recordset
          If rs_datos12.State = 1 Then rs_datos12.Close
          rs_datos12.Open "Select * from ao_ventas_cobro_anl where cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " and cobranza_nro_factura_anl = " & Ado_datos.Recordset!cobranza_nro_factura & " ", db, adOpenKeyset, adLockOptimistic
          If rs_datos12.RecordCount > 0 Then
            MsgBox "NO se puede ANULAR el registro que ya fue Aprobado o previamente Anulado.", , "Atencion"
          Else
            'wwwwwwwwwwwwwwwwwwwww
              ' hora_registro
            rs_datos12.AddNew
            rs_datos12!ges_gestion = glGestion
            rs_datos12!cobranza_codigo = Ado_datos.Recordset!cobranza_codigo
            rs_datos12!venta_codigo = Ado_datos.Recordset!venta_codigo
            
            rs_datos12!cobranza_nro_factura_anl = Ado_datos.Recordset!cobranza_nro_factura
            rs_datos12!cobranza_prog_codigo = Ado_datos.Recordset!cobranza_prog_codigo
            rs_datos12!beneficiario_codigo_fac = Ado_datos.Recordset!beneficiario_codigo_fac
            rs_datos12!cobranza_anuladal_bs = Ado_datos.Recordset!cobranza_total_bs
            rs_datos12!cobranza_anulada_dol = Ado_datos.Recordset!cobranza_total_dol
            
            rs_datos12!cobranza_fecha_anl = Ado_datos.Recordset!cobranza_fecha_fac      'Format(Date, "dd/mm/yyyy")
            rs_datos12!cobranza_fecha_fac2 = Ado_datos.Recordset!cobranza_fecha_fac2
            rs_datos12!cobranza_observaciones = Ado_datos.Recordset!cobranza_observaciones
            rs_datos12!cobranza_codigo_control_anl = Ado_datos.Recordset!cobranza_codigo_control
            rs_datos12!Literal = Ado_datos.Recordset!Literal
        
            rs_datos12!cobranza_nro_autorizacion_anl = Ado_datos.Recordset!cobranza_nro_autorizacion
            rs_datos12!correl_contab_anl = Ado_datos.Recordset!correl_contab
            rs_datos12!estado_codigo_anl = "APR"            'Ado_datos.Recordset!estado_codigo_anl
            rs_datos12!usr_codigo_anl = glusuario           'Ado_datos.Recordset!usr_codigo_anl
            rs_datos12!fecha_registro = Ado_datos.Recordset!fecha_registro
        
            rs_datos12!trans_codigo = Ado_datos.Recordset!trans_codigo
            rs_datos12!cmpbte_deposito = Ado_datos.Recordset!cmpbte_deposito
            rs_datos12!cta_codigo = Ado_datos.Recordset!cta_codigo
            rs_datos12.Update
          End If
      End If
        '  rs_datos12!beneficiario_codigo_resp = dtc_codigo4A.Text                                                     'Codigo Cobrador
          'wwwwwwwwwwwwwwwwwwwww
          'marca1 = Ado_datos.Recordset.Bookmark
          'Call OptFilGral2_Click
          'Ado_datos.Recordset.Move marca1 - 1
    Else
      MsgBox "NO se puede ANULAR, porque el registro NO fue Facturado o ya fue Cobrado...", , "Atencion"
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atenci�n!"
  End If
End Sub

Private Sub cambiarEtiquetaFactura()
    If lbl_fac.Caption <> "R-101" Then
       TxtCmpbte = False
       TxtCmpbte.backColor = &H80000005
       TxtCmpbte.ForeColor = &H80000008
       lbl_factura.Caption = "Nro.de Recibo"
    Else
       TxtCmpbte = True
       TxtCmpbte.backColor = &H404040
       TxtCmpbte.ForeColor = &HFFFFFF
       lbl_factura.Caption = "Nro.de Factura"
    End If
End Sub

Private Sub BtnGrabar_Click()
  Call cambiarEtiquetaFactura
  If dtc_codigo4A.Text = "" Then
    MsgBox "Debe Elejir " + Lbl_Cobrador.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atenci�n"
    Exit Sub
  End If
  If dtc_codigo5.Text = "" Then
    MsgBox "Debe Elejir <<Factura a Nombre de:>> !! Vuelva a Intentar ...", vbExclamation, "Atenci�n"
    Exit Sub
  End If
  If TxtMonto = "" Or TxtMonto = "0" Or TxtMonto = "0.00" Then
    MsgBox "Debe Registrar el " + lbl_monto.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atenci�n"
    Exit Sub
  End If
  If TxtObs = "" Then
    MsgBox "Debe Registrar el " + lbl_obs.Caption + " de la Cobranza, !! Vuelva a Intentar ...", vbExclamation, "Atenci�n"
    Exit Sub
  End If
  'If swnuevo = 2 Then
  'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
'  If DTPFechaProg.Visible = False Then
'    If TxtCmpbte = "" Or TxtCmpbte = "0" Then
'       MsgBox "Debe Registrar el " + lbl_factura.Caption + " a emitir al Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atenci�n"
'      Exit Sub
'    End If
'  End If
  'fin PARA COBRANZA WWWWWWWWWWWWWWWWWWW
  'valida = 1
  'If valida = 1 And dtc_codigo4A <> "" Then
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
    db.BeginTrans
    If swnuevo = 1 Then
'      rstdestino.Open "select * from ao_ventas_detalle where correl_venta = " & lblcorrelVenta & " and venta_codigo = " & TxtNroVenta, db, adOpenKeyset, adLockOptimistic
'      Set Ado_datos16.Recordset = rstdestino
'      Ado_datos16.Recordset.AddNew
      Ado_datos.Recordset!venta_codigo = Ado_datos.Recordset("venta_codigo")
      Ado_datos.Recordset!ges_gestion = glGestion       'Ado_datos.Recordset("ges_gestion")
      'Ado_datos.Recordset!cobranza_fecha_prog = DTPFechaProg                                'Fecha Programada a Cobrar
    End If
'      If Ado_datos.Recordset!beneficiario_codigo = "0" Then
'        Ado_datos.Recordset!beneficiario_codigo = dtc_codigo5.Text        'lbl_nit.Caption                                  'Codigo Beneficiario (Cliente)
'      End If
      Ado_datos.Recordset!beneficiario_codigo_fac = IIf(dtc_codigo5.Text = "", "0", dtc_codigo5.Text)       ' dtc_codigo5.Text  'dtc_codigo2A.Text                            'Beneficiario (Factura a nombre de ...)
      Ado_datos.Recordset!beneficiario_codigo_resp = dtc_codigo4A.Text                                                     'Codigo Cobrador
      Ado_datos.Recordset!trans_codigo = IIf(dtc_codigo6.Text = "", "O", dtc_codigo6.Text) 'tipo de Transaccion
      'Ado_datos.Recordset!nombre_cobrador = dtc_desc4A.Text   '+ " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
      Ado_datos.Recordset!cmpbte_deposito = IIf(Txt_deposito.Text = "", "0", Txt_deposito.Text)
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
      Ado_datos.Recordset!cta_codigo = IIf(dtc_cta.Text = "", "NN", dtc_cta.Text)
      Ado_datos.Recordset!cta_codigo2 = IIf(dtc_codigo7.Text = "", "NN", dtc_codigo7.Text)
      If TxtMonto.Text = "" Then
        Ado_datos.Recordset!cobranza_deuda_bs = "0"                                  'Monto Cobrado Bs.
        Ado_datos.Recordset!cobranza_deuda_dol = "0"        'Monto en Dolares
      Else
        Ado_datos.Recordset!cobranza_tdc = IIf(IsNull(Txt_tdc = ""), 6.96, CDbl(Txt_tdc.Text))                               'Monto Cobrado Bs.
'        Ado_datos.Recordset!cobranza_deuda_dol = CDbl(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
        Ado_datos.Recordset!cobranza_total_bs = CDbl(TxtMonto.Text)                                  'Monto Cobrado Bs.
        Ado_datos.Recordset!cobranza_total_dol = CDbl(TxtMontoDol)        'CDbl(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
      End If
      'VAR_GLOSA = Trim(Ado_datos02.Recordset!cobranza_observaciones) + " - Nro.: " + Trim(VAR_CITE)
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
      If Ado_datos.Recordset!cobranza_total_bs <> 0 Then
            Ado_datos.Recordset!Literal = Literal(CStr(Ado_datos.Recordset!cobranza_total_bs)) + " BOLIVIANOS"
      End If
      'Ado_datos.Recordset!cobranza_fecha_cobro = DTPFechaCobro.Value                                'Fecha de Cobranza
      'Call acumulaMont(Ado_datos.Recordset!ges_gestion, Ado_datos.Recordset!correl_venta, Ado_datos.Recordset!venta_codigo)
      Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"))
      '        '===== ini GENERA NRO. AUTORIZACION DE FACTURA ====
'        Set rs_aux1 = New ADODB.Recordset
'        rs_aux1.CursorLocation = adUseClient
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "select * from fc_Correl  where tipo_tramite = 'FAC_AUTORIZA'", db, adOpenDynamic, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
'          VAR_COD2 = CDbl(rs_aux1!numero_correlativo)
'          'rs_aux1!numero_correlativo = Trim(Str(VAR_COD2))
'          'rs_aux1.Update
'        End If
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        '===== fin TERMINA GENERACION NRO. AUTORIZACION DE FACTURA =====
'        'GENERA CORREL NOTA DEBITO POR DEPTO INI
'        Set rs_aux5 = New ADODB.Recordset
'        If rs_aux5.State = 1 Then rs_aux5.Close
'        'rs_aux5.Open "Select correl_contab as Codigo from gc_departamento where depto_codigo = '" & Left(VAR_PROY3, 1) & "'    ", db, adOpenStatic
'        rs_aux5.Open "Select * from fc_correl where tipo_tramite  = 'NDEBITO '    ", db, adOpenStatic
'        If Not rs_aux5.EOF Then
'            VAR_CONTAB = IIf(IsNull(rs_aux5!numero_correlativo), 1, CDbl(rs_aux5!numero_correlativo) + 1)
'        End If
'        'rs_aux5!Codigo = VAR_CONTAB
'        'rs_aux5.Update
'        db.Execute "update ao_ventas_cobranza set correl_contab = " & VAR_CONTAB & " Where ao_ventas_cobranza.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'        db.Execute "update fc_correl set numero_correlativo = " & VAR_CONTAB & " Where tipo_tramite = 'NDEBITO' "
'        'Ado_datos.Recordset!correl_contab = VAR_CONTAB
'        'GENERA CORREL NOTA DEBITO POR DEPTO FIN
'        If VAR_CONTAB < 10 Then
'            Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
'        End If
'        If VAR_CONTAB > 9 And VAR_CONTAB < 100 Then
'           Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
'        End If
'        If VAR_CONTAB > 99 And VAR_CONTAB < 1000 Then
'           Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
'        End If
      'Ado_datos.Recordset!proceso_codigo = "FIN"
      'Ado_datos.Recordset!subproceso_codigo = "FIN-01"
      Ado_datos.Recordset!etapa_codigo = "FIN-01-02"
      'Ado_datos.Recordset!clasif_codigo = "ADM"
      'Ado_datos.Recordset!doc_codigo = IIf(lbl_doc1 = "", "R-105", lbl_doc1)
      'Ado_datos.Recordset!doc_numero = IIf(lbl_docnro = "", "0", lbl_docnro)
      Ado_datos.Recordset!cmpbte_deposito = IIf(Txt_deposito = "", "0", Txt_deposito)
      If lbl_fac <> "R-101" Then
        Ado_datos01.Recordset!doc_codigo_fac = "R-103"
      Else
        Ado_datos01.Recordset!doc_codigo_fac = "R-101"
      End If
      If Ado_datos.Recordset!factura_impresa = "N" Then
         TxtCmpbte = "0"
         Ado_datos.Recordset!cobranza_nro_factura = IIf(TxtCmpbte = "", "0", Trim(TxtCmpbte))
      Else
         Ado_datos.Recordset!cobranza_nro_factura = IIf(TxtCmpbte = "", "0", Trim(TxtCmpbte))
      End If
      Ado_datos.Recordset!cobranza_nro_autorizacion = IIf(TxtAutorizacion = "", "0", Trim(TxtAutorizacion))
      Ado_datos.Recordset!poa_codigo = "3.1.2"
      Ado_datos.Recordset!cobranza_fecha_fac = DTPFechaCobro.Value         'Fecha de Facturacion
        'VAR_ANIO = CStr(glGestion)
        'VAR_MES = CStr(Month(Date))
        'VAR_DIA = CStr(Day(Date))
      Ado_datos.Recordset!cobranza_fecha_fac2 = ""        'VAR_ANIO & VAR_MES & VAR_DIA          'Fecha de Facturacion Texto
      Ado_datos.Recordset!estado_codigo_fac = "REG"
      Ado_datos.Recordset!usr_codigo = glusuario
      Ado_datos.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
      Ado_datos.Recordset!hora_registro = Format(Time, "hh:mm:ss")
      Ado_datos.Recordset.Update
    db.CommitTrans
    MsgBox "El registro se guardo correctamente"
    'Ado_datos.Recordset!doc_numero = Ado_datos.Recordset!cobranza_codigo       'Txt_cod_cobro.Text     ' "0"
  If swnuevo = 1 Then
    'Call abre_solicitud_lista
    'rc_Cobranza.Requery
    'Ado_datos.Refresh
    'Ado_datos.Recordset.MoveLast
  End If
    SSTab1.Tab = 1
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    FraNavega.Enabled = True
'    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FrmDetalle.Enabled = True
    FrmCobranza.Visible = True
    FrmCobros.Enabled = False
'    TxtCobrador.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
'    BtnImprimir2.Visible = True
    BtnImprimir3.Visible = True
    
    swnuevo = 0
    
  'Else
  '  MsgBox "Error en registro de datos, vuelva a intentar.!", vbCritical, ""
  'End If

End Sub

Private Sub BtnGrabar1_Click()
  If TxtMonto1 = "" Or TxtMonto1 = "0" Or TxtMonto1 = "0.00" Then
    MsgBox "Debe Registrar el " + lbl_monto1.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atenci�n"
    Exit Sub
  End If
    db.BeginTrans
      Ado_datos01.Recordset!cmpbte_deposito = IIf(Txt_deposito1.Text = "", "0", Txt_deposito1.Text)
      Ado_datos01.Recordset!cta_codigo = IIf(dtc_cta.Text = "", "NN", dtc_cta.Text)
      Ado_datos01.Recordset!cta_codigo2 = IIf(dtc_codigo7.Text = "", "NN", dtc_codigo7.Text)
      If TxtMonto1.Text = "" Then
        Ado_datos01.Recordset!cobranza_deuda_bs = "0"         'Monto Cobrado Bs.
        Ado_datos01.Recordset!cobranza_deuda_dol = "0"        'Monto en Dolares
      Else
        Ado_datos01.Recordset!cobranza_deuda_bs = CDbl(TxtMonto1.Text)                                  'Monto Cobrado Bs.
        Ado_datos01.Recordset!cobranza_deuda_dol = CDbl(TxtMonto1.Text) / GlTipoCambioMercado        'Monto en Dolares
      End If
      If TxtDscto1.Text = "" Then
        Ado_datos01.Recordset!cobranza_deuda_dol2 = "0"                              'Monto Cobrado Dolares
        Ado_datos01.Recordset!cobranza_deuda_bs2 = "0"         'Monto en Bs
      Else
        Ado_datos01.Recordset!cobranza_deuda_dol2 = "0"         ' IIf(TxtDscto1.Text = "", 0, CDbl(TxtDscto1.Text))                              'Monto Cobrado Dolares
        Ado_datos01.Recordset!cobranza_deuda_bs2 = "0"         'CDbl(TxtDscto1.Text) * GlTipoCambioMercado         'Monto en Bs
      End If
'      Ado_datos01.Recordset!cobranza_descuento_bs = "0"     'Ado_datos01.Recordset!cobranza_deuda_bs + Ado_datos01.Recordset!cobranza_deuda_bs2              'Monto Total Bs
'      Ado_datos01.Recordset!cobranza_descuento_dol = "0"     'Ado_datos01.Recordset!cobranza_deuda_dol + Ado_datos01.Recordset!cobranza_deuda_dol2           'Monto Total Dol
      Ado_datos01.Recordset!cobranza_solicitado_bs = Ado_datos01.Recordset!cobranza_programada_bs               'Monto Total Bs
      Ado_datos01.Recordset!cobranza_solicitado_dol = Ado_datos01.Recordset!cobranza_programada_dol           'Monto Total Dol
      Ado_datos01.Recordset!cobranza_total_bs = Ado_datos01.Recordset!cobranza_deuda_bs + Ado_datos01.Recordset!cobranza_deuda_bs2                  'Monto Total Bs
      Ado_datos01.Recordset!cobranza_total_dol = Ado_datos01.Recordset!cobranza_deuda_dol + Ado_datos01.Recordset!cobranza_deuda_dol2               'Monto Total Dol
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
      If Ado_datos01.Recordset!cobranza_total_bs <> 0 Then
            Ado_datos01.Recordset!Literal = Literal(CStr(Ado_datos01.Recordset!cobranza_total_bs)) + " BOLIVIANOS"
      End If
      Call acumulaMont(Ado_datos01.Recordset("ges_gestion"), Ado_datos01.Recordset("venta_codigo"))

      Ado_datos01.Recordset!proceso_codigo = "FIN"
      Ado_datos01.Recordset!subproceso_codigo = "FIN-01"
      Ado_datos01.Recordset!etapa_codigo = "FIN-01-03"
      Ado_datos01.Recordset!clasif_codigo = "ADM"
      Ado_datos01.Recordset!doc_codigo = IIf(lbl_doc01 = "", "R-105", lbl_doc01)
      Ado_datos01.Recordset!doc_numero = Txt_cod_cobro1.Caption          'IIf(lbl_docnro1 = "", "0", lbl_docnro1)
      If cmd_fac = "RECIBO" Then
        Ado_datos01.Recordset!doc_codigo_fac = "R-103"
      Else
        Ado_datos01.Recordset!doc_codigo_fac = "R-101"
      End If
      Ado_datos01.Recordset!cobranza_nro_factura = IIf(TxtCmpbte1 = "", "0", Trim(TxtCmpbt1))
      Ado_datos01.Recordset!cobranza_nro_autorizacion = IIf(TxtAutorizacion1 = "", "0", Trim(TxtAutorizacion1))
      Ado_datos01.Recordset!poa_codigo = "3.1.2"
      Ado_datos01.Recordset!cobranza_fecha_sol = DTPfechasol.Value         'Fecha de Cobranza
      Ado_datos01.Recordset!estado_codigo_sol = "REG"
      Ado_datos01.Recordset!usr_codigo_sol = glusuario
      Ado_datos01.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
      Ado_datos01.Recordset!hora_registro = Format(Time, "hh:mm:ss")
      Ado_datos01.Recordset.Update
    db.CommitTrans
    If glusuario = "RCUELA" Or glusuario = "ADMIN" Or glusuario = "MVALDIVIA" Or glusuario = "VSPAREDES" Or glusuario = "CSALINAS" Then
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
    Else
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
    End If
    FraNavega1.Enabled = True
'      fraOpciones1.Visible = True
      FraGrabarCancelar1.Visible = False
      FrmDetalle.Enabled = True
      FrmCobranza.Enabled = True
      FrmABMDet.Visible = True
      FrmABMDet2.Visible = True
      FrmCobros1.Enabled = False
    swnuevo = 0
  End Sub

Private Sub BtnGrabar2_Click()
    If (TxtMonto02 = "" Or TxtMonto02 = "0" Or TxtMonto02 = "0.00") Then
      'MsgBox "Debe Registrar el " + lbl_monto1.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atenci�n"
        MsgBox "Debe Registrar el Monto Cobrado Bs. o Cobrado USD, !! Vuelva a Intentar ...", vbExclamation, "Atenci�n"
        Exit Sub
    End If
    'RONALD
'    If (CDate(Ado_datos02.Recordset!cobranza_fecha_fac) > CDate(DTPFechaCobro2.Value)) Then
'        MsgBox "La <<Fecha Cobranza1>> No puede ser MENOR a la <<Fecha de Facturaci�n = " + CStr(Ado_datos02.Recordset!cobranza_fecha_fac) + ">>, Vuelva a Intentar !! ", vbExclamation, "Atenci�n!"
'        Exit Sub
'    End If
  
  
    db.BeginTrans
      Ado_datos02.Recordset!cmpbte_deposito = IIf(Txt_deposito2.Text = "", "0", Txt_deposito2.Text)
      Ado_datos02.Recordset!cmpbte_deposito2 = IIf(Txt_deposito3.Text = "", "0", Txt_deposito3.Text)
      Ado_datos02.Recordset!cta_codigo = IIf(dtc_cta2.Text = "", "NN", dtc_cta2.Text)
      Ado_datos02.Recordset!cta_codigo2 = IIf(dtc_codigo7_2.Text = "", "NN", dtc_codigo7_2.Text)
      If TxtMonto02.Text = "" Then
        Ado_datos02.Recordset!cobranza_deuda_bs = "0"         'Monto Cobrado Bs.
        Ado_datos02.Recordset!cobranza_deuda_dol = "0"        'Monto en Dolares
      Else
        Ado_datos02.Recordset!cobranza_deuda_bs = CDbl(TxtMonto02.Text)                               'Monto Cobrado Bs.
        Ado_datos02.Recordset!cobranza_deuda_dol = CDbl(TxtMonto02D.Text)        'CDbl(TxtMonto02.Text) / GlTipoCambioMercado        'Monto en Dolares
      End If
      If TxtDscto2.Text = "" Or TxtDscto2.Text = "0" Or TxtDscto2.Text = "0.00" Then
        Ado_datos02.Recordset!cobranza_deuda_dol2 = "0"                             'Segundo Monto Cobrado Dolares
        Ado_datos02.Recordset!cobranza_deuda_bs2 = "0"                              'Segundo Monto en Bs
      Else
        Ado_datos02.Recordset!cobranza_deuda_bs2 = "0"                              'IIf(TxtDscto2.Text = "", 0, CDbl(TxtDscto2.Text))   'Segundo Monto en Bs                           'Monto Cobrado Dolares
        Ado_datos02.Recordset!cobranza_deuda_dol2 = "0"                              ' CDbl(TxtDscto2.Text) / GlTipoCambioMercado         'Segundo Monto en Dolares
      End If
'        Ado_datos02.Recordset!cobranza_descuento_bs = Ado_datos02.Recordset!cobranza_deuda_bs + Ado_datos02.Recordset!cobranza_deuda_bs2              'Monto Total Bs
'        Ado_datos02.Recordset!cobranza_descuento_dol = Ado_datos02.Recordset!cobranza_deuda_dol + Ado_datos02.Recordset!cobranza_deuda_dol2               'Monto Total Dol
        
'      Ado_datos02.Recordset!cobranza_total_bs = Ado_datos02.Recordset!cobranza_deuda_bs + Ado_datos02.Recordset!cobranza_deuda_bs2              'Monto Total Bs
'      Ado_datos02.Recordset!cobranza_total_dol = Ado_datos02.Recordset!cobranza_deuda_dol + Ado_datos02.Recordset!cobranza_deuda_dol2               'Monto Total Dol
      
      Ado_datos02.Recordset!tipo_moneda = cmd_moneda1                                 'Tipo Moneda1
      Ado_datos02.Recordset!tipo_moneda2 = cmd_moneda2                                'Tipo Moneda2
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW RONALD
'      If (CDate(DTPFechaCobro2.Value) > CDate(DTPFechaCobro02.Value)) Then
'        Ado_datos02.Recordset!cobranza_fecha_cobro1 = Format(DTPFechaCobro2.Value, "dd/mm/yyyy")         'Fecha de Cobranza1
'        Ado_datos02.Recordset!cobranza_fecha_cobro = Format(DTPFechaCobro2.Value, "dd/mm/yyyy")         'Fecha de Cobranza2
'      Else
      'If (CDate(DTPFechaCobro2.Value) = "01/01/1900") Then
        Ado_datos02.Recordset!cobranza_fecha_cobro1 = IIf(IsNull(DTPFechaCobro2.Value) Or (CDate(DTPFechaCobro2.Value) = "01/01/1900"), Date, Format(DTPFechaCobro2.Value, "dd/mm/yyyy"))         'Fecha de Cobranza1
        Ado_datos02.Recordset!cobranza_fecha_cobro = IIf(IsNull(DTPFechaCobro02.Value) Or (CDate(DTPFechaCobro02.Value) = "01/01/1900"), Date, Format(DTPFechaCobro02.Value, "dd/mm/yyyy"))        'Fecha de Cobranza2
      'Else
      '  Ado_datos02.Recordset!cobranza_fecha_cobro1 = IIf(IsNull(DTPFechaCobro2.Value), Date, Format(DTPFechaCobro2.Value, "dd/mm/yyyy"))        'Fecha de Cobranza1
      '  Ado_datos02.Recordset!cobranza_fecha_cobro = IIf(IsNull(DTPFechaCobro02.Value), Date, Format(DTPFechaCobro02.Value, "dd/mm/yyyy"))        'Fecha de Cobranza2
      'End If
'      End If
      
      COBR_BS = Ado_datos02.Recordset!cobranza_deuda_bs + Ado_datos02.Recordset!cobranza_deuda_bs2            'Monto Total Cobrado Bs
      If COBR_BS > 0 Then
            Ado_datos02.Recordset!Literal = Literal(CStr(COBR_BS)) + " BOLIVIANOS"
      Else
            Ado_datos02.Recordset!Literal = "CERO 00/100 BOLIVIANOS"
      End If
      Call acumulaMont(Ado_datos02.Recordset("ges_gestion"), Ado_datos02.Recordset("venta_codigo"))
      
      'Ado_datos02.Recordset!proceso_codigo = "FIN"
      'Ado_datos02.Recordset!subproceso_codigo = "FIN-01"
      Ado_datos02.Recordset!etapa_codigo2 = "FIN-01-03"
      'Ado_datos02.Recordset!clasif_codigo = "ADM"
      Ado_datos02.Recordset!doc_codigo_cobr = "RE-426"        'IIf(lbl_doc01 = "", "R-105", lbl_doc01)
      Ado_datos02.Recordset!doc_numero = IIf(Txt_docnro = "", "0", Txt_docnro)
      'Ado_datos02.Recordset!doc_codigo_fac = "R-101"
      'Ado_datos02.Recordset!cobranza_nro_factura = IIf(TxtCmpbte2 = "", "0", Trim(TxtCmpbte2))
      'Ado_datos02.Recordset!cobranza_nro_autorizacion = IIf(TxtAutorizacion2 = "", "0", Trim(TxtAutorizacion2))
      'Ado_datos02.Recordset!poa_codigo = "3.1.2"
      Ado_datos02.Recordset!estado_codigo_bco = "REG"
      Ado_datos02.Recordset!usr_codigo_bco = glusuario
      Ado_datos02.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
      Ado_datos02.Recordset!hora_registro = Format(Time, "hh:mm:ss")
      Ado_datos02.Recordset.Update
    db.CommitTrans
    If glusuario = "RCUELA" Or glusuario = "ADMIN" Or glusuario = "CSALINAS" Then
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
    Else
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
    End If
    FraNavega2.Enabled = True
'      fraOpciones1.Visible = True
      FraGrabarCancelar2.Visible = False
      FrmDetalle.Enabled = True
      FrmCobranza.Enabled = True
      FrmABMDet.Visible = True
      FrmABMDet2.Visible = True
      FrmCobros2.Enabled = False
'      BtnAprobar3.Visible = False
    swnuevo = 0

End Sub

Private Sub BtnGrabarBen_Click()
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from gc_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' and beneficiario_codigo = '" & dtc_codigo8.Text & "'  ", db, adOpenStatic
    If rs_datos10.RecordCount = 0 Then
        'abrir gc_edificio_vs_beneficiario
        db.Execute "INSERT INTO gc_edificio_vs_beneficiario (edif_codigo, beneficiario_codigo, estado_codigo, fecha_registro, usr_codigo) VALUES ('" & VAR_PROY3 & "', '" & dtc_codigo8.Text & "', 'APR', '" & Date & "', '" & glusuario & "')"
        'Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
        Set rs_datos5 = New ADODB.Recordset
        If rs_datos5.State = 1 Then rs_datos5.Close
        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
        Set Ado_datos5.Recordset = rs_datos5
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        dtc_aux5.BoundText = dtc_codigo5.BoundText
        'FraGrabarCancelar.Enabled = True
'        lbl_nit_fac.Caption = dtc_codigo8.Text
'        lbl_benef_fac.Caption = dtc_desc8.Text

'        lbl_nit_fac.Visible = True
'        lbl_benef_fac.Visible = True
        
    Else
        MsgBox "Ya existe el Beneficiario relacionado, en: <<Facturado a Nombre de>>. Vuelva a intentar ...", , "Atenci�n"
    End If
    FraGrabarCancelar.Enabled = True
    frm_benef.Visible = False
End Sub

Private Sub BtnImprimir_Click()
    Select Case SSTab1.Tab
        Case 0
          If Ado_datos01.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza.rpt"
            CryR01.WindowShowRefreshBtn = True
            CryR01.StoredProcParam(0) = Me.Ado_datos01.Recordset!venta_codigo
            CryR01.StoredProcParam(1) = Me.Ado_datos01.Recordset!cobranza_codigo
            CryR01.Formulas(1) = "literalcobro = '" & Ado_datos01.Recordset!Literal & "' "
            CryR01.Formulas(2) = "correlcobro = '" & Ado_datos01.Recordset!cobranza_codigo & "' "
            '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
            iResult = CryR01.PrintReport
            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresi�n"
          Else
            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
          End If
        Case 1
          If Ado_datos.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza.rpt"
            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
            CryR01.WindowShowRefreshBtn = True
            CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            CryR01.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            CryR01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
            '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
            iResult = CryR01.PrintReport
            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresi�n"
          Else
            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
          End If
'        Case 2
'          If Ado_datos02.Recordset.RecordCount > 0 Then
''            Dim iResult As Variant  ', i%, y%
'            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza.rpt"
'            CryR01.WindowShowRefreshBtn = True
'            CryR01.StoredProcParam(0) = Me.Ado_datos02.Recordset!venta_codigo
'            CryR01.StoredProcParam(1) = Me.Ado_datos02.Recordset!cobranza_codigo
'            CryR01.Formulas(1) = "literalcobro = '" & Ado_datos02.Recordset!Literal & "' "
'            CryR01.Formulas(2) = "correlcobro = '" & Ado_datos02.Recordset!cobranza_codigo & "' "
'            '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'            iResult = CryR01.PrintReport
'            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresi�n"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
'          End If
    End Select

'  If Ado_datos.Recordset.RecordCount > 0 Then
'    Dim iResult As Variant, i%, y%
'    Dim co As New ADODB.Command
'
''    Dim rs As New ADODB.Recordset
''    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
''            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly
''    i = 1
''    y = 1
'    CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_nota_de_venta.rpt"
'    CryV01.WindowShowRefreshBtn = True
'    CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'    CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
'    CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
'    iResult = CryV01.PrintReport
'    If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresi�n"
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
'  End If
End Sub

Private Sub BtnImprimir2_Click()
    Call generarRepRecibo
End Sub

' Genera reporte de recibo
Private Sub generarRepRecibo()
    ' Verifica si codigo y numero son validos para recibo.
    'If Label38.Caption <> "" And TxtCmpbte2 <> "" And Label38.Caption <> "R-101" Then
    If Ado_datos.Recordset!doc_codigo_fac <> "R-101" Then
        Dim iResult As Integer
        Dim montoLiteral As String
        Dim Monto As Double
        Monto = 0
        If TxtMonto.Text = "0" Or TxtMonto.Text = "" Then
            MsgBox "No se puede emitir un Recibo con Monto cero, vuelva a intentar ..."
            Exit Sub
        Else
            Monto = TxtMonto.Text
        End If
        'If TxtCmpbte2 <> "" Then Monto = TxtCmpbte2
        If TxtDscto2.Text <> "" Then Monto = Monto + CInt(TxtDscto2.Text)
        montoLiteral = Monto
        montoLiteral = Literal(montoLiteral)
        crRecibo.WindowShowPrintSetupBtn = True
        crRecibo.WindowShowRefreshBtn = True
        crRecibo.ReportFileName = App.Path & "\Reportes\Ventas\ar_recibo_oficial.rpt"
        crRecibo.StoredProcParam(0) = Label38.Caption ' codigo
        crRecibo.StoredProcParam(1) = TxtCmpbte2 ' numero
        crRecibo.StoredProcParam(2) = "Bs. " + TxtMonto02 + "(" + montoLiteral + " BOLIVIANOS)" ' monto
        crRecibo.WindowState = crptMaximized
        iResult = crRecibo.PrintReport
        If iResult <> 0 Then
              MsgBox crRecibo.LastErrorNumber & " : " & crRecibo.LastErrorString, vbExclamation + vbOKOnly, "Error"
        End If
   Else
       MsgBox "No se puede generar reporte por falta de codigo y numero de recibo."
   End If
End Sub

'Private Sub BtnImprimir1_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    Dim iResult As Variant  ', i%, y%
'    CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'    CryR01.WindowShowRefreshBtn = True
''    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
''    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
''    CryR01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_codigo
'    CryR01.StoredProcParam(0) = Me.Ado_datos02.Recordset!venta_codigo
'    CryR01.StoredProcParam(1) = Me.Ado_datos02.Recordset!cobranza_codigo
'
'    CryR01.Formulas(1) = "literalcobro = '" & Ado_datos02.Recordset!Literal & "' "
'    CryR01.Formulas(2) = "correlcobro = '" & Ado_datos01.Recordset!cobranza_codigo & "' "
'    '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'    iResult = CryR01.PrintReport
'    If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresi�n"
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
'  End If
'
'End Sub

Private Sub BtnImprimir3_Click()
  If Ado_datos.Recordset.RecordCount > 0 And (dtc_aux5.Text <> "") Then
    If (Ado_datos.Recordset!factura_impresa = "N") And (Ado_datos.Recordset!cobranza_deuda_bs <> "0.00") Then
      If Ado_datos.Recordset!doc_codigo_fac = "R-101" Then
        '===== ini GENERA EL CODIGO DE FACTURA ====
        Set rs_aux1 = New ADODB.Recordset
        rs_aux1.CursorLocation = adUseClient
        If rs_aux1.State = 1 Then rs_aux1.Close
        rs_aux1.Open "select * from fc_dosificacion_docs where doc_codigo = 'R-101' AND estado_codigo = 'APR' ", db, adOpenDynamic, adLockOptimistic
        'rs_aux1.Open "select * from fc_dosificacion_docs  where doc_codigo = 'R-101'  ", db, adOpenDynamic, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            gestion0 = glGestion        'Ado_datos.Recordset("ges_gestion")
            correlv = Ado_datos.Recordset("venta_codigo")
            nroventa = Ado_datos.Recordset("venta_codigo")
            NRO_COBR = Me.Ado_datos.Recordset!cobranza_codigo
            VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
            VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
            'VAR_GLOSA = Ado_datos.Recordset!cobranza_observaciones
            VAR_GLOSA = Trim(Ado_datos.Recordset!cobranza_observaciones) + " - Tram.: " + Trim(VAR_CITE)
            'VAR_DOL2 = Round(Ado_datos.Recordset!cobranza_deuda_dol, 2)
            'VAR_BS2 = Round(Ado_datos.Recordset!cobranza_deuda_bs, 2)
            VAR_DOL2 = Round(Ado_datos.Recordset!cobranza_total_dol, 2)
            VAR_BS2 = Round(Ado_datos.Recordset!cobranza_total_bs, 2)
            'VAR_CTA = IIf(Ado_datos.Recordset!Cta_Codigo = "", "NN", Ado_datos.Recordset!Cta_Codigo)
            var_literal = Ado_datos.Recordset!Literal
            
            Llave = Trim(rs_aux1!dosifica_llave)
            If dtc_aux5.Text Like " " Then
                MsgBox "Error en el NIT del Cliente, Contactese con el Administrador y vuelva a intentar ...", , "Atenci�n"
                Exit Sub
            Else
                NitCi = IIf(dtc_aux5.Text = "", Ado_datos.Recordset!beneficiario_codigo_fac, dtc_aux5.Text)    'VAR_BENEF
            End If
            Autorizacion = rs_aux1!dosifica_autorizacion
            'Fecha = Val(Format((Date), "YYYYMMDD"))
            'Monto = Redondeo((VAR_BS2), 0)
            'CodigoContro = CodigoControl(Autorizacion, NroFactura, NitCi, Fecha, Monto, Llave)
            VAR_PROY2 = Ado_datos16.Recordset!edif_codigo
            VAR_COD4 = Ado_datos16.Recordset!unidad_codigo
            VAR_TIPOV = Ado_datos16.Recordset!venta_tipo
            VAR_SOL = Ado_datos16.Recordset!solicitud_codigo
            VAR_MONEDA = Ado_datos.Recordset!tipo_moneda
            'CodigoContro = CodigoControl(NroFactura)
            If Autorizacion <> "" And NitCi <> "" And Llave <> "" And VAR_BS2 <> "0" And rs_aux1!CORREL >= 0 Then
                VAR_SW = 1
            Else
                VAR_SW = 0
                MsgBox "Error en Autorizacion, NIT o Llave, Contactese con el Administrador y vuelva a intentar ...", , "Atenci�n"
                Exit Sub
            End If
            VAR_COD1 = CDbl(rs_aux1!CORREL) + 1
            sino = MsgBox("Esta seguro(a) de IMPRIMIR la Factura Nro. " + Str(VAR_COD1) + " ?", vbYesNo, "Confirmando")
            If sino = vbYes Then
                rs_aux1!CORREL = Trim(Str(VAR_COD1))
                rs_aux1.Update
                'GENERA CORREL NOTA DEBITO POR DEPTO INI
                Set rs_aux5 = New ADODB.Recordset
                If rs_aux5.State = 1 Then rs_aux5.Close
                'rs_aux5.Open "Select correl_contab as Codigo from gc_departamento where depto_codigo = '" & Left(VAR_PROY3, 1) & "'    ", db, adOpenStatic
                rs_aux5.Open "Select * from fc_correl where tipo_tramite  = 'NDEBITO '    ", db, adOpenStatic
                If Not rs_aux5.EOF Then
                    VAR_CONTAB = IIf(IsNull(rs_aux5!numero_correlativo), 1, CDbl(rs_aux5!numero_correlativo) + 1)
                End If
                'rs_aux5!Codigo = VAR_CONTAB
                'rs_aux5.Update
                
                VAR_COD2 = rs_aux1!dosifica_autorizacion
                NroFactura = Trim(Str(VAR_COD1))
                Fecha = Val(Format((Date), "YYYYMMDD"))
                Monto = Redondeo((VAR_BS2), 0)
                
                CodigoContro = CodigoControl(Autorizacion, NroFactura, NitCi, Fecha, Monto, Llave)
                If CodigoContro = "" Or CodigoContro = "0" Then
                    VAR_SW = 0
                    MsgBox "Error en Codigo de Control, Contactese con el Administrador o vuelva a intentar ...", , "Atenci�n"
                    Exit Sub
                Else
                    VAR_SW = 1
                End If
                db.Execute "update ao_ventas_cobranza set correl_contab = " & VAR_CONTAB & " Where ao_ventas_cobranza.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
                db.Execute "update fc_correl set numero_correlativo = " & VAR_CONTAB & " Where tipo_tramite = 'NDEBITO' "
                'Ado_datos.Recordset!correl_contab = VAR_CONTAB
                If VAR_CONTAB < 10 Then
                    'Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
                    VAR_GLOSA = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
                End If
                If VAR_CONTAB > 9 And VAR_CONTAB < 100 Then
                   'Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
                   VAR_GLOSA = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
                End If
                If VAR_CONTAB > 99 Then
'                    If VAR_CONTAB > 1200 Then
'                        MsgBox "El ND Finaliza en 6564 ... ", , "Atenci�n"
'                    End If
                   'Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
                   VAR_GLOSA = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
                End If
                db.Execute "update ao_ventas_cobranza set cobranza_observaciones = '" & VAR_GLOSA & "' Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'               'GENERA CORREL NOTA DEBITO POR DEPTO FIN
                
                '===== ini nombre archivo de la FACTURA ====
                'db.Execute "update ao_ventas_cobranza set archivo_foto = '" & doc_codigo & "' + '-' + '" & Str(VAR_COD1) & "' + '.JPG' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
                db.Execute "update ao_ventas_cobranza set archivo_foto = 'R101-' + '" & Str(VAR_COD1) & "' + '.JPG' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
                db.Execute "update ao_ventas_cobranza set archivo_foto_cargado = 'N' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
                '===== fin nombre archivo de la FACTURA ====
                ' ACTUALIZA NRO FAC. EN ao_ventas_cobranza
                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac = '" & Date & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
                db.Execute "update ao_ventas_cobranza set cobranza_nro_factura = " & VAR_COD1 & " Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
                db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion = " & VAR_COD2 & " Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
                'IMPRIMIR FACTURA
'                VAR_ANIO = CStr(glGestion)
'                VAR_MES = CStr(Month(Date))
'                VAR_DIA = CStr(Day(Date))
'                VAR_FECHA = VAR_ANIO & VAR_MES & VAR_DIA
                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac2 = '" & Fecha & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & NRO_COBR & " "
                'Dim F1
                'FI = Ado_datos.Recordset!cobranza_fecha_cobro
                'frm_qr.txt_texto = GlParametro + "|" + GlParametroDes + "|" + Trim(str(VAR_COD1)) + "|" + TrimSTR((VAR_COD2)) + "|" + '" & Ado_datos.Recordset!cobranza_fecha_cobro & "' + "|" + " & Ado_datos.Recordset!cobranza_deuda_bs & " + "|" + '" & rs_aux1!dosifica_codigo_control & "' + "|" + '" & rs_aux1!dosifica_fecha_limite & "' + "|" + "0" + "|" + "0" + "|" + '" & Ado_datos.Recordset!beneficiario_codigo & "' + "|" + '" & dtc_desc2A.Text & "'
                'frm_qr.Show vbModal
                'NIT del emisor, Nombre o Raz�n Social del emisor, N�mero correlativo de Factura, N�mero de Autorizaci�n, Fecha de emisi�n, Importe de la compra, C�digo de Control, Fecha L�mite de Emisi�n, 0, 0, NIT / NDI Comprador, Nombre o Raz�n Social del comprador
                
                'MsgBox "Se est� Imprimiendo la Factura Nro. " + Str(VAR_COD1), , "Atenci�n"
                db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
                db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'APR' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
                db.Execute "update ao_ventas_cobranza set estado_codigo_bco = 'REG' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
                
                db.Execute "update ao_ventas_cobranza set cobranza_codigo_control = '" & CodigoContro & "' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
                
'                Ado_datos.Recordset!estado_codigo_fac = "APR"
'                Ado_datos.Recordset.Update
                'INI QR
                'sFile = "C:\Tmp\QRCode.bmp"
                '1003579028
                '& "|" & Format(Trim("0"), "###0.00") _
                'dtc_aux5.Text
                sFile = App.Path & "\CLIENTES\QRCode.bmp"
                CadenaQ = Trim("1018533029") _
                & "|" & Trim(VAR_COD1) _
                & "|" & Trim(VAR_COD2) _
                & "|" & Format(Trim(Date), "DD/MM/YYYY") _
                & "|" & Format(Trim(VAR_BS2), "###0.00") _
                & "|" & Format(Trim(VAR_BS2), "###0.00") _
                & "|" & Trim(CodigoContro) _
                & "|" & Trim(dtc_aux5.Text) _
                & "|" & Trim("0") _
                & "|" & Trim("0") _
                & "|" & Trim("0") _
                & "|" & Trim("0")
                
                'CadenaQ = Trim(txtNitEmisor.Text) _
                '& "|" & Trim(txtNumeroFactura.Text) _
                '& "|" & Trim(txtNumeroAutorizacion.Text) _
                '& "|" & Format(Trim(txtFechaEmision.Text), "DD/MM/YYYY") _
                '& "|" & Format(Trim(txtImporteCompra.Text), "###0.00") _
                '& "|" & Format(Trim(txtFiscal.Text), "###0.00") _
                '& "|" & Trim(txtCodigoControl.Text) _
                '& "|" & Trim(txtNitComprador.Text) _
                '& "|" & Trim(txtImporteICE.Text) _
                '& "|" & Trim(txtGravadas.Text) _
                '& "|" & Trim(txtNoFiscal) _
                '& "|" & Trim(TxtDescuento)
'                MsgBox CadenaQ
'                FastQRCode CadenaQ, sFile
                Set Picture1.Picture = LoadPicture(sFile)
                'FIN QR
                'Call IMPRIME_FACTURA
                Call IMPRIME_QR
                'MsgBox CadenaQ
                If VAR_TIPOV = "C" Then
                    Call Contabiliza_venta
                End If
            Else
                VAR_COD1 = "0"
                If rs_aux1.State = 1 Then rs_aux1.Close
                Exit Sub
            End If
        End If
        If rs_aux1.State = 1 Then rs_aux1.Close
        '===== fin TERMINA GENERACION DE FACTURA =====
        

'        '===== ini GENERA NRO. AUTORIZACION DE FACTURA ====
'        Set rs_aux1 = New ADODB.Recordset
'        rs_aux1.CursorLocation = adUseClient
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "select * from fc_Correl  where tipo_tramite = 'FAC_AUTORIZA'", db, adOpenDynamic, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
'          VAR_COD2 = CDbl(rs_aux1!numero_correlativo)
'          'rs_aux1!numero_correlativo = Trim(Str(VAR_COD2))
'          'rs_aux1.Update
'        End If
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        '===== fin TERMINA GENERACION NRO. AUTORIZACION DE FACTURA =====
        
'        Dim iResult As Variant  ', i%, y%
'        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R-101_factura.rpt"
'        CryF01.WindowShowRefreshBtn = True
'        CryF01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'        CryF01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
'        CryF01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_codigo
'
'        CryF01.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
'        CryF01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
'        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'        iResult = CryF01.PrintReport
'        If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresi�n"
        
        TxtCmpbte = VAR_COD1
        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
          Call OptFilGral1_Click
        Else
          Call OptFilGral2_Click
        End If
      Else
        Call generarRepRecibo
      End If
      If Ado_datos.Recordset!doc_codigo_fac = "R-103" Then
      'WWWWWWWWWWWWWWWWWWWWWWWWW
        '===== ini GENERA EL CODIGO DE RECIBO ====
        Set rs_aux1 = New ADODB.Recordset
        rs_aux1.CursorLocation = adUseClient
        If rs_aux1.State = 1 Then rs_aux1.Close
        rs_aux1.Open "select * from gc_documentos_respaldo where doc_codigo = 'R-103' AND estado_codigo = 'APR' ", db, adOpenDynamic, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            gestion0 = glGestion        'Ado_datos.Recordset("ges_gestion")
            correlv = Ado_datos.Recordset("venta_codigo")
            nroventa = Ado_datos.Recordset("venta_codigo")
            NRO_COBR = Me.Ado_datos.Recordset!cobranza_codigo
            VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
            VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
            'VAR_GLOSA = Ado_datos.Recordset!cobranza_observaciones
            VAR_GLOSA = Trim(Ado_datos.Recordset!cobranza_observaciones) + " - Tram.: " + Trim(VAR_CITE)
            VAR_DOL2 = Round(Ado_datos.Recordset!cobranza_deuda_dol, 2)
            VAR_BS2 = Round(Ado_datos.Recordset!cobranza_deuda_bs, 2)
            'VAR_CTA = IIf(Ado_datos.Recordset!Cta_Codigo = "", "NN", Ado_datos.Recordset!Cta_Codigo)
            var_literal = Ado_datos.Recordset!Literal
            'Llave = Trim(rs_aux1!dosifica_llave)
            NitCi = IIf(dtc_aux5.Text = "", Ado_datos.Recordset!beneficiario_codigo_fac, dtc_aux5.Text)    'VAR_BENEF
            'Autorizacion = rs_aux1!dosifica_autorizacion
            VAR_PROY2 = Ado_datos16.Recordset!edif_codigo
            VAR_COD4 = Ado_datos16.Recordset!unidad_codigo
            VAR_TIPOV = Ado_datos16.Recordset!venta_tipo
            VAR_SOL = Ado_datos16.Recordset!solicitud_codigo
            VAR_MONEDA = Ado_datos.Recordset!tipo_moneda
        
            VAR_COD1 = CDbl(rs_aux1!correl_doc) + 1
            sino = MsgBox("Esta seguro(a) de IMPRIMIR la Recibo Nro. " + Str(VAR_COD1) + " ?", vbYesNo, "Confirmando")
            If sino = vbYes Then
                rs_aux1!correl_doc = Trim(Str(VAR_COD1))
                rs_aux1.Update
                'GENERA CORREL NOTA DEBITO POR DEPTO INI
'                Set rs_aux5 = New ADODB.Recordset
'                If rs_aux5.State = 1 Then rs_aux5.Close
'                rs_aux5.Open "Select * from fc_correl where tipo_tramite  = 'NDEBITO '    ", db, adOpenStatic
'                If Not rs_aux5.EOF Then
'                    VAR_CONTAB = IIf(IsNull(rs_aux5!numero_correlativo), 1, CDbl(rs_aux5!numero_correlativo) + 1)
'                End If
'                db.Execute "update ao_ventas_cobranza set correl_contab = " & VAR_CONTAB & " Where ao_ventas_cobranza.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'                db.Execute "update fc_correl set numero_correlativo = " & VAR_CONTAB & " Where tipo_tramite = 'NDEBITO' "
'                If VAR_CONTAB < 10 Then
'                    VAR_GLOSA = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
'                End If
'                If VAR_CONTAB > 9 And VAR_CONTAB < 100 Then
'                   VAR_GLOSA = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
'                End If
'                If VAR_CONTAB > 99 And VAR_CONTAB < 6564 Then
'                    If VAR_CONTAB > 1200 Then
'                        MsgBox "El ND Finaliza en 6564 ... ", , "Atenci�n"
'                    End If
'                   VAR_GLOSA = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
'                End If
                VAR_GLOSA = TxtObs.Text
                db.Execute "update ao_ventas_cobranza set cobranza_observaciones = '" & VAR_GLOSA & "' Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                'GENERA CORREL NOTA DEBITO POR DEPTO FIN
                
                VAR_COD2 = "0"  'rs_aux1!dosifica_autorizacion
                NroFactura = Trim(Str(VAR_COD1))
                '===== ini nombre archivo de la FACTURA ====
                db.Execute "update ao_ventas_cobranza set archivo_foto = 'R103-' + '" & Str(VAR_COD1) & "' + '.JPG' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
                db.Execute "update ao_ventas_cobranza set archivo_foto_cargado = 'N' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
                '===== fin nombre archivo de la FACTURA ====
                ' ACTUALIZA NRO FAC. EN ao_ventas_cobranza
                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac = '" & Date & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
                db.Execute "update ao_ventas_cobranza set cobranza_nro_factura = " & VAR_COD1 & " Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
                db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion = " & VAR_COD2 & " Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
                'IMPRIMIR FACTURA
                Fecha = Val(Format((Date), "YYYYMMDD"))
                Monto = Redondeo((VAR_BS2), 0)
                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac2 = '" & Fecha & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & NRO_COBR & " "
                'Dim F1
                'FI = Ado_datos.Recordset!cobranza_fecha_cobro
                'frm_qr.txt_texto = GlParametro + "|" + GlParametroDes + "|" + Trim(str(VAR_COD1)) + "|" + TrimSTR((VAR_COD2)) + "|" + '" & Ado_datos.Recordset!cobranza_fecha_cobro & "' + "|" + " & Ado_datos.Recordset!cobranza_deuda_bs & " + "|" + '" & rs_aux1!dosifica_codigo_control & "' + "|" + '" & rs_aux1!dosifica_fecha_limite & "' + "|" + "0" + "|" + "0" + "|" + '" & Ado_datos.Recordset!beneficiario_codigo & "' + "|" + '" & dtc_desc2A.Text & "'
                'frm_qr.Show vbModal
                'NIT del emisor, Nombre o Raz�n Social del emisor, N�mero correlativo de Factura, N�mero de Autorizaci�n, Fecha de emisi�n, Importe de la compra, C�digo de Control, Fecha L�mite de Emisi�n, 0, 0, NIT / NDI Comprador, Nombre o Raz�n Social del comprador
                
                'MsgBox "Se est� Imprimiendo la Factura Nro. " + Str(VAR_COD1), , "Atenci�n"
                db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
                db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'APR' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
                db.Execute "update ao_ventas_cobranza set estado_codigo_bco = 'REG' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
                
                VAR_SW = 1
                'CodigoContro = CodigoControl(Autorizacion, NroFactura, NitCi, Fecha, Monto, Llave)
                'db.Execute "update ao_ventas_cobranza set cobranza_codigo_control = '" & CodigoContro & "' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
                Call IMPRIME_RECIBO
                If VAR_TIPOV = "C" Then
                    Call Contabiliza_venta
                End If
            Else
                VAR_COD1 = "0"
                If rs_aux1.State = 1 Then rs_aux1.Close
                Exit Sub
            End If
        End If
        If rs_aux1.State = 1 Then rs_aux1.Close
        '===== fin TERMINA GENERACION DE FACTURA =====
        TxtCmpbte = VAR_COD1
        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
          Call OptFilGral1_Click
        Else
          Call OptFilGral2_Click
        End If
      'WWWWWWWWWWWWWWWWWWWWWWWWW
      End If
    Else
        MsgBox "La Factura Nro. " + Ado_datos.Recordset!cobranza_nro_factura + " ya fue Impresa", , "Atenci�n"
        'Call IMPRIME_FACTURA
        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
          Call OptFilGral1_Click
        Else
          Call OptFilGral2_Click
        End If
    End If
  Else
    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
  End If
End Sub

Private Sub generar(Autorizacion As String, Numero As String, NitCi As String, Fecha As String, Monto As String, Llave As String)
' paso 1
'    Dim suma As String
'    Dim digitos As String
'    Dim digitossum(4) As Integer
'    Dim cadenas(4) As String
'    Dim inicio As Integer
'    Dim x As Integer
'
'    Dim arc4 As String
'    Dim suma_total As Long
'    Dim sumas(4) As Long
'    Dim strlen_arc4 As Integer
'    Dim i As Integer
'    Dim total As Long
'
'    Dim mensaje As String
'    Dim last As String
'
'        numero = verhoeff_add_recursive(numero, 2)
'        nitci = verhoeff_add_recursive(nitci, 2)
'        fecha = verhoeff_add_recursive(fecha, 2)
'        monto = verhoeff_add_recursive(monto, 2)
''            Dim suma As String = CType((Long.Parse(numero) _
''                        + (Long.Parse(nitci) _
''                        + (Long.Parse(fecha) + Long.Parse(monto)))),Long).ToString
'        suma = (CStr(numero) + (CStr(nitci) + (Trim(fecha) + CStr(monto))))
'        suma = verhoeff_add_recursive(suma, 5)
'' paso2
''            Dim digitos As String = ("" + suma.Substring((suma.Length - 5), 5))
''            Dim digitossum() As Integer = New Integer() {0, 0, 0, 0, 0}
''            Dim cadenas() As String = New String() {"", "", "", "", ""}
''            Dim inicio As Integer = 0
''            Dim x As Integer = 0
'    digitos = ("" + suma.Substring((suma.Length - 5), 5))
'    digitossum(0) = 0
'    digitossum(1) = 0
'    digitossum(2) = 0
'    digitossum(3) = 0
'    digitossum(4) = 0
'    cadenas(0) = ""
'    cadenas(1) = ""
'    cadenas(2) = ""
'    cadenas(3) = ""
'    cadenas(4) = ""
'    inicio = 0
'    x = 0
''    For Each d As Char In digitos.ToCharArray
''                digitossum(x) = (Integer.Parse(d.ToString) + 1)
''                cadenas(x) = llave.Substring(inicio, (Integer.Parse(d.ToString) + 1))
''                inicio = (inicio _
''                            + (Integer.Parse(d.ToString) + 1))
''                x = (x + 1)
''    Next
'    For x = 0 To Len(digitos)
'        digitossum(x) = (CInt(digitos) + 1)
'        cadenas(x) = llave.Substring(inicio, (CInt(digitos) + 1))
'        inicio = (inicio + (CInt(digitos) + 1))
'        x = (x + 1)
'    Next x
'            autorizacion = (autorizacion + cadenas(0))
'            numero = (numero + cadenas(1))
'            nitci = (nitci + cadenas(2))
'            fecha = (fecha + cadenas(3))
'            monto = (monto + cadenas(4))
'' paso3
'    arc4 = allegedrc4((autorizacion + (numero + (nitci + (fecha + monto)))), (llave + digitos))
'' paso4
'    suma_total = 0
'    sumas(0) = 0
'    sumas(1) = 0
'    sumas(2) = 0
'    sumas(3) = 0
'    sumas(4) = 0
'    strlen_arc4 = Len(arc4)
'    i = 0
'    Do While (i < strlen_arc4)
'                x = CInt(arc4(i))
'                sumas((i Mod 5)) = (sumas((i Mod 5)) + x)
'                suma_total = (suma_total + x)
'                i = (i + 1)
'    Loop
'' paso5
'    total = 0
'    i = 0
'    Do While (i < Len(sumas))
'                total = (total + (suma_total * (sumas(i) / digitossum(i))))
'                i = (i + 1)
'    Loop
'    mensaje = big_base_convert(total, 64)
'    last = allegedrc4(mensaje, (llave + digitos)).Insert(2, "-").Insert(5, "-").Insert(8, "-")
'            If (last.Length > 11) Then
'                last = last.Insert(11, "-")
'            End If
'    'Return last

End Sub

Private Sub big_base_convert(ByVal Numero As Long, ByVal baseconv As Long)
'    Dim dic(63) As Char
'    Dim cociente As Long
'    Dim resto As Long
'    Dim palabra As String
'
'    dic(0) = Microsoft.VisualBasic.ChrW(48)
'    dic(1) = Microsoft.VisualBasic.ChrW(49)
'    dic(2) = Microsoft.VisualBasic.ChrW(50)
'    dic(3) = Microsoft.VisualBasic.ChrW(51)
'    dic(4) = Microsoft.VisualBasic.ChrW(52)
'    dic(5) = Microsoft.VisualBasic.ChrW(53)
'    dic(6) = Microsoft.VisualBasic.ChrW(54)
'    dic(7) = Microsoft.VisualBasic.ChrW(55)
'    dic(8) = Microsoft.VisualBasic.ChrW(56)
'    dic(9) = Microsoft.VisualBasic.ChrW(57)
'    dic(10) = Microsoft.VisualBasic.ChrW(65)
'    dic(11) = Microsoft.VisualBasic.ChrW(66)
'    dic(12) = Microsoft.VisualBasic.ChrW(67)
'    dic(13) = Microsoft.VisualBasic.ChrW(68)
'    dic(14) = Microsoft.VisualBasic.ChrW(69)
'    dic(15) = Microsoft.VisualBasic.ChrW(70)
'    dic(16) = Microsoft.VisualBasic.ChrW(71)
'    dic(17) = Microsoft.VisualBasic.ChrW(72)
'    dic(18) = Microsoft.VisualBasic.ChrW(73)
'    dic(19) = Microsoft.VisualBasic.ChrW(74)
'    dic(20) = Microsoft.VisualBasic.ChrW(75)
'    dic(21) = Microsoft.VisualBasic.ChrW(76)
'    dic(22) = Microsoft.VisualBasic.ChrW(77)
'    dic(23) = Microsoft.VisualBasic.ChrW(78)
'    dic(24) = Microsoft.VisualBasic.ChrW(79)
'    dic(25) = Microsoft.VisualBasic.ChrW(80)
'    dic(26) = Microsoft.VisualBasic.ChrW(81)
'    dic(27) = Microsoft.VisualBasic.ChrW(82)
'    dic(28) = Microsoft.VisualBasic.ChrW(83)
'    dic(29) = Microsoft.VisualBasic.ChrW(84)
'    dic(30) = Microsoft.VisualBasic.ChrW(85)
'    dic(31) = Microsoft.VisualBasic.ChrW(86)
'    dic(32) = Microsoft.VisualBasic.ChrW(87)
'    dic(33) = Microsoft.VisualBasic.ChrW(88)
'    dic(34) = Microsoft.VisualBasic.ChrW(89)
'    dic(35) = Microsoft.VisualBasic.ChrW(90)
'    dic(36) = Microsoft.VisualBasic.ChrW(97)
'    dic(37) = Microsoft.VisualBasic.ChrW(98)
'    dic(38) = Microsoft.VisualBasic.ChrW(99)
'    dic(39) = Microsoft.VisualBasic.ChrW(100)
'    dic(40) = Microsoft.VisualBasic.ChrW(101)
'    dic(41) = Microsoft.VisualBasic.ChrW(102)
'    dic(42) = Microsoft.VisualBasic.ChrW(103)
'    dic(43) = Microsoft.VisualBasic.ChrW(104)
'    dic(44) = Microsoft.VisualBasic.ChrW(105)
'    dic(45) = Microsoft.VisualBasic.ChrW(106)
'    dic(46) = Microsoft.VisualBasic.ChrW(107)
'    dic(47) = Microsoft.VisualBasic.ChrW(108)
'    dic(48) = Microsoft.VisualBasic.ChrW(109)
'    dic(49) = Microsoft.VisualBasic.ChrW(110)
'    dic(50) = Microsoft.VisualBasic.ChrW(111)
'    dic(51) = Microsoft.VisualBasic.ChrW(112)
'    dic(52) = Microsoft.VisualBasic.ChrW(113)
'    dic(53) = Microsoft.VisualBasic.ChrW(114)
'    dic(54) = Microsoft.VisualBasic.ChrW(115)
'    dic(55) = Microsoft.VisualBasic.ChrW(116)
'    dic(56) = Microsoft.VisualBasic.ChrW(117)
'    dic(57) = Microsoft.VisualBasic.ChrW(118)
'    dic(58) = Microsoft.VisualBasic.ChrW(119)
'    dic(59) = Microsoft.VisualBasic.ChrW(120)
'    dic(60) = Microsoft.VisualBasic.ChrW(121)
'    dic(61) = Microsoft.VisualBasic.ChrW(122)
'    dic(62) = Microsoft.VisualBasic.ChrW(43)
'    dic(63) = Microsoft.VisualBasic.ChrW(47)
'
'    cociente = 1
'    resto = 0
'    palabra = ""
'    While (cociente > 0)
'                cociente = (numero / baseconv)
'                resto = (numero Mod baseconv)
'                palabra = (dic(resto) + palabra)
'                numero = cociente
'
'    End
'    '        Return palabra
End Sub
        
Private Sub SWAP(ByRef num1 As Integer, ByRef num2 As Integer)
    Dim temp As Integer
    temp = num2
    num2 = num1
    num1 = temp
End Sub
        
'Private Sub allegedrc4(mensaje As String, llaverc4 As String)
'            Dim state() As Integer = New Integer((256) - 1) {}
'            Dim x As Integer = 0
'            Dim y As Integer = 0
'            Dim index1 As Integer = 0
'            Dim index2 As Integer = 0
'            Dim nmen As Integer = 0
'            Dim i As Integer = 0
'            Dim cifrado As String = ""
'            i = 0
'            Do While (i < 256)
'                state(i) = i
'                i = (i + 1)
'            Loop
'            Dim strlen_llave As Integer = llaverc4.Length
'            Dim strlen_mensaje As Integer = mensaje.Length
'            i = 0
'            Do While (i < 256)
'                index2 = ((CType(llaverc4(index1),Integer) _
'                            + (state(i) + index2)) _
'                            Mod 256)
'                swap(state(index2), state(i))
'                index1 = ((index1 + 1) _
'                            Mod strlen_llave)
'                i = (i + 1)
'            Loop
'            Dim cadtemp As String = ""
'            i = 0
'            Do While (i < strlen_mensaje)
'                x = ((x + 1) _
'                            Mod 256)
'                y = ((state(x) + y) _
'                            Mod 256)
'                swap(state(y), state(x))
'                ' ^ = XOR function
'                nmen = (CType(mensaje(i),Integer) Or state(((state(x) + state(y)) _
'                            Mod 256)))
'                'The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
'                cadtemp = ("0" + big_base_convert(nmen, 16))
'                cifrado = (cifrado + cadtemp.Substring((cadtemp.Length - 2), 2))
'                i = (i + 1)
'            Loop
'            Return cifrado
'End Sub
'
'Private Shared Function calcsum(ByVal number As String) As Integer
'            Dim c As Integer = 0
'            Dim n As String = reverse(number)
'            Dim len As Integer = n.Length
'            Dim nchar() As Char = n.ToCharArray
'            Dim i As Integer = 0
'            Do While (i < len)
'                c = table_d(c, table_p(((i + 1) _
'                            Mod 8), Integer.Parse(nchar(i).ToString)))
'                i = (i + 1)
'            Loop
'            Return table_inv(c)
'End Sub
'
'Private Shared Function verhoeff_add_recursive(ByVal number As String, ByVal digits As Integer) As String
'            Dim temp As String = number
'
'            While (digits > 0)
'                temp = (temp + calcsum(temp))
'                digits = (digits - 1)
'
'            End While
'            Return temp
'End Sub
'
'Private Shared Function reverse(ByVal cadena As String) As String
'            Dim str() As Char = cadena.ToCharArray
'            Array.Reverse(str)
'            Return New String(str)
'End Sub

Private Sub IMPRIME_FACTURA()
        'IMPRIMIR FACTURA
    Dim iResult As Variant  ', i%, y%
    sino = MsgBox("Imprimir� con el detalle de Bienes ? ", vbYesNo, "Confirmando")
    If sino = vbYes Then
        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_anterior_rep.rpt"
    Else
        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_anterior.rpt"
    End If
        CryF01.WindowShowRefreshBtn = True
        CryF01.StoredProcParam(0) = glGestion       'Me.Ado_datos.Recordset!ges_gestion
        CryF01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
        CryF01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
        'var_literal = "-"   'Ado_datos.Recordset!Literal
        CryF01.Formulas(1) = "literalcobro = '" & var_literal & "' "
        CryF01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
        iResult = CryF01.PrintReport
        If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresi�n"

End Sub

Private Sub IMPRIME_QR()
'    'IMPRIMIR FACTURA con QR
'    'Dim Exel As Object
'    'Set Exel = CreateObject("Excel.Application")
'    'Exel.Workbooks.Open "c:\tmp\Factura.xlt", , , , "123", "123"
'    'Exel.Visible = True
'    Call CmdFoto_Click
'    ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"
'
'    Picture2.AutoRedraw = True
'    Picture2.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
'
'    ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"
'    ' MsgBox CadenaQr
'    FastQRCode CadenaQr, ImagenQr
'    Picture1.AutoRedraw = True
'    Picture1.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
'    Clipboard.Clear
'    Clipboard.SetData Picture2.Image
''    Exel.Application.Range("a2").Select
''    Exel.Application.ActiveSheet.Paste
'
'    Dim iResult As Variant  ', i%, y%
'    sino = MsgBox("Imprimir� con el detalle de Bienes ? ", vbYesNo, "Confirmando")
'    If sino = vbYes Then
'        CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_rep.rpt"
'    Else
'        CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura.rpt"
'    End If
'        CryQ01.WindowShowRefreshBtn = True
'        CryQ01.StoredProcParam(0) = glGestion       'Me.Ado_datos.Recordset!ges_gestion
'        CryQ01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
'        CryQ01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
'        'var_literal = "-"   'Ado_datos.Recordset!Literal
'        CryQ01.Formulas(1) = "literalcobro = '" & var_literal & "' "
'        CryQ01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
'        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
'        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'        iResult = CryQ01.PrintReport
'        If iResult <> 0 Then MsgBox CryQ01.LastErrorNumber & " : " & CryQ01.LastErrorString, vbCritical, "Error de impresi�n"

End Sub

Private Sub IMPRIME_RECIBO()
        'IMPRIMIR FACTURA
        Dim iResult As Variant  ', i%, y%
        'CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R-101_factura.rpt"
        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_oficial.rpt"
        CryF01.WindowShowRefreshBtn = True
        CryF01.StoredProcParam(0) = glGestion       'Me.Ado_datos.Recordset!ges_gestion
        CryF01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
        CryF01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
        'var_literal = "-"   'Ado_datos.Recordset!Literal
        CryF01.Formulas(1) = "literalcobro = '" & var_literal & "' "
        CryF01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
        iResult = CryF01.PrintReport
        If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresi�n"

End Sub
Private Sub BtnImprimir4_Click()
    Select Case SSTab1.Tab
        Case 0
            If Ado_datos16.Recordset.RecordCount > 0 Then
              'CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_R105_kardex.rpt"
              CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
              CryV01.WindowShowRefreshBtn = True
              CryV01.StoredProcParam(0) = Me.Ado_datos01.Recordset!ges_gestion            'glGestion
              CryV01.StoredProcParam(1) = Me.Ado_datos01.Recordset!venta_codigo           'nroventa        '
              CryV01.StoredProcParam(2) = Me.Ado_datos01.Recordset!cobranza_prog_codigo   'NRO_COBR        '
              'Literal por el Total de la Compra
              var_literal = Literal(CStr(Ado_datos16.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
              CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
              CryV01.Formulas(2) = "correlcobro = '" & Ado_datos01.Recordset!cobranza_prog_codigo & "' "
              iResult = CryV01.PrintReport
              If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresi�n"
            Else
              MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
            End If
        Case 1
            If Ado_datos16.Recordset.RecordCount > 0 Then
              'CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_R105_kardex.rpt"
              CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
              CryV01.WindowShowRefreshBtn = True
              CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion            'glGestion
              CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo           'nroventa        '
              CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_prog_codigo   'NRO_COBR        '
              'Literal por el Total de la Compra
              var_literal = Literal(CStr(Ado_datos16.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
              CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
              'CryV01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
              CryV01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_prog_codigo & "' "
              '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
              iResult = CryV01.PrintReport
              If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresi�n"
            Else
              MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
            End If
'        Case 2  'Ado_datos02
'            If Ado_datos16.Recordset.RecordCount > 0 Then
'              'CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_R105_kardex.rpt"
'              CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
'              CryV01.WindowShowRefreshBtn = True
'              CryV01.StoredProcParam(0) = Me.Ado_datos02.Recordset!ges_gestion            'glGestion
'              CryV01.StoredProcParam(1) = Me.Ado_datos02.Recordset!venta_codigo           'nroventa        '
'              CryV01.StoredProcParam(2) = Me.Ado_datos02.Recordset!cobranza_prog_codigo   'NRO_COBR        '
'              'Literal por el Total de la Compra
'              var_literal = Literal(CStr(Ado_datos16.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
'              CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
'              'CryV01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
'              CryV01.Formulas(2) = "correlcobro = '" & Ado_datos02.Recordset!cobranza_prog_codigo & "' "
'              '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'              iResult = CryV01.PrintReport
'              If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresi�n"
'            Else
'              MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
'            End If
    End Select
  
End Sub

Private Sub BtnModificar_Click()
  codigo_doc = lbl_fac.Caption
  If Ado_datos.Recordset.RecordCount > 0 Then
    If codigo_doc <> "R-101" Then
         Call cambiarEtiquetaFactura
         Dim Cmd1 As ADODB.Command
         Dim rs  As ADODB.Recordset
         Set Cmd1 = New ADODB.Command
         Set rs = New ADODB.Recordset
        
         Cmd1.ActiveConnection = db 'sqlServer
         Cmd1.CommandType = adCmdStoredProc
         Cmd1.CommandText = "ap_genera_codigoregistro"
         Set Parm1 = Cmd1.CreateParameter("@codigo_doc", adVarChar, adParamInput, 200, codigo_doc)
         Cmd1.Parameters.Append Parm1
         rs.Open Cmd1
         rs.MoveFirst
         TxtCmpbte.Text = rs!Codigo
         rs.Close
    Else
        Call cambiarEtiquetaFactura
    End If
    If (Ado_datos.Recordset!estado_codigo_sol = "APR" And Ado_datos.Recordset!estado_codigo_fac = "REG") And (Ado_datos16.Recordset!venta_tipo = "E" Or Ado_datos16.Recordset!venta_tipo = "V" Or Ado_datos16.Recordset!venta_tipo = "C" Or Ado_datos16.Recordset!venta_tipo = "L") Then
      FraNavega.Enabled = False
'      fraOpciones.Visible = False
      FraGrabarCancelar.Visible = True
      FrmDetalle.Enabled = False
      FrmCobranza.Enabled = False
      'swgrabar = 0
      swnuevo = 2
      Txt_tdc.Text = GlTipoCambioMercado    'GlTipoCambioOficial
      SSTab1.Tab = 1
      SSTab1.TabEnabled(0) = False
      SSTab1.TabEnabled(1) = True
      SSTab1.TabEnabled(2) = False
      FrmCobros.Visible = True
      FrmCobros.Enabled = True
      FrmABMDet.Visible = False
      FrmABMDet2.Visible = False
      
'      BtnImprimir2.Visible = False
      BtnImprimir3.Visible = False
      CmdFoto.Visible = False
      If Ado_datos.Recordset!factura_impresa = "N" Then
      '    sino = MsgBox("Registrar� la cobranza efectiva, ahora ? ", vbYesNo, "Confirmando")
      '    If sino = vbYes Then
              'DTPFechaProg.Visible = True
              DTPFechaCobro.Visible = True
              DTPFechaCobro.Value = Date
'              Lbl_nombre_fac.Caption = "Factura a Nombre de:"
'              lbl_fechas.Caption = "Fecha de Cobranza"
              TxtCmpbte.Text = "0"
      '        Txt_parche.Visible = False      '&H80000013&
      '        'dtc_desc2A.BackColor = &H80000013
      '    Else
      '        DTPFechaProg.Visible = True
      '        DTPFechaCobro.Visible = False
      '        Lbl_nombre_fac.Caption = "Cliente :"
      '        lbl_fechas.Caption = "Fecha Programada de Cobranza"
      '        Txt_parche.Visible = True       '&H80000005&
      '        'dtc_desc2A.BackColor = &H80000005
      '    End If
      Else
      '    DTPFechaProg.Visible = True
      '    DTPFechaCobro.Visible = False
      '    lbl_fechas.Caption = "Fecha Programada de Cobranza"
      End If
      'TxtMonto.Text = CDbl(TxtDsctoTot)
      TxtMonto.SetFocus
    Else
      MsgBox "La Venta NO tiene saldo para cobrar o el Registro ya fue Aprobado !! ", vbExclamation, "Atenci�n!"
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atenci�n!"
  End If
End Sub

Private Sub BtnModificar1_Click()
  If Ado_datos01.Recordset.RecordCount > 0 Then
    If Ado_datos01.Recordset!estado_codigo_sol = "REG" Then
      SSTab1.Tab = 0
      SSTab1.TabEnabled(0) = True
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(2) = False
      'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
      FraNavega1.Enabled = False
'      fraOpciones1.Visible = False
      FraGrabarCancelar1.Visible = True
      FrmDetalle.Enabled = False
      FrmCobranza.Enabled = False
      FrmABMDet.Visible = False
      FrmABMDet2.Visible = False
      FrmCobros1.Visible = True
      FrmCobros1.Enabled = True
      swnuevo = 2
      DTPfechasol.Value = Date
      Txt_deposito.Text = "0"
      TxtMonto1.SetFocus
    Else
      MsgBox "No se puede editar, porque el Registro ya fue Aprobado !! ", vbExclamation, "Atenci�n!"
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atenci�n!"
  End If

End Sub

Private Sub BtnModificar2_Click()
  If Ado_datos02.Recordset.RecordCount > 0 Then
    If Ado_datos02.Recordset!estado_codigo_fac = "APR" And Ado_datos02.Recordset!estado_codigo_bco = "REG" Then
      SSTab1.Tab = 2
      SSTab1.TabEnabled(0) = False
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(2) = True
      'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
      FraNavega2.Enabled = False
'      fraOpciones1.Visible = False
      FraGrabarCancelar2.Visible = True
      FrmDetalle.Enabled = False
      FrmCobranza.Enabled = False
      FrmABMDet.Visible = False
      FrmABMDet2.Visible = False
      FrmCobros2.Visible = True
      FrmCobros2.Enabled = True
      swnuevo = 2
'      BtnAprobar3.Visible = True
      'DTPFechaCobro2.Value = Date
      'DTPFechaCobro02.Value = Date
      'Txt_deposito.Text = "0"
      TxtMonto02.SetFocus
    Else
      If Ado_datos02.Recordset!estado_codigo_bco = "APR" And Ado_datos02.Recordset!estado_codigo = "REG" And usr_codigo = "ASANTIVA�EZ" Then
            SSTab1.Tab = 2
          SSTab1.TabEnabled(0) = False
          SSTab1.TabEnabled(1) = False
          SSTab1.TabEnabled(2) = True
          'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
          FraNavega2.Enabled = False
    '      fraOpciones1.Visible = False
          FraGrabarCancelar2.Visible = True
          FrmDetalle.Enabled = False
          FrmCobranza.Enabled = False
          FrmABMDet.Visible = False
          FrmABMDet2.Visible = False
          FrmCobros2.Visible = True
          FrmCobros2.Enabled = True
          swnuevo = 2
          'DTPFechaCobro2.Value = Date
          'DTPFechaCobro02.Value = Date
          'Txt_deposito.Text = "0"
          TxtMonto02.SetFocus
      Else
            MsgBox "No se puede editar, porque el Registro ya fue Aprobado !! ", vbExclamation, "Atenci�n!"
      End If
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atenci�n!"
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
End Sub

Private Sub BtnModDetalle2_Click()
'  If ado_datos14.Recordset.RecordCount > 0 Then
'    SSTab1.Tab = 2
'    SSTab1.TabEnabled(2) = True
'    SSTab1.TabEnabled(0) = False
'    SSTab1.TabEnabled(1) = False
'
'    FrmEdita.Visible = True
''    BtnImprimir2.Visible = False
''    BtnImprimir3.Visible = False
'  Else
'    MsgBox "No Existen Bienes Registrados, Verifique por favor !! ", vbExclamation, "Atenci�n!"
'  End If

    'If Ado_datos.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            'CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_cobranzas_facturadas.rpt"
            CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_diaria_facturas.rpt"
            CryF02.WindowShowRefreshBtn = True
            'CryF02.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            'CryF02.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            'CryF02.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            'CryF02.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
        CryF02.Formulas(1) = "titulo = 'COBRANZAS' "
        CryF02.Formulas(2) = "subtitulo = 'LISTADO DE FACTURACION' "
            iResult = CryF02.PrintReport
            If iResult <> 0 Then MsgBox CryF02.LastErrorNumber & " : " & CryF02.LastErrorString, vbCritical, "Error de impresi�n"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
     'End If

End Sub

Private Sub BtnDesAprobar_Click()
'  sino = MsgBox("Esta seguro de Desaprobar el registro?", vbYesNo, "Confirmando")
'  If sino = vbYes Then
'    Dim rstdestino As New ADODB.Recordset
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correl_venta = " & Ado_datos.Recordset("correl_venta") & " and venta_codigo = " & Ado_datos.Recordset("venta_codigo") & " ", db, adOpenDynamic, adLockOptimistic
'    If Not rstdestino.BOF Then rstdestino.MoveFirst
'    If Not rstdestino.BOF And Not rstdestino.EOF Then
'      rstdestino("estado_codigo") = "REG"
'      rstdestino.Update
'    End If
'    If rstdestino.State = 1 Then rstdestino.Close
'    marca1 = Ado_datos.Recordset.Bookmark
'    Call OptFilGral1_Click
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo_fac = "REG" And Ado_datos.Recordset!factura_impresa = "N" Then
        Ado_datos.Recordset!estado_codigo_sol = "REG"
        Ado_datos.Recordset!estado_codigo_fac = "REG"
        Ado_datos.Recordset.Update
          'db.Execute "update ao_ventas_cobranza set estado_codigo_sol = 'APR' Where cobranza_codigo = " & Ado_datos01.Recordset("cobranza_codigo") & " "
    Else
        MsgBox "No se puede DEVOLVER, el registro ya fue FACTURADO, verifique los datos y vuelva a intentar ...", , "Atenci�n"
        Exit Sub
    End If
 Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atenci�n!"
 End If
End Sub

'Private Sub CmdDetallePoa_Click()
'  If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
'   MsgBox "No Existen Registros ", vbInformation, "Formulario 11"
'  Else
'    marca1 = Ado_datos.Recordset.BookMark
'    FrmPoasCapturaALB.Lblformulario = "F11"
'    FrmPoasCapturaALB.lblges_gestion = Ado_datos.Recordset!ges_gestion
'    FrmPoasCapturaALB.lblcodigo_unidad = Ado_datos.Recordset!codigo_unidad
'    FrmPoasCapturaALB.lblcodigo_solicitud = Ado_datos.Recordset!codigo_solicitud
'    FrmPoasCapturaALB.lbltipo_beneficiario = "N" 'Ado_datos.Recordset!tipoben_codigo
'    FrmPoasCapturaALB.Show vbModal
'  If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
'    '
'  Else
'    Ado_datos.Refresh
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
'  End If
'End Sub

'Private Sub cmdElige_Click()
'  With ALFrmMateriales
'        .ALPrincipal
'        If .QResp Then
'            txtCodigo.Text = .QCodigo
'            txtDesc.Text = .QItem
'        End If
'    End With
'    Txtcant_alm = 0
'    Cant_Alm = 0
'    DE.dbo_albSacaDetalleMaterial Mid(txtCodigo, 3, 12), descri_bien, Cant_Alm
'    Txtcant_alm = Cant_Alm
'    If Cant_Alm >= TxtCantPedi Then
'        optSi = True
'    Else
'        optNo = True
'    End If
'End Sub

Private Sub Contabiliza_venta()
'    Call graba_proyecto
    If VAR_SW = 1 Then
        Call graba_ingreso
    End If
    'If VAR_SW = 1 Then
        Set rstdestino = New ADODB.Recordset
        If VAR_TIPOV = "V" Or VAR_TIPOV = "C" Then
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
        Else
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEY' ", db, adOpenDynamic, adLockOptimistic
        End If
        If rstdestino.RecordCount > 0 Then
            VAR_CODANT = rstdestino!ingreso_codigo
            VAR_ORG = rstdestino!org_codigo
            VAR_FTE = rstdestino!fte_codigo
            If VAR_SW = 1 Then
                VAR_CODTIPO = "REF"
            Else
                VAR_CODTIPO = "REC"
            End If
            'Modificar con CASE WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW MAY-2015
            If VAR_COD4 = "DVTA" Then
               VAR_TSOL = "3"
               VAR_PARTIDA = "11200"
            Else
               VAR_TSOL = "10"
               VAR_PARTIDA = "11300"
            End If
            If VAR_COD4 = "DNMAN" Then
               VAR_TSOL = "10"
               VAR_PARTIDA = "11320"
            End If
            If VAR_COD4 = "DNREP" Then
               VAR_TSOL = "7"
               VAR_PARTIDA = "11330"
            End If
            If VAR_COD4 = "DNMOD" Then
               VAR_TSOL = "9"
               VAR_PARTIDA = "11340"
            End If
        End If
    'End If
  '===== Proceso para generar Asientos Contables Autom�ticos "DEI" y "REC"
  'sino = MsgBox("�Est� seguro de aprobar el Registro?", vbYesNo + vbQuestion, "CONFIRMAR...")
  'If sino = vbYes Then
    ' INI CORRECCION 18-JUN-2014
    Dim i As Integer
    Dim j As Integer
    Dim v_Tipo_Comp(1, 2)

'               gestion0 = Ado_datos.Recordset("ges_gestion")
'               correlv = Ado_datos.Recordset("venta_codigo")
'               nroventa = Ado_datos.Recordset("venta_codigo")
'
'               VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'               VAR_GLOSA = Ado_datos.Recordset!cobranza_observaciones
'               VAR_DOL2 = Round(Ado_datos.Recordset("cobranza_deuda_dol"), 2)
'               VAR_BS2 = Round(Ado_datos.Recordset("cobranza_deuda_bs"), 2)
'
'               VAR_PROY2 = Ado_datos16.Recordset!edif_codigo
'               VAR_COD4 = Ado_datos16.Recordset!unidad_codigo
'               VAR_TIPOV = Ado_datos16.Recordset!venta_tipo
'               VAR_SOL = Ado_datos16.Recordset!solicitud_codigo
'               VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
'                VAR_CODANT = rstdestino!ingreso_codigo
'            VAR_ORG = rstdestino!org_codigo
'            VAR_FTE = rstdestino!org_codigo
'            VAR_CODTIPO = "REC"
'            VAR_PARTIDA = "11200"

    fte_codigo1 = VAR_FTE
    '**** INI VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    Select Case VAR_CODTIPO
        Case "DEI"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
              'cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
              'Subcta_deb11 = rstdestino!Subcta_cred1
              'Subcta_deb21 = rstdestino!Subcta_cred2
    
              'cta_credito1 = rstdestino2!cta_deb
              'Subcta_cred11 = rstdestino2!Subcta_deb1
              'Subcta_cred21 = rstdestino2!Subcta_deb2
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor cont�ctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
        Case "REC"
            If VAR_MONEDA = "BOB" Then
                rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and SubCta_Deb1 = '01' ", db, adOpenKeyset, adLockReadOnly
            'rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            Else
                rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "  and SubCta_Deb1 = '02' ", db, adOpenKeyset, adLockReadOnly
            End If
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor cont�ctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
                        
            If rs_aux1.State = 1 Then rs_aux1.Close
            rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
            'rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
            If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
              If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
                MsgBox "El monto que est� intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
                'JQA FEB-2016
                'Exit Sub
              End If
            End If
            If rs_aux1.State = 1 Then rs_aux1.Close
        Case "REF"
            If rstdestino.State = 1 Then rstdestino.Close
            'rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REF' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REF' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "  ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor cont�ctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
                        
            If rs_aux1.State = 1 Then rs_aux1.Close
            rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
            'rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
            If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
              If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
                MsgBox "El monto que est� intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
                Exit Sub
              End If
            End If
            If rs_aux1.State = 1 Then rs_aux1.Close
            
        Case "DYR"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DYR' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor cont�ctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
            
        Case "DES"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DES' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor cont�ctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "ANI"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor cont�ctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "DVI"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DVI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor cont�ctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
            
            '' 02/07/2014 VERIFICAR
            'If rstdestino.State = 1 Then rstdestino.Close
            'rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
            'If rstdestino2.State = 1 Then rstdestino2.Close
            'rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            'If rstdestino.RecordCount < 1 Or rstdestino2.RecordCount < 1 Then
            '  MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor cont�ctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
            '  Exit Sub
            'End If
        Case Else
            MsgBox "No se ha definido el tipo " & vbCrLf & " de registro que est� procesando", vbOKOnly + vbCritical, "Error de aprobaci�n... "
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

    'If DtCCta_codigo.Text <> "01" Then
    '  If rstdestino.State = 1 Then rstdestino.Close
    '  rstFc_cuenta_bancaria.Find " cta_codigo = '" & DtCCta_codigo & "'", , adSearchForward, 1
    '  If Not rstFc_cuenta_bancaria.EOF Then
    '    fte_codigo1 = rstFc_cuenta_bancaria("fte_codigo")
    '  Else
    '  End If
    'Else
    '    fte_codigo1 = Me.DtCFte_codigo.Text
    'End If
    'If VAR_CODTIPO = "DEI" Or VAR_CODTIPO = "DES" Then
    '  fte_codigo1 = Me.DtCFte_codigo.Text
    'End If
    
'    fte_codigo1 = VAR_FTE
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim v_Tipo_Comp(1, 2)
'
'    v_Tipo_Comp(1, 1) = VAR_CODTIPO
    
'    If VAR_CODTIPO = "DYR" Then
'      'j = 2
'      'v_Tipo_Comp(1, 1) = "CAD"
'      'v_Tipo_Comp(1, 2) = "CAR"
'      j = 2
'      v_Tipo_Comp(1, 1) = "DYR"
'    Else
'      j = 1
'      v_Tipo_Comp(1, 1) = IIf(VAR_CODTIPO = "DEI", "DEI", IIf(VAR_CODTIPO = "REC", "REC", IIf(VAR_CODTIPO = "DES", "DES", IIf(VAR_CODTIPO = "ANI", "ANI", ""))))
'    End If
'
'    If VAR_CODTIPO = "DVI" Then
'      j = 1
'      v_Tipo_Comp(1, 1) = "DVI"
'    End If

'    For i = 1 To j
'      If rstdestino.State = 1 Then rstdestino.Close
'      If v_Tipo_Comp(1, i) = "DEI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "REC" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DYR" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DYR' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DES" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DES' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "ANI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DVI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DVI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "" Then
'        MsgBox "Antes de aprobar defina que tipo " & vbCrLf & "de registro est� procesando", vbOKOnly + vbCritical, "Error de aprobaci�n... "
'        Exit Sub
'      End If

    ' INI CORRECCION 18-JUN-2014
'      If v_Tipo_Comp(1, i) = "DVI" Then
'        ' 02/07/2014 VERIFICAR
'        If rs_aux2.State = 1 Then rs_aux2.Close
'        rs_aux2.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
'        If rstdestino2.State = 1 Then rstdestino2.Close
'        rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'        If rs_aux2.RecordCount < 1 Or rstdestino2.RecordCount < 1 Then
'          MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor cont�ctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'          Exit Sub
'        End If
'      End If
'
'      If rs_aux2.RecordCount < 1 Then
'        MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor cont�ctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'        Exit Sub
'      End If
'    Next

    'If rstdestino.State = 1 Then rstdestino.Close

    v_Tipo_Comp(1, 1) = VAR_CODTIPO
    
    db.BeginTrans
'    Frmmensaje.Visible = True
'    LblMensaje.Caption = "Este proceso tomar� solo unos segundos, gracias"
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
        rstCodComp.Open "select * from fc_Correl  where tipo_tramite = 'CMBTE'", db, adOpenDynamic, adLockOptimistic
        If rstCodComp.RecordCount > 0 Then
          Var_Comp = CDbl(rstCodComp!numero_correlativo)
          Var_Comp = Var_Comp + 1
          rstCodComp!numero_correlativo = Trim(Str(Var_Comp))
          rstCodComp.Update
        End If
        If rstCodComp.State = 1 Then rstCodComp.Close
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
      rs_aux2("ges_gestion") = glGestion        'Year(Date)
      'rstdestino("Num_Respaldo") = Ado_datos.Recordset("numero_documento")
      If yacontabilizo = 0 Then
        rs_aux2("Fecha_transacion") = Date
      End If
      rs_aux2("beneficiario_codigo") = VAR_BENEF
      rs_aux2("glosa") = "CONTABILIZA: " + VAR_GLOSA
      rs_aux2("unidad_codigo") = VAR_COD4           'Ado_datos16.Recordset("unidad_codigo")
      rs_aux2("solicitud_codigo") = VAR_SOL         'Ado_datos16.Recordset("solicitud_codigo")
      rs_aux2("tipo_moneda") = VAR_MONEDA
      rs_aux2("unidad_codigo_ant") = VAR_CITE
      
      rs_aux2("proceso_codigo") = "FIN"
      rs_aux2("subproceso_codigo") = "FIN-02"
      Select Case VAR_CODTIPO
        Case "DEI"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "REC"
            rs_aux2("etapa_codigo") = "FIN-02-03"
        Case "DYR"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "DES"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "ANI"
            rs_aux2("etapa_codigo") = "FIN-02-02"
        Case "REF"
            rs_aux2("etapa_codigo") = "FIN-02-02"
      End Select
      
      rs_aux2("clasif_codigo") = "ADM"
      rs_aux2("doc_codigo") = "R-110"
      rs_aux2("doc_numero") = Var_Comp
      rs_aux2("pro_codigo_det") = VAR_PROY2
    
      rs_aux2("estado_codigo") = "APR"

      If yacontabilizo = 0 Then
        rs_aux2("usr_codigo") = glusuario
        rs_aux2("Fecha_registro") = Format(Date, "dd/mm/yyyy")
        rs_aux2("Hora_registro") = Format(Time, "hh:mm:ss")
      End If
      rs_aux2.Update
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
'      If v_Tipo_Comp(1, i) = "DEI" Then     'Devengado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "REC" Then     'Recaudado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DYR" Then     'Devengado y Recaudado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DYR' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DES" Then     'Desafectado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DES' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "ANI" Then     'Anulado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DVI" Then     'Desafectado y Anulado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If

'      If v_Tipo_Comp(1, i) = "DVI" Then
'        ' VERIFICAR SI SE ESTA CONTROLANDA con el DYR
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
'        If rstdestino2.State = 1 Then rstdestino2.Close
'        rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'        If rstdestino.RecordCount > 0 And rstdestino2.RecordCount > 0 Then
'          cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
'          Subcta_deb11 = rstdestino!Subcta_cred1
'          Subcta_deb21 = rstdestino!Subcta_cred2
'
'          cta_credito1 = rstdestino2!cta_deb
'          Subcta_cred11 = rstdestino2!Subcta_deb1
'          Subcta_cred21 = rstdestino2!Subcta_deb2
'        Else
'          MsgBox "Rubro no presupuestado", vbCritical + vbOKOnly, "ERROR... "
''          Exit Sub
'        End If
'      End If
'
'      If rstdestino.RecordCount > 0 And v_Tipo_Comp(1, i) <> "DVI" Then
'        cta_deb1 = rstdestino("cta_deb")
'        Subcta_deb11 = rstdestino("Subcta_deb1")
'        Subcta_deb21 = rstdestino("Subcta_deb2")
'        cta_credito1 = rstdestino("cta_cred")
'        Subcta_cred11 = rstdestino("Subcta_cred1")
'        Subcta_cred21 = rstdestino("Subcta_cred2")
'      Else
'        'MsgBox "Rubro no presupuestado", vbCritical + vbOKOnly, "ERROR... "
'        'Exit Sub
'
'      End If
      
      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Or (VAR_CODTIPO = "DEY") Or (VAR_CODTIPO = "REF") Then
        cta_deb1 = rstdestino("cta_deb")
        Subcta_deb11 = rstdestino("Subcta_deb1")
        Subcta_deb21 = rstdestino("Subcta_deb2")
        
        cta_credito1 = rstdestino("cta_cred")
        Subcta_cred11 = rstdestino("Subcta_cred1")
        Subcta_cred21 = rstdestino("Subcta_cred2")
      Else
        cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
        Subcta_deb11 = rstdestino!Subcta_cred1
        Subcta_deb21 = rstdestino!Subcta_cred2
    
        cta_credito1 = rstdestino!cta_deb
        Subcta_cred11 = rstdestino!Subcta_deb1
        Subcta_cred21 = rstdestino!Subcta_deb2
      End If
      
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and SubCta1 = '" & Subcta_deb11 & "' and SubCta2 = '" & Subcta_deb21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        d_cta_nombre_1 = RTrim(rs_aux1("NombreCta"))
        d_aux1_1 = rs_aux1("aux1")
        d_aux2_1 = rs_aux1("aux2")
        d_aux3_1 = rs_aux1("aux3")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and SubCta1 = '" & Subcta_cred11 & "' and SubCta2 = '" & Subcta_cred21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        h_cta_nombre_1 = RTrim(rs_aux1("NombreCta"))
        h_aux1_1 = rs_aux1("aux1")
        h_aux2_1 = rs_aux1("aux2")
        h_aux3_1 = rs_aux1("aux3")
      End If
    ' nuevo fin
    
      '===== ini registra CO_diaRIO =========
      Set rstdestino2 = New ADODB.Recordset
      If rstdestino2.State = 1 Then rstdestino2.Close
      rstdestino2.Open "select * from co_diario where Cod_Comp = " & Var_Comp, db, adOpenKeyset, adLockOptimistic
      'If rstdestino2.RecordCount > 0 Then
      '  MsgBox "Ya Existe el asiento, se reemplazar� con los nuevos datos..."
      'Else
        rstdestino2.AddNew
        rstdestino2("Cod_Comp") = Var_Comp
      'End If
        rstdestino2("Cod_Comp_Detalle") = rstdestino2.RecordCount
      'rstdestino2("Tipo_Comp") = "DEI"   'v_Tipo_Comp(1, i)
      'rstdestino2("Cod_Comp_C") = Var_Comp
      'If v_Tipo_Comp(1, i) = "DEI" Or v_Tipo_Comp(1, i) = "REC" Then
      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Or (VAR_CODTIPO = "DEY") Or (VAR_CODTIPO = "REF") Then
        rstdestino2("D_Cuenta") = cta_deb1
        rstdestino2("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino2("D_Subcta1") = Subcta_deb11
        rstdestino2("D_SubCta2") = Subcta_deb21
        rstdestino2("D_Aux1") = d_aux1_1
        rstdestino2("D_Aux2") = d_aux2_1
        rstdestino2("D_Aux3") = d_aux3_1
        ' para Aux1
'        Select Case d_aux1_1
'                Case "01"
'                    VAR_COD1 = VAR_BENEF
'                Case "02"
'                    VAR_COD1 = VAR_CTA
'                Case "03"
'                    VAR_COD1 = VAR_PROY2
'                Case "04"
'                    VAR_COD1 = Ado_datos.Recordset("unidad_codigo")
'                Case "05"
'                    VAR_COD1 = ""
'                Case "06"
'                    VAR_COD1 = ""
'                Case "07"
'                    VAR_COD1 = ""
'                Case "08"
'                    VAR_COD1 = ""
'                Case "09"
'                    VAR_COD1 = VAR_ORG
'                Case "10"
'                    VAR_COD1 = ""
'                Case "11"
'                    VAR_COD1 = ""
'                Case "12"
'                    VAR_COD1 = ""
'        End Select
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
        Select Case d_aux1_1
            Case "01"
                rstdestino2("D_Cta_Aux1") = VAR_BENEF
                Call DESCAUX(d_aux1_1, CStr(VAR_BENEF))    'DESAUX =
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
                Call DESCAUX(d_aux1_1, CStr(VAR_CTA))
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
                Call DESCAUX(d_aux1_1, CStr(VAR_PROY2))
            Case "04"
                rstdestino2("D_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux1_1, CStr(VAR_COD4))
            Case "05"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux1") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux1_1, rstdestino2!D_Cta_Aux1)
            Case "07"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("D_Cta_Aux1") = VAR_ORG
                Call DESCAUX(d_aux1_1, CStr(VAR_ORG))
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
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
                Call DESCAUX(d_aux2_1, CStr(VAR_CTA))
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
                Call DESCAUX(d_aux2_1, CStr(VAR_PROY2))
            Case "04"
                rstdestino2("D_Cta_Aux2") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux2_1, CStr(VAR_COD4))
            Case "05"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux2") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux2_1, rstdestino2!D_Cta_Aux2)
            Case "07"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("D_Cta_Aux2") = VAR_ORG
                Call DESCAUX(d_aux2_1, CStr(VAR_ORG))
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
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
                Call DESCAUX(d_aux3_1, CStr(VAR_CTA))
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
                Call DESCAUX(d_aux3_1, CStr(VAR_PROY2))
            Case "04"
                rstdestino2("D_Cta_Aux3") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux3_1, CStr(VAR_COD4))
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux3_1, rstdestino2!D_Cta_Aux3)
            Case "07"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("D_Cta_Aux3") = VAR_ORG
                Call DESCAUX(d_aux3_1, CStr(VAR_ORG))
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
        If cta_deb1 = "6112" Then
            rstdestino2("D_MontoBs") = VAR_BS2 * 0.03
            rstdestino2("D_MontoDl") = VAR_DOL2 * 0.03
        Else
            If cta_credito1 = "2112" Then
                rstdestino2("D_MontoBs") = VAR_BS2 * 0.13
                rstdestino2("D_MontoDl") = VAR_DOL2 * 0.13
            Else
                If cta_deb1 = "1111" Then
                    rstdestino2("D_MontoBs") = VAR_BS2
                    rstdestino2("D_MontoDl") = VAR_DOL2
                Else
                    rstdestino2("D_MontoBs") = VAR_BS2 * 0.87
                    rstdestino2("D_MontoDl") = VAR_DOL2 * 0.87
                    'rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
                    'rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
                End If
            End If
        End If
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
        'rstdestino2("H_Cta_Aux1") = ""
        Select Case h_aux1_1
            Case "01"
                rstdestino2("H_Cta_Aux1") = VAR_BENEF
                Call DESCAUX(h_aux1_1, CStr(VAR_BENEF))
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
                Call DESCAUX(h_aux1_1, CStr(VAR_CTA))
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
                Call DESCAUX(h_aux1_1, CStr(VAR_PROY2))
            Case "04"
                rstdestino2("H_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux1_1, CStr(VAR_COD4))
            Case "05"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux1") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux1_1, rstdestino2!H_Cta_Aux1)
            Case "07"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("H_Cta_Aux1") = VAR_ORG
                Call DESCAUX(h_aux1_1, CStr(VAR_ORG))
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
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
                Call DESCAUX(h_aux2_1, CStr(VAR_CTA))
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
                Call DESCAUX(h_aux2_1, CStr(VAR_PROY2))
            Case "04"
                rstdestino2("H_Cta_Aux2") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux2_1, CStr(VAR_COD4))
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux2_1, rstdestino2!H_Cta_Aux2)
            Case "07"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("H_Cta_Aux2") = VAR_ORG
                Call DESCAUX(h_aux2_1, CStr(VAR_ORG))
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
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
                Call DESCAUX(h_aux3_1, CStr(VAR_CTA))
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
                Call DESCAUX(h_aux3_1, CStr(VAR_PROY2))
            Case "04"
                rstdestino2("H_Cta_Aux3") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux3_1, CStr(VAR_COD4))
            Case "05"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux3") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux3_1, rstdestino2!H_Cta_Aux3)
            Case "07"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("H_Cta_Aux3") = VAR_ORG
                Call DESCAUX(h_aux3_1, CStr(VAR_ORG))
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
        If cta_deb1 = "6112" Then
            rstdestino2("H_MontoBs") = VAR_BS2 * 0.03
            rstdestino2("H_MontoDl") = VAR_DOL2 * 0.03
        Else
            If cta_credito1 = "2112" Then
                rstdestino2("H_MontoBs") = VAR_BS2 * 0.13
                rstdestino2("H_MontoDl") = VAR_DOL2 * 0.13
            Else
                If cta_deb1 = "1111" Then
                    rstdestino2("H_MontoBs") = VAR_BS2
                    rstdestino2("H_MontoDl") = VAR_DOL2
                Else
                    rstdestino2("H_MontoBs") = VAR_BS2 * 0.87
                    rstdestino2("H_MontoDl") = VAR_DOL2 * 0.87
                    'rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
                    'rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
                End If
            End If
            'rstdestino2("H_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
            'rstdestino2("H_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
        End If
        rstdestino2("H_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
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

      If yacontabilizo = 0 Then
        rstdestino2("Usr_codigo") = glusuario
        rstdestino2("Fecha_registro") = Date
        rstdestino2("Hora_registro") = Format(Time, "hh:mm:ss")
      End If
      
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
    'LblMensaje.Caption = "El proceso concluy� exitosamente, gracias"
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
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "08"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
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
    End Select
    If rsAuxDetalle.RecordCount > 0 Then
      DESAUX = RTrim(rsAuxDetalle!DESAUX2)
    Else
      DESAUX = ""
    End If
End Function

'Private Sub f_actual_rec(org, codant)
'  Dim acumDl As Double
'  Dim rsrecalc As New ADODB.Recordset
'  Set rsrecalc = New ADODB.Recordset
'  If rsrecalc.State = 1 Then rsrecalc.Close
'  rsrecalc.Open "select sum(monto_dolares) as acumDl from fo_ingresos_cabecera where org_codigo = '" & org & "' and  correlativo_anterior = '" & codant & "' and codigo_tipo = 'REC' and estado_recaudado= 'S'", db, adOpenKeyset, adLockReadOnly
'  If rsrecalc.RecordCount > 0 Then
'    acumDl = IIf(IsNull(rsrecalc!acumDl), 0, rsrecalc!acumDl)
'  Else
'    acumDl = 0
'  End If
'  If rsrecalc.State = 1 Then rsrecalc.Close
'  rsrecalc.Open "select * from fo_ingresos_cabecera where org_codigo = '" & org & "' and correlativo_ingreso = '" & codant & "' ", db, adOpenKeyset, adLockOptimistic
'  If rsrecalc.RecordCount > 0 Then
'    rsrecalc!monto_recaudado_dolares = acumDl
'  End If
'  rsrecalc.Update
'  If rsrecalc.State = 1 Then rsrecalc.Close
'
'End Sub

Private Sub graba_proyecto()
'    Select Case Ado_datos.Recordset!unidad_codigo
'        Case "DNAJS", "DNEME", "DNINS", "DNMAN", "DNMOD", "DNREP"
'            VAR_PROY = 12
'        Case "UCOM"
'            VAR_PROY = 17
'        Case "DVTA"
'            VAR_PROY = 18
'
'    End Select
'
'    Set rs_aux1 = New ADODB.Recordset
'    If rs_aux1.State = 1 Then rs_aux1.Close
'    SQL_FOR = "select * from fo_proyectos_ejecucion where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
'    rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'    If rs_aux1.RecordCount > 0 Then
'        db.Execute "update fo_proyectos_ejecucion set pro_codigo_det_descripcion = '" & dtc_desc3.Text & "' Where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
'    Else
'        db.Execute "INSERT INTO fo_proyectos_ejecucion (pro_codigo, pro_codigo_det, pro_codigo_det_descripcion, unidad_codigo, ges_gestion, estado_codigo, usr_codigo, fecha_registro) " & _
'           "VALUES (" & VAR_PROY & ", '" & Ado_datos.Recordset!edif_codigo & "', '" & dtc_desc3.Text & "', '" & Ado_datos.Recordset!unidad_codigo & "', " & Ado_datos.Recordset!ges_gestion & ", 'APR', '" & GlUsuario & "', '" & Date & "')"
'    End If
End Sub

Private Sub graba_ingreso()
    '======= Ini grabado de datos
   'swgraba = 0
   'Call valida
    
'   If swgraba = 1 Then
'      FraOpciones2.Visible = False
'      fraOpciones.Visible = True
'      FraIngresosNav.Enabled = True
'      FraIngresosDat.Enabled = False
      
      'If v_a�adir = 1 Then
        'EFECTIVO o a CREDITO
         'db.BeginTrans
         'Call add_correl
         Set rstdestino = New ADODB.Recordset
         If VAR_TIPOV = "V" Or VAR_TIPOV = "C" Then
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
         Else
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEY' ", db, adOpenDynamic, adLockOptimistic
         End If
         
         If rstdestino.RecordCount > 0 Then
            VAR_CODANT = rstdestino!ingreso_codigo
            VAR_ORG = rstdestino!org_codigo
            VAR_FTE = rstdestino!fte_codigo
            'Call add_correl
         Else
            Call add_correl
            'EXEPCION PARA GRABAR CONTRATO EN INGRESOS
             rstdestino.AddNew
             rstdestino("Ges_Gestion") = Year(Date)     'Ado_datos.Recordset("ges_gestion")
             rstdestino("ingreso_codigo") = correlativo1
             rstdestino("org_codigo") = VAR_ORG
             rstdestino("ingreso_codigo_anterior") = VAR_ORG
             'rstdestino("Codigo_tipo") = "DEI"
             rstdestino("proceso_codigo") = "FIN"
             rstdestino("subproceso_codigo") = "FIN-01"
             rstdestino("etapa_codigo") = "FIN-01-02"
             rstdestino("clasif_codigo") = "ADM"
             rstdestino("doc_codigo") = "R-110"
             rstdestino("doc_numero") = correlativo1
             rstdestino("unidad_codigo") = VAR_COD4     'Ado_datos16.Recordset("unidad_codigo")
             rstdestino("solicitud_codigo") = VAR_SOL   'Ado_datos16.Recordset("solicitud_codigo")
             If VAR_COD4 = "DVTA" Then
                rstdestino("solicitud_tipo") = "3"
                VAR_PARTIDA = "11200"
                rstdestino("tipo_comp") = "DEY"
                rstdestino("Codigo_tipo") = "DEY"
             Else
                rstdestino("solicitud_tipo") = "10"
                VAR_PARTIDA = "11300"
                rstdestino("tipo_comp") = "DEI"
                rstdestino("Codigo_tipo") = "DEI"
             End If
             If VAR_COD4 = "DNMAN" Then
                rstdestino("solicitud_tipo") = "10"
                VAR_PARTIDA = "11320"
             End If
             If VAR_COD4 = "DNREP" Then
                rstdestino("solicitud_tipo") = "7"
                VAR_PARTIDA = "11330"
             End If
             If VAR_COD4 = "DNMOD" Then
                rstdestino("solicitud_tipo") = "9"
                VAR_PARTIDA = "11340"
             End If
             'OJO JQA
             rstdestino("beneficiario_codigo") = VAR_BENEF      'Ado_datos.Recordset("beneficiario_codigo")
             rstdestino("fecha_ingreso") = Date
             rstdestino("tipo_cambio") = GlTipoCambioMercado        'GlTipoCambioOficial
             rstdestino("tipo_moneda") = VAR_MONEDA
             'VAR_MONEDA = "BOB"
             rstdestino("ingreso_concepto") = "INGRESO POR: " + VAR_GLOSA       'Ado_datos.Recordset("cobranza_observaciones")
             'CAMBIAR FTE
             rstdestino("fte_codigo") = VAR_FTE
             'CAMBIAR RUBROS
             rstdestino("rubro_codigo") = VAR_PARTIDA
             'CAMBIAR RUBROS
             rstdestino("cheque_o_trf") = "T"
             'CAMBIAR CTA
             rstdestino("cta_codigo") = VAR_CTA
             If VAR_CTA = "NN" Then
                rstdestino("Bco_codigo") = "BCP"
             Else
                rstdestino("Bco_codigo") = "BMS"
             End If
             'CAMBIAR CTA
             rstdestino("numero_documento") = VAR_COD1
             rstdestino("unidad_codigo_ant") = VAR_CITE
             rstdestino("monto_dolares") = VAR_DOL2 * 12
             rstdestino("monto_bolivianos") = VAR_BS2 * 12
             rstdestino("monto_recaudado_dolares") = VAR_DOL2 * 12 'Round(Ado_datos.Recordset("cobranza_total_dol"), 2)
             rstdestino("monto_recaudado_bolivianos") = VAR_BS2 * 12   'Round(Ado_datos.Recordset("cobranza_total_bs"), 2)
             rstdestino("convenio_codigo") = "NN"
             rstdestino("pro_codigo_det") = VAR_PROY2       'Ado_datos16.Recordset("edif_codigo")
             rstdestino("estado_CODIGO") = "APR"
             'rstdestino("estado_codigo_dr") = "DEI"
    
             rstdestino("usr_CODIGO") = glusuario
             rstdestino("fecha_registro") = Date
             rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
             
             rstdestino.Update
             VAR_CODANT = rstdestino!ingreso_codigo
             VAR_ORG = rstdestino!org_codigo
             VAR_FTE = rstdestino!fte_codigo
             If rstdestino.State = 1 Then rstdestino.Close
             If VAR_TIPOV = "V" Or VAR_TIPOV = "C" Then
                rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
             Else
                rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEY' ", db, adOpenDynamic, adLockOptimistic
             End If
         End If
         Call add_correl
         ' OJO CAMBIAR FINANCIADOR WWWWWWWWWWWWWWWWWWWWW
         rstdestino.AddNew
         rstdestino("Ges_Gestion") = Year(Date)     'Ado_datos.Recordset("ges_gestion")
         rstdestino("ingreso_codigo") = correlativo1
         'VAR_CODANT = correlativo1
         'CAMBIAR org_codigo
         'rstdestino("org_codigo") = "111"
         'VAR_ORG = "111"
         rstdestino("org_codigo") = VAR_ORG
         'CAMBIAR org_codigo
         'CAMBIAR COD ingreso_codigo_anterior
         rstdestino("ingreso_codigo_anterior") = VAR_CODANT
         'CAMBIAR COD ingreso_codigo_anterior
         'CAMBIAR DEI O REC
         rstdestino("Codigo_tipo") = "REC"
         VAR_CODTIPO = "REC"
         'CAMBIAR DEI O REC
         rstdestino("proceso_codigo") = "FIN"
         rstdestino("subproceso_codigo") = "FIN-01"
         rstdestino("etapa_codigo") = "FIN-01-02"
         rstdestino("clasif_codigo") = "ADM"
         rstdestino("doc_codigo") = "R-110"
         rstdestino("doc_numero") = correlativo1
         rstdestino("unidad_codigo") = VAR_COD4
         rstdestino("solicitud_codigo") = VAR_SOL
         If VAR_COD4 = "DVTA" Then
            rstdestino("solicitud_tipo") = "3"
            VAR_PARTIDA = "11200"
         Else
            rstdestino("solicitud_tipo") = "10"
            VAR_PARTIDA = "11300"
         End If
         If VAR_COD4 = "DNMAN" Then
            rstdestino("solicitud_tipo") = "10"
            VAR_PARTIDA = "11320"
         End If
         If VAR_COD4 = "DNREP" Then
            rstdestino("solicitud_tipo") = "7"
            VAR_PARTIDA = "11330"
         End If
         If VAR_COD4 = "DNMOD" Then
            rstdestino("solicitud_tipo") = "9"
            VAR_PARTIDA = "11340"
         End If
         'OJO JQA
         rstdestino("beneficiario_codigo") = VAR_BENEF      'Ado_datos.Recordset("beneficiario_codigo")
         rstdestino("fecha_ingreso") = Date
         rstdestino("tipo_cambio") = GlTipoCambioMercado        'GlTipoCambioOficial
         rstdestino("tipo_moneda") = VAR_MONEDA
         'VAR_MONEDA = "BOB"
         rstdestino("ingreso_concepto") = "INGRESO POR: " + VAR_GLOSA       'Ado_datos.Recordset("cobranza_observaciones")
         'VAR_GLOSA = "INGRESO POR: " + Ado_datos.Recordset("cobranza_observaciones")
         If Ado_datos16.Recordset("venta_tipo") = "E" Then
            rstdestino("tipo_comp") = "DYR"
         Else
            rstdestino("tipo_comp") = "REC"
         End If
         'CAMBIAR FTE
         rstdestino("fte_codigo") = VAR_FTE
         'CAMBIAR FTE OJO JQAW
         'CAMBIAR RUBROS
         rstdestino("rubro_codigo") = VAR_PARTIDA
         'CAMBIAR RUBROS
         rstdestino("cheque_o_trf") = "T"
         'CAMBIAR CTA
         rstdestino("cta_codigo") = VAR_CTA
         If VAR_CTA = "2015046557-03-054" Then
            rstdestino("Bco_codigo") = "BCP"
         Else
            rstdestino("Bco_codigo") = "BMS"
         End If
         'CAMBIAR CTA
         NroFactura = Trim(Str(VAR_COD1))
         rstdestino("numero_documento") = NroFactura        'Ado_datos.Recordset!cobranza_nro_factura
         rstdestino("unidad_codigo_ant") = VAR_CITE
         rstdestino("monto_dolares") = VAR_DOL2
         rstdestino("monto_bolivianos") = VAR_BS2
         rstdestino("monto_recaudado_dolares") = VAR_DOL2   'Round(Ado_datos.Recordset("cobranza_total_dol"), 2)
         rstdestino("monto_recaudado_bolivianos") = VAR_BS2     'Round(Ado_datos.Recordset("cobranza_total_bs"), 2)
         rstdestino("convenio_codigo") = "NN"
         rstdestino("pro_codigo_det") = VAR_PROY2       'Ado_datos16.Recordset("edif_codigo")
         rstdestino("estado_CODIGO") = "APR"
         'rstdestino("estado_codigo_dr") = "DEI"

         rstdestino("usr_CODIGO") = glusuario
         rstdestino("fecha_registro") = Date
         rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
         
         rstdestino.Update
         If rstdestino.State = 1 Then rstdestino.Close
        'db.CommitTrans
          
'          If rstIngresos.State = 1 Then rstIngresos.Close
'          rstIngresos.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
'          rstIngresos.Sort = "ingreso_codigo"
'          rstIngresos.Requery
          
'          rstIngresos.Requery
'          Set AdoIngresos.Recordset = rstIngresos
'          AdoIngresos.Refresh
'          AdoIngresos.Recordset.Find "ultimo = 'S'"
'          If Not (AdoIngresos.Recordset.EOF) Then
'            marca1 = AdoIngresos.Recordset.Bookmark
'            AdoIngresos.Recordset("ultimo") = "N"
'            AdoIngresos.Recordset.Update
'          End If

'          AdoIngresos.Recordset.Move marca1 - 1

'          marca1 = 0
      'End If
'   Else*
'      MsgBox "ERROR Los datos no est�n completos, no se realizar� la grabaci�n..."
''      FraOpciones2.Visible = False
''      FraOpciones.Visible = True
''      FraIngresosNav.Enabled = True
''      FraIngresosDat.Enabled = False
''      AdoIngresos.Refresh
'   End If
'   LblAccion = ""
'AAQQQQQUIIIIIIIIII    JQA

End Sub

Private Sub add_correl()
  Set rstcorrel_ing = New ADODB.Recordset
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
  rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '" & VAR_ORG & "' ", db, adOpenDynamic, adLockOptimistic
  'rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '111' ", db, adOpenDynamic, adLockOptimistic
  If rstcorrel_ing.RecordCount = 0 Then
     VAR_ORG = "112"
     VAR_FTE = "10"
     If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
     rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '" & VAR_ORG & "'  ", db, adOpenDynamic, adLockOptimistic
     rstcorrel_ing("correlativo_ingreso") = rstcorrel_ing("correlativo_ingreso") + 1
     rstcorrel_ing.Update
     correlativo1 = rstcorrel_ing("correlativo_ingreso")
'     rstcorrel_ing.AddNew
'     rstcorrel_ing("org_codigo") = VAR_ORG   'Trim(DtCorg_codigo.Text)
'     rstcorrel_ing("ges_gestion") = Ado_datos.Recordset("ges_gestion")  'Trim(lblges_gestion.Caption)
'     rstcorrel_ing("fte_codigo") = "10"
'     'rstcorrel_ing("correlativo") = 1
'     rstcorrel_ing("correlativo_ingreso") = 1
'     rstcorrel_ing.Update
'     correlativo1 = rstcorrel_ing("correlativo_ingreso")
'     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo_ingreso")
  Else
     rstcorrel_ing("correlativo_ingreso") = rstcorrel_ing("correlativo_ingreso") + 1
     rstcorrel_ing.Update
     correlativo1 = rstcorrel_ing("correlativo_ingreso")
     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo")
  End If
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close

End Sub

'Private Sub CmdGrabaCobranza()
'    If swnuevo = 1 Then
''      rstdestino.Open "select * from ao_ventas_detalle where correl_venta = " & lblcorrelVenta & " and venta_codigo = " & TxtNroVenta, db, adOpenKeyset, adLockOptimistic
''      Set Ado_datos16.Recordset = rstdestino
''      Ado_datos16.Recordset.AddNew
'      Ado_datos16.Recordset!correl_venta = Val(lblcorrelVenta.Caption)
'      Ado_datos16.Recordset!venta_codigo = Val(TxtNroVenta.Text)
'      Ado_datos16.Recordset!ges_gestion = Year(Date)    'Trim(LblGestion.Caption)
'    End If
'      Ado_datos16.Recordset!beneficiario_codigo = dtc_codigo2A.Text                                 'Codigo Beneficiario/Cliente
'      Ado_datos16.Recordset!ci = dtc_codigo4A.Text                                                     'Codigo Cobrador
'      Ado_datos16.Recordset!nombre_cobrador = dtc_desc4A.Text + " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
'      Ado_datos16.Recordset!deuda_cobrada = Val(TxtMonto.Text)                                  'Monto Cobrado
'      Ado_datos16.Recordset!deuda_cobrada_dol = Val(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
'      Ado_datos16.Recordset!fecha_cobranza = DTPFechaCobro.Value                                'Fecha de Cobranza
'      'Call acumulaMont(Ado_datos16.Recordset!ges_gestion, Ado_datos16.Recordset!correl_venta, Ado_datos16.Recordset!venta_codigo)
'      Call acumulaMont(Ado_datos16.Recordset("ges_gestion"), Ado_datos16.Recordset("venta_codigo"))
'
'      Ado_datos16.Recordset!obs_cobranza = TxtObs
'      Ado_datos16.Recordset!nro_cmpbte = Trim(TxtCmpbte)
'      Ado_datos16.Recordset!usr_usuario = GlUsuario
'      Ado_datos16.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
'      Ado_datos16.Recordset!hora_registro = Format(Time, "hh:mm:ss")
'      Ado_datos16.Recordset.Update
'End Sub

'Private Sub CmdModDetalle_Click()
'  FraDetalle.Visible = True
'  FraDetalle.Enabled = True
'  txtnosolicitud1.Enabled = False
'  txtcorrdet.Enabled = False
'  dtccodpar.SetFocus
'  CmdGraDetalle.Enabled = True
'  CmdAddDetalle.Enabled = False
'  CmdModDetalle.Enabled = False
'  CmdSalDetalle.Enabled = False
'  CmdCanDetalle.Enabled = True
'  swgrabar = 2
'End Sub

'Private Sub CmdGraDetalle_Click()
'    If swgrabar = 1 Then
'        Dim rstdestino As New ADODB.Recordset
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_solicitud_detalle_correl where formulario = '" & "F11" & "' and correl_solicitud = " & Ado_datos.Recordset("codigo_solicitud"), db, adOpenDynamic, adLockOptimistic
'        If Not (rstdestino.EOF) Then
'            rstdestino("correl_solicitud_detalle") = rstdestino("correl_solicitud_detalle") + 1
'        Else
'            rstdestino.AddNew
'            rstdestino("formulario") = "F11"
'            rstdestino("correl_solicitud") = Ado_datos.Recordset("codigo_solicitud")
'            rstdestino("correl_solicitud_detalle") = 1
'        End If
'        correldetalle = rstdestino("correl_solicitud_detalle")
'        rstdestino.Update
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_solicitud_detalle where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correlativo_solicitud = " & Ado_datos.Recordset("codigo_solicitud"), db, adOpenDynamic, adLockOptimistic
'        rstdestino.AddNew
'        rstdestino("ges_gestion") = Ado_datos.Recordset("ges_gestion")
'        rstdestino("correlativo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
'        rstdestino("correlativo_detalle") = correldetalle
'        rstdestino("Par_codigo") = dtccodpar.Text
'        rstdestino("Importe_nacional") = txtsolpeso.Text
'        rstdestino("formulario") = "F11"
'        rstdestino.Update
'        If rstdestino.State = 1 Then rstdestino.Close
'        Set rs_datos14 = New ADODB.Recordset
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        rs_datos14.Open "select * from ao_solicitud_detalle WHERE ges_gestion = '" & Trim(Ado_datos.Recordset("ges_gestion")) & "' and correlativo_solicitud = " & Trim(Ado_datos.Recordset("codigo_solicitud")) & " and formulario = 'F11'", db, ad0OpenKeyset, adLockOptimistic
'        Set adoDetalleSolicitud.Recordset = rs_datos14
'        adoDetalleSolicitud.Refresh
'    End If
'    If swgrabar = 2 Then
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adoDetalleSolicitud.Recordset("ges_gestion") & "' and correlativo_solicitud = " & adoDetalleSolicitud.Recordset("correlativo_solicitud") & " and correlativo_detalle =" & adoDetalleSolicitud.Recordset("correlativo_detalle"), db, adOpenDynamic, adLockOptimistic
'        If Not (rstdestino.EOF) Then
'            rstdestino("ges_gestion") = Ado_datos.Recordset("ges_gestion")
'            rstdestino("correlativo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
'            rstdestino("correlativo_detalle") = correldetalle
'            rstdestino("Par_codigo") = dtccodpar.Text
'            rstdestino("Importe_nacional") = txtsolpeso.Text
'            rstdestino("formulario") = "F11"
'            rstdestino.Update
'        End If
'        If rstdestino.State = 1 Then rstdestino.Close
'        Set rs_datos14 = New ADODB.Recordset
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        rs_datos14.Open "select * from ao_solicitud_detalle WHERE ges_gestion = '" & Trim(Ado_datos.Recordset("ges_gestion")) & "' and correlativo_solicitud = " & Trim(Ado_datos.Recordset("codigo_solicitud")) & " and formulario = 'F11'", db, ad0OpenKeyset, adLockOptimistic
'        Set adoDetalleSolicitud.Recordset = rs_datos14
'        adoDetalleSolicitud.Refresh
'    End If
'    CmdGraDetalle.Enabled = False
'    CmdAddDetalle.Enabled = True
'    CmdModDetalle.Enabled = True
'    CmdSalDetalle.Enabled = True
'    CmdCanDetalle.Enabled = False
'    FraDetalle.Enabled = False
'    swgrabar = 0
'End Sub

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
                rstpagos("ges_gestion") = Ado_datos.Recordset("ges_gestion")
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

Private Sub CmdGrabaCobro_Click()
End Sub

'Private Sub CmdGrabaDet_Click()
''If dtc_desc12 = "" Then
''    MsgBox "Debe Elejir un Descuento X Tipo de Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atenci�n"
''    Exit Sub
''  End If
'  If dtc_codigo15 = "" Then
'     MsgBox "Debe Elejir un Producto para Vender, !! Vuelva a Intentar ...", vbExclamation, "Atenci�n"
'    Exit Sub
'  End If
''  If dtc_desc13 = "" Then
''    MsgBox "Debe Elejir el Almacen de Origen, !! Vuelva a Intentar ...", vbExclamation, "Atenci�n"
''    Exit Sub
''  End If
'    'If Val(dtc_stocktotal15.Text) >= Val(TxtCantidad.Text) Then
'    '    VAR_PARTIDA = "OK"
'    If Val(Dtc_Stock13.Text) >= Val(TxtCantidad.Text) Or Dtc_partida15.Text = "43340" Then
'          'fraOpciones.Visible = True
'          'FraGrabarCancelar.Visible = False
'          'TxtNroVenta.Enabled = True
'          FrmEdita.Enabled = False
'        '  DtGListaN.Enabled = True
'          'cmdElige.Enabled = False
'        '  dtc_codigo15.Visible = False
'        '  dtc_desc15.Visible = False
'          'txt_descripcion_venta.Enabled = False
'        If swnuevo = 1 Then
'          'ado_datos14.Recordset!venta_codigo_det = Ado_datos.Recordset("correl_venta")
'          ado_datos14.Recordset!venta_codigo = Ado_datos.Recordset("venta_codigo")
'          ado_datos14.Recordset!ges_gestion = Ado_datos.Recordset("ges_gestion")
'        End If
'          'ado_datos14.Recordset!nro_licitacion = dtc_partida15.Text                       'Compra ??
'          'ado_datos14.Recordset!nro_adjudica = 0 'Trim(DtcNroAdjudica.Text)                 'Codigo de Adjudicacion
'          ado_datos14.Recordset!bien_codigo = Trim(dtc_codigo15.Text)                       'Codigo Bien (Equipo, Producto, etc)
'          ado_datos14.Recordset!grupo_codigo = Trim(dtc_grupo15.Text)
'          ado_datos14.Recordset!subgrupo_codigo = Trim(dtc_subgrupo15.Text)
'          ado_datos14.Recordset!par_codigo = Dtc_partida15                              'Partida
'          ado_datos14.Recordset!tipo_descuento = IIf(dtc_codigo12.Text = "", "0", dtc_codigo12.Text)                      ' Tipo de Descuento
'          ado_datos14.Recordset!concepto_venta = txt_descripcion_venta                  'Descripcion y Caracteristicas
'          ado_datos14.Recordset!almacen_codigo = IIf(dtc_codigo13.Text = "", "0", dtc_codigo13.Text)
'          If TxtCantidad.Text = "" Then
'            TxtCantidad.Text = "1"
'          End If
'          ado_datos14.Recordset!venta_det_cantidad = Val(IIf(TxtCantidad = "", 1, TxtCantidad)) 'Cantidad Vendida
'          'ado_datos14.Recordset!codigo_solicitud = 0                                     'Nro.Solicitud de compra
'          ado_datos14.Recordset!venta_precio_unitario_bs = CDbl(TxtPrecioU.Text)             'Precio Unitario de Venta
'          If CDbl(TxtDescuento) > 0 Then
'            ado_datos14.Recordset!venta_descuento_bs = CDbl(TxtDescuento.Text)      'Dcto por producto CON DESCUENTO
'            ado_datos14.Recordset!venta_descuento_dol = Val(TxtDescuento) / GlTipoCambioMercado
'          Else
'            'ado_datos14.Recordset!descuento_venta = (Val(TxtCantidad) * CDbl(TxtPrecioU.Text)) * (CDbl(Dtc_aux12)) 'Dcto por producto DE LA TABLA
'            TxtDescuento.Text = "0"
'            ado_datos14.Recordset!venta_descuento_bs = 0
'            ado_datos14.Recordset!venta_descuento_dol = 0
'          End If
'          ado_datos14.Recordset!venta_precio_total_bs = (Val(TxtCantidad) * CDbl(TxtPrecioU.Text)) - (CDbl(TxtDescuento)) 'Precio Total Producto
'          'If Val(lbltipo_Cambio) = 0 Then lbltipo_Cambio = 1
'          ado_datos14.Recordset!venta_precio_unitario_dol = CDbl(TxtPrecioU.Text) / GlTipoCambioMercado                'Precio Unitario Dolares
'          ado_datos14.Recordset!venta_precio_total_dol = (ado_datos14.Recordset!venta_precio_total_bs) / GlTipoCambioMercado
'          'Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"), Ado_datos.Recordset("venta_codigo"))
'          ado_datos14.Recordset!modelo_codigo = Txt_modelo.Text
'          ado_datos14.Recordset!modelo_codigo1 = Txt_modelo1.Text
'          ado_datos14.Recordset!modelo_codigo_h = Txt_modelo2.Text
'          ado_datos14.Recordset!modelo_codigo_x = Txt_modelo3.Text
'          If OpMod1.Value = True Then
'            ado_datos14.Recordset!modelo_elegido = "S"
'          Else
'            ado_datos14.Recordset!modelo_elegido = "N"
'          End If
'          If OpMod2.Value = True Then
'            ado_datos14.Recordset!modelo_elegido_h = "S"
'          Else
'            ado_datos14.Recordset!modelo_elegido_h = "N"
'          End If
'          If OpMod2.Value = True Then
'            ado_datos14.Recordset!modelo_elegido_x = "S"
'          Else
'            ado_datos14.Recordset!modelo_elegido_x = "N"
'          End If
'          ado_datos14.Recordset!estado_codigo = "REG"
'          ado_datos14.Recordset!usr_codigo = GlUsuario
'          ado_datos14.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
'          ado_datos14.Recordset!hora_registro = Format(Time, "hh:mm:ss")
'          ado_datos14.Recordset.Update
'        'db.CommitTrans
'
'        'Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"), Ado_datos.Recordset("venta_codigo"))
'        Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"))
'        SSTab1.Tab = 0
'        SSTab1.TabEnabled(0) = True
'        SSTab1.TabEnabled(1) = False
'        SSTab1.TabEnabled(2) = False
'        FraNavega.Enabled = True
'        FrmDetalle.Enabled = True
'        'FrmDetalle.Visible = True
'        FrmCobranza.Visible = True
'        FrmABMDet.Visible = True
'        FrmABMDet2.Visible = True
'        Call OptFilGral1_Click
'        If swnuevo = 1 Then
'          'Call abre_ventas_det
'          'rs_datos14.Requery
'          'ado_datos14.Refresh
'          'ado_datos14.Recordset.MoveLast
'
'        End If
'        swnuevo = 0
'    Else
'        MsgBox "Saldo Insuficiente en Almacen Origen, debe realizar Transferencia de otro Almacen, Luego Intente nuevamente !..."
'    End If
'  'Else
'  '  MsgBox "Saldo Insuficiente en Stock General (Todos los Almacenes), Intente nuevamente !..."
'  'End If
'End Sub

'Private Sub BtnImprimir2_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    Dim iResult As Variant  ', i%, y%
'    CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'    CryR01.WindowShowRefreshBtn = True
''    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
''    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
''    CryR01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_codigo
'    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
'    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
'
'    CryR01.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
'    CryR01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
'    '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'    iResult = CryR01.PrintReport
'    If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresi�n"
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
'  End If
'End Sub

'Private Sub BtnModDetalle_Click()
'  If Ado_datos16.Recordset.RecordCount > 0 Then
'
'    SSTab1.Tab = 1
'    SSTab1.TabEnabled(1) = True
'    SSTab1.TabEnabled(0) = False
'    SSTab1.TabEnabled(2) = False
'
'    FrmCabecera.Visible = True
''    BtnImprimir2.Visible = False
''    BtnImprimir3.Visible = False
'  Else
'    MsgBox "No existen datos de la Venta, Verifique por favor !! ", vbExclamation, "Atenci�n!"
'  End If
'End Sub

Private Sub BtnSalir2_Click()
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
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
'
'    FrmCabecera.Visible = False
''    BtnImprimir2.Visible = True
''    BtnImprimir3.Visible = True
End Sub

Private Sub BtnSalir3_Click()
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
'
'    FrmEdita.Visible = False
''    BtnImprimir2.Visible = True
''    BtnImprimir3.Visible = True
End Sub

Private Sub BtnSalir1_Click()
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

Private Sub cmd_benef_Click()
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

Private Sub cmd_moneda1_LostFocus()
    Set rs_datos20 = New ADODB.Recordset
    If rs_datos20.State = 1 Then rs_datos20.Close
    rs_datos20.Open "Select * from fc_cuenta_bancaria where tipo_moneda = '" & cmd_moneda1.Text & "' ", db, adOpenStatic
    Set Ado_datos20.Recordset = rs_datos20
    dtc_ctades.BoundText = dtc_cta.BoundText
End Sub

Private Sub cmd_moneda2_LostFocus()
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from fc_cuenta_bancaria where tipo_moneda = '" & cmd_moneda2.Text & "' ", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub CmdFoto_Click()
'    Frm_Imprime_Factura.Show

    On Error GoTo QError
    Set fs = New FileSystemObject   'Creamos la Nueva referencia Fso
    
    Set rs_aux6 = New ADODB.Recordset     'Iniciales del Cliente - gc_beneficiario
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "Select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' ", db, adOpenStatic
    If rs_aux6.RecordCount > 0 Then
        db.Execute "update ao_ventas_cobranza set beneficiario_iniciales = '" & rs_aux6!beneficiario_iniciales & "'   Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
    End If
    'If Ado_datos.Recordset!ARCHIVO_FOTO = "Cargar_Archivo" Then
    If Ado_datos.Recordset!archivo_foto_cargado = "N" Or IsNull(Ado_datos.Recordset!archivo_foto_cargado) Then
      NombreCarpeta = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
      DirOrigen = App.Path & "\CLIENTES\"
      DirDestino = App.Path & "\CLIENTES\"
      'DirDestino = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
      fs.CopyFile DirOrigen & "\QRCode.bmp", DirDestino & "\" & Ado_datos.Recordset!doc_codigo_fac & "-" & Trim(Str(VAR_COD1)) & ".JPG"       'Ado_datos.Recordset!cobranza_nro_factura        'ARCHIVO_Foto
      Ado_datos.Recordset!ARCHIVO_Foto = Trim(Ado_datos.Recordset!doc_codigo_fac & "-" & Trim(Str(VAR_COD1)) & ".JPG")
      Ado_datos.Recordset!archivo_foto_cargado = "S"
      
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "Q_R"
''      If GlServidor = "SERVIDOR2" Then
''         e = "\\" & Trim(GlServidor) & "\SIGPER\CLIENTES\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
''      Else
'         e = NombreCarpeta
''      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
    Else
      'MsgBox ""
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atenci�n")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
          DirOrigen = App.Path & "\CLIENTES\"
          DirDestino = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
          fs.CopyFile DirOrigen & "\QRCode.bmp", DirDestino & "\" & Ado_datos.Recordset!ARCHIVO_Foto
          frmBeneficiario_Admin.Adolista.Recordset!archivo_foto_cargado = "S"
          
    '      Frmexporta.DirDestino.Path = NombreCarpeta
    '      GlArch = "Q_R"
    ''      If GlServidor = "SERVIDOR2" Then
    ''         e = "\\" & Trim(GlServidor) & "\SIGPER\CLIENTES\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
    ''      Else
    '         e = NombreCarpeta
    ''      End If
    '      Frmexporta.DirDestino2.Path = e
    '      Frmexporta.Show vbModal      End If
      End If
    End If

    Dim ARCH_FOTO As String
'    If GlServidor = "SERVIDOR2" Then
'        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\CLIENTES\" + Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    Else
        'ARCH_FOTO = App.Path + "\CLIENTES\" + Trim(rs_aux6!beneficiario_iniciales) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_Foto)
        ARCH_FOTO = App.Path + "\CLIENTES\" + Trim(Ado_datos.Recordset!ARCHIVO_Foto)
'    End If
    'ARCH_FOTO = App.Path + "\" + "CLIENTES" + "\" + Ado_datos.Recordset!beneficiario_codigo + "\" + Ado_datos.Recordset("beneficiario_codigo") + "-FOTO.JPG"
    CodBenef = Ado_datos.Recordset!cobranza_codigo
    'If Guardar_Imagen(db, "Select Foto From Gc_beneficiario Where beneficiario_codigo= '" & CodBenef & "' ", "Foto", ARCH_FOTO) Then
    If Guardar_Imagen(db, "Select Foto From ao_ventas_cobranza Where cobranza_codigo= '" & CodBenef & "' ", "Foto", ARCH_FOTO) Then
        MsgBox "Se cargo la Imagen Correctamente !!"
        Exit Sub
    Else
        MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    End If
QError:
    ' Manejo de errores
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atenci�n"
'    db.RollbackTrans
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    'dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub BntImprimir3_Click()
    'If Ado_datos.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_cobranzas_facturadas_dol.rpt"
            'CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_diaria_facturas.rpt"
            CryF02.WindowShowRefreshBtn = True
            'CryF02.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            'CryF02.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            'CryF02.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            'CryF02.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
        CryF02.Formulas(1) = "titulo = 'COBRANZAS' "
        CryF02.Formulas(2) = "subtitulo = 'ESTADO DE CUENTAS' "
            iResult = CryF02.PrintReport
            If iResult <> 0 Then MsgBox CryF02.LastErrorNumber & " : " & CryF02.LastErrorString, vbCritical, "Error de impresi�n"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
     'End If
End Sub



Private Sub dtc_aux8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_aux8.BoundText
    dtc_codigo8.BoundText = dtc_aux8.BoundText
End Sub

Private Sub dtc_codigo4A1_Click(Area As Integer)
    dtc_desc4A1.BoundText = dtc_codigo4A1.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

'Private Sub dtc_codigo61_Click(Area As Integer)
'    dtc_desc61.BoundText = dtc_codigo61.BoundText
'End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    dtc_aux8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_cta_Click(Area As Integer)
    dtc_ctades.BoundText = dtc_cta.BoundText
End Sub

Private Sub dtc_ctades_Click(Area As Integer)
    dtc_cta.BoundText = dtc_ctades.BoundText
End Sub

Private Sub dtc_desc4A1_Click(Area As Integer)
    dtc_codigo4A1.BoundText = dtc_desc4A1.BoundText
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
End Sub

'Private Sub dtc_desc61_Click(Area As Integer)
'    dtc_desc61.BoundText = dtc_codigo61.BoundText
'End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
    dtc_aux8.BoundText = dtc_desc8.BoundText
End Sub

Private Sub dtc_aux5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo4A_Click(Area As Integer)
    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    DataCombo2.Text = DataCombo1.BoundText
End Sub

Private Sub DataCombo2_Click(Area As Integer)
    DataCombo1.Text = DataCombo2.BoundText
End Sub

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

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    dtc_aux5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub DTPFechaCobro02_LostFocus()
    If (CDate(DTPFechaCobro2.Value) > CDate(DTPFechaCobro02.Value)) Then
        MsgBox "La <<Fecha Cobranza2>> No puede ser MENOR a la <<Fecha Cobranza1>>, Vuelva a Intentar !! ", vbExclamation, "Atenci�n!"
        DTPFechaCobro02.SetFocus
    End If
End Sub

'Private Sub DTPfechasol_Change()
'    txtGes_gestion = CStr(Year(DTPfechasol.Value))
'End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = 0
    parametro = Aux
    '
    Call ABRIR_TABLAS_AUX
    Call OptFilGral01_Click
    'txt_codigo.Enabled = True
    mbDataChanged = False
'    FrmCabecera.Enabled = False
    FrmCobros.Enabled = False
    FrmCobros1.Enabled = False
    dg_datos.Enabled = True
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    GlNombFor = "F04"
    FraGrabarCancelar1.Visible = False
    'LblUsuario.Caption = GlUsuario
    marca1 = 1
    deta2 = 0
'    BtnImprimir2.Visible = True
    If glusuario = "RCUELA" Or glusuario = "ADMIN" Or glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS" Then
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
'        SSTab1.TabEnabled(2) = True
    Else
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
'        SSTab1.TabEnabled(2) = True
    End If
'    FrmEdita.Enabled = False
'    Cmd_Cliente.Visible = False
    swnuevo = 0
    FraNavega.Caption = lbl_titulo.Caption
    'lbl_titulo2.Caption = lbl_titulo.Caption
    'lbl_titulo1.Caption = lbl_titulo.Caption
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    'dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset     'Beneficiario Personas Nat. y Juridicas
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset     'Proyecto de Edificaci�n
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from gc_edificaciones order by edif_denominacion", db, adOpenStatic
    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    Set rs_datos4 = New ADODB.Recordset     'Beneficiario Funcionario - Cobrador en Fac.
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
    'rs_datos4.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_datos4.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
    
    Set rs_datos4A = New ADODB.Recordset     'Beneficiario Funcionario - Cobrador
    If rs_datos4A.State = 1 Then rs_datos4A.Close
    'rs_datos4A.Open "gp_listar_gc_beneficiario_funcionario ", db, adOpenStatic  '4333735
    'rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set ado_datos4A.Recordset = rs_datos4A
    dtc_desc4A1.BoundText = dtc_codigo4A1.BoundText

    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from gc_tipo_transaccion order by trans_descripcion", db, adOpenStatic
    'rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
    
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    rs_datos11.Open "ac_tipo_compra_venta", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
'    dtc_desc11.BoundText = dtc_codigo11.BoundText

    Set rs_datos13 = New ADODB.Recordset    'Detalle por cada Almacen
    If rs_datos13.State = 1 Then rs_datos13.Close
    'rs_datos13.Open "select * from Av_DestinoDet", db, adOpenKeyset, adLockReadOnly
    rs_datos13.Open "select * from av_almacen_detalle", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos13.Recordset = rs_datos13
    Ado_datos13.Refresh
    
    'Solo para Equipos (*)
    Set rs_datos15 = New ADODB.Recordset
    If rs_datos15.State = 1 Then rs_datos15.Close
    'rs_datos15.Open "select * from av_lista_productos where saldo_actual >= 0 order by DescDetalle ", db, adOpenKeyset, adLockReadOnly  'JQA 06/2008
    rs_datos15.Open "select * from av_solicitud_cotiza_venta ", db, adOpenKeyset, adLockReadOnly
    Set ado_datos15.Recordset = rs_datos15
    ado_datos15.Refresh
    
   'wwwwwwwwwwwwwwwwwwww
    'db.Execute "DELETE ao_ventas_cabecera where venta_codigo = 0 "
    'Call ABREVENTAS
  
'    Set rs_Dsctos = New ADODB.Recordset
'    If rs_Dsctos.State = 1 Then rs_Dsctos.Close
'    rs_Dsctos.Open "select * from ac_ventas_descuentos ", db, adOpenKeyset, adLockReadOnly     'where venta_codigo = '" & TxtNroVenta.Text & "'
'    Set AdoDsctos.Recordset = rs_Dsctos
'    AdoDsctos.Refresh

    Set rs_datos17 = New ADODB.Recordset
    If rs_datos17.State = 1 Then rs_datos17.Close
    rs_datos17.Open "select * from ac_bienes_grupo", db, adOpenKeyset, adLockReadOnly
    Set ado_datos17.Recordset = rs_datos17
    ado_datos17.Refresh
       
    Set rs_datos20 = New ADODB.Recordset
    If rs_datos20.State = 1 Then rs_datos20.Close
    rs_datos20.Open "Select * from fc_cuenta_bancaria", db, adOpenStatic
    Set Ado_datos20.Recordset = rs_datos20
'    dtc_ctades.BoundText = dtc_cta.BoundText
    
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from fc_cuenta_bancaria", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
'    dtc_desc7.BoundText = dtc_codigo7.BoundText

End Sub

Private Sub Form_Unload(Cancel As Integer)
'  If glPersNew = "P" Then
'    frmmo_formulario_M1.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre = rs_Personal!pers_nombres
'    frmmo_formulario_M1.Dtc_Pers_Cargo = rs_Personal!cargo_codigo
'  End If
'  If glPersNew = "L" Then
'    frmmo_formulario_M1.Dtc_doc_id_lab = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell_lab = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2apell_lab = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre_lab = rs_Personal!pers_nombres
'  End If
'  If glPersNew = "PL" Then
'    frmeo_Larvas_mosquitos.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmeo_Larvas_mosquitos.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmeo_Larvas_mosquitos.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmeo_Larvas_mosquitos.Dtc_Pers_nombre = rs_Personal!pers_nombres
'  End If
'  If glPersNew = "PMA" Then
'    frmeo_mosquito_adulto.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmeo_mosquito_adulto.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmeo_mosquito_adulto.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmeo_mosquito_adulto.Dtc_Pers_nombre = rs_Personal!pers_nombres
'  End If
'  glPersNew = "N"

End Sub

Private Sub OpMod1_Click()
'    Txt_modelo.Text = Txt_modelo1.Text
'    Set rs_datos18 = New ADODB.Recordset
'    If rs_datos18.State = 1 Then rs_datos18.Close
'    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    If rs_datos18.RecordCount > 0 Then
'        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs
'    End If
'    'Set ado_datos17.Recordset = rs_datos18
'    'ado_datos17.Refresh
End Sub

Private Sub OpMod2_Click()
'    Txt_modelo.Text = Txt_modelo2.Text
'    Set rs_datos18 = New ADODB.Recordset
'    If rs_datos18.State = 1 Then rs_datos18.Close
'    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    If rs_datos18.RecordCount > 0 Then
'        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs_h
'    End If
End Sub

Private Sub OpMod3_Click()
'    Txt_modelo.Text = Txt_modelo3.Text
'    Set rs_datos18 = New ADODB.Recordset
'    If rs_datos18.State = 1 Then rs_datos18.Close
'    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    If rs_datos18.RecordCount > 0 Then
'        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs_x
'    End If
End Sub

Private Sub OptFilGral01_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    'Set Ado_datos9.Recordset = rs_datos9
    'dtc_desc1.BoundText = dtc_codigo1.BoundText
    Set rs_datos01 = New Recordset
    If rs_datos01.State = 1 Then rs_datos01.Close
    If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "RCUELA" Or glusuario = "CSALINAS" Then
        BtnAprobar1.Visible = True
        queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR' and estado_codigo_fac1 = 'REG' and  DOC_codigo_fac = 'R-101'  AND cobranza_nro_factura <> '0' order by cobranza_fecha_fac, cobranza_nro_factura"
    Else
        BtnAprobar1.Visible = False
        queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR' and estado_codigo_fac1 = 'REG' and DOC_codigo_fac = 'R-101'  AND cobranza_nro_factura <> '0' order by cobranza_fecha_fac, cobranza_nro_factura "
    End If
    'queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos01.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos01.Sort = "cobranza_fecha_prog"
    Set Ado_datos01.Recordset = rs_datos01.DataSource
    Set dg_datos1.DataSource = Ado_datos01.Recordset
End Sub

Private Sub OptFilGral02_Click()
'===== Proceso para filtrado general de datos (todos los registros )
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    
    Set rs_datos01 = New Recordset
    If rs_datos01.State = 1 Then rs_datos01.Close
    'queryinicial1 = "select * From av_ventas_cobranza WHERE beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
    If glusuario = "ADMIN" Or glusuario = "CSALINAS" Or glusuario = "RCUELA" Then
        queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG'  "
        'queryinicial = "SELECT ao_ventas_cobranza.*, ao_ventas_cabecera.* FROM ao_ventas_cobranza INNER JOIN ao_ventas_cabecera ON ao_ventas_cobranza.venta_codigo = ao_ventas_cabecera.venta_codigo"
    Else
        If glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "DLAURA" Then
            queryinicial = "select * From av_ventas_cobranza WHERE unidad_codigo = 'DVTA' and estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' "
        Else
            queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' and beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
        End If
    End If
    rs_datos01.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos01.Sort = "cobranza_fecha_prog"
    Set Ado_datos01.Recordset = rs_datos01.DataSource
    Set dg_datos1.DataSource = Ado_datos01.Recordset
End Sub

Private Sub OptFilGral03_Click()
    '===== Proceso para filtrado de datos(registros Pendientes para Cobrar)
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    'Set Ado_datos9.Recordset = rs_datos9
    'dtc_desc1.BoundText = dtc_codigo1.BoundText
    Set rs_datos02 = New Recordset
    If rs_datos02.State = 1 Then rs_datos02.Close
    If glusuario = "ADMIN" Or glusuario = "RCUELA" Or glusuario = "CSALINAS" Then
        queryinicial2 = "select * From av_ventas_cobranza where estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG'  "
    Else
        If glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "DLAURA" Then
            'queryinicial2 = "select * From av_ventas_cobranza WHERE unidad_codigo = 'DVTA' "
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' AND unidad_codigo = 'DVTA'  "
        Else
            'queryinicial2 = "select * From av_ventas_cobranza WHERE beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
        End If
    End If
        'queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
'    End If
    rs_datos02.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
    rs_datos02.Sort = "cobranza_fecha_fac"
    Set Ado_datos02.Recordset = rs_datos02.DataSource
    Set dg_datos2.DataSource = Ado_datos02.Recordset
End Sub

Private Sub OptFilGral04_Click()
    '===== Proceso para filtrado general de datos(Todos los registros)
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    'Set Ado_datos9.Recordset = rs_datos9
    'dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos02 = New Recordset
    If rs_datos02.State = 1 Then rs_datos02.Close
    If glusuario = "ASANTIVA�EZ" Or glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "DLAURA" Then
        queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR'  "
    Else
        If glusuario = "ADMIN" Or glusuario = "CSALINAS" Or glusuario = "RCUELA" Then
            queryinicial2 = "select * From av_ventas_cobranza estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG'  "
        Else
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
        End If
    End If
    rs_datos02.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
    rs_datos02.Sort = "cobranza_fecha_fac"
    Set Ado_datos02.Recordset = rs_datos02.DataSource
    Set dg_datos2.DataSource = Ado_datos02.Recordset
End Sub

Private Sub OptFilGral05_Click()
'===== Proceso para filtrado de datos(registros Pendientes para Cobrar)
    Set rs_datos02 = New Recordset
    If rs_datos02.State = 1 Then rs_datos02.Close
        'queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG'  "
    If glusuario = "RCUELA" Or glusuario = "ADMIN" Or glusuario = "CSALINAS" Then
        queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG' "
    Else
        If glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "DLAURA" Then
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG' AND unidad_codigo = 'DVTA' "
        Else
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
        End If
    End If
    rs_datos02.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
    rs_datos02.Sort = "cobranza_fecha_cobro"
    Set Ado_datos02.Recordset = rs_datos02.DataSource
    Set dg_datos2.DataSource = Ado_datos02.Recordset

End Sub

Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    'Set Ado_datos9.Recordset = rs_datos9
    'dtc_desc1.BoundText = dtc_codigo1.BoundText
        Set rs_datos = New Recordset
        If rs_datos.State = 1 Then rs_datos.Close
        If glusuario = "RCUELA" Or glusuario = "ADMIN" Or glusuario = "CSALINAS" Then
            queryinicial1 = "select * From av_ventas_cobranza_nd WHERE estado_codigo_fac = 'APR' AND estado_codigo_fac1 = 'APR' and doc_codigo_fac <> 'R-103' "      'ORDER BY cobranza_fecha_prog
        Else
            queryinicial1 = "select * From av_ventas_cobranza_nd WHERE estado_codigo_fac = 'APR' AND estado_codigo_fac1 = 'REG' and doc_codigo_fac = 'R-103' "      'ORDER BY cobranza_fecha_prog
        End If
        rs_datos.Open queryinicial1, db, adOpenKeyset, adLockOptimistic
        rs_datos.Sort = "cobranza_fecha_sol"
        Set Ado_datos.Recordset = rs_datos.DataSource
        Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    If glusuario = "RCUELA" Then
        queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' and doc_codigo_fac <> 'R-103' "     'ORDER BY cobranza_fecha_prog
    Else
        If glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "DLAURA" Then
                queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR'  AND estado_codigo_bco = 'REG' and doc_codigo_fac = 'R-103' "      'ORDER BY cobranza_fecha_prog
            Else
                If glusuario = "ADMIN" Or glusuario = "CSALINAS" Then
                    queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' "      'ORDER BY cobranza_fecha_prog
                End If
            End If
    '    queryinicial = "select * From av_ventas_cobranza WHERE beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
    End If
    'queryinicial = "select * From ao_ventas_cobranza  ORDER BY cobranza_fecha_prog "
    rs_datos.Open queryinicial1, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "cobranza_fecha_sol"
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
  txtTDC.Text = GlTipoCambioMercado ' GlTipoCambioOficial
  
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
  rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot from ao_ventas_detalle where ges_gestion = '" & ges & "' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  If IsNull(rstacumdet!totbs) Then
    VAR_AUX = 0
    VAR_AUX2 = 0
    VAR_CANT = 1
  Else
    VAR_AUX = Round(rstacumdet!totbs, 2)
    VAR_AUX2 = Round(rstacumdet!totdl, 2)
    VAR_CANT = rstacumdet!CANTOT
  End If
  
  rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza where ges_gestion = '" & ges & "' and estado_codigo = 'APR' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  If IsNull(rs_datos19!totbs2) Then
    Cobrobs = 0
    VAR_COBR = 0
  Else
    Cobrobs = Round(rs_datos19!totbs2, 2)
    VAR_COBR = Round(rs_datos19!totdl2, 2)
  End If
  
  VAR_Bs = VAR_AUX - Cobrobs
  VAR_Dol = VAR_AUX2 - VAR_COBR
  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & VAR_AUX & " , ao_ventas_cabecera.venta_monto_total_dol = " & VAR_AUX2 & ", ao_ventas_cabecera.venta_cantidad_total = " & VAR_CANT & ", ao_ventas_cabecera.venta_monto_cobrado_bs = " & Cobrobs & ", ao_ventas_cabecera.venta_monto_cobrado_dol = " & VAR_COBR & ",  ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_Bs & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_Dol & "  Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & Nro & " "
  
'  TxtMontoBs.Text = VAR_AUX
'  TxtCobrado.Text = Cobrobs
'  TxtBstotal.Text = VAR_Bs
  
  If rstacumdet.State = 1 Then rstacumdet.Close
  
End Sub

Private Sub sstab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
        Case 0
            lbl_titulo1.Caption = SSTab1.Caption
'            lbl_titulo3.Caption = SSTab1.Caption
            FraNavega1.Caption = SSTab1.Caption
            FraGrabarCancelar1.Visible = False
            OptFilGral01.Value = True
            Call OptFilGral01_Click
        Case 1
            If glusuario = "RCUELA" Or glusuario = "ADMIN" Or glusuario = "CSALINAS" Then
                lbl_titulo = SSTab1.Caption
'                lbl_titulo2 = SSTab1.Caption
                FraNavega.Caption = SSTab1.Caption
                FraGrabarCancelar.Visible = False
                Call ABRIR_TABLAS_AUX
                OptFilGral1.Value = True
                Call OptFilGral1_Click
                'FACTURA O RECIBO
            Else
                SSTab1.Tab = 0
            End If
            Picture1.Visible = True
'        Case 2
'            lbl_titulo2 = SSTab1.Caption
''            lbl_titulo5 = SSTab1.Caption
'            FraNavega2.Caption = SSTab1.Caption
'            FraGrabarCancelar2.Visible = False
'            OptFilGral03.Value = True
'            Call OptFilGral03_Click
    End Select
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtCantidad_LostFocus()
  If (TxtCantidad.Text) = "" Then
    TxtCantidad.Text = 1
  End If
  If dtc_codigo11.Text = "E" Then
    If (dtc_codigo12.Text) = "" Or IsNull(dtc_codigo12.Text) Then
        TxtDescuento.Text = "0"
    Else
        TxtDescuento.Text = CDbl(TxtCantidad.Text) * (CDbl(TxtPrecioU.Text) * CDbl(Dtc_aux12.Text))
    End If
    'TxtPrecioU.Text = dtc_precioventabase15.Text
    'TxtTotal.Text = CDbl(TxtCantidad.Text) * (CDbl(TxtPrecioU.Text) - CDbl(TxtDescuento.Text))
  End If
  If dtc_codigo11.Text = "C" Then
     TxtDescuento.Text = "0"
     'TxtDescuento.Text = CDbl(Dtc_aux12) * (CDbl(TxtCantidad) * CDbl(TxtPrecioU))
     TxtPrecioU.Text = dtc_precioventafinal15.Text
  End If
  If (dtc_codigo11.Text <> "E" And dtc_codigo11.Text <> "C") Then
     TxtDescuento.Text = "0"
     TxtPrecioU.Text = "0"
  End If
  TxtTotal.Text = (CDbl(TxtCantidad.Text) * CDbl(TxtPrecioU.Text)) - CDbl(TxtDescuento.Text)
  
End Sub

Private Sub TxtCobrado_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtDscto2_LostFocus()
    TxtDscto2D.Text = Round(CDbl(TxtDscto2.Text) / Ado_datos02.Recordset!cobranza_tdc, 2)
End Sub

Private Sub TxtDscto2D_LostFocus()
    TxtDscto2.Text = Round(CDbl(TxtDscto2D.Text) * Ado_datos02.Recordset!cobranza_tdc, 2)
End Sub

Private Sub TxtMonto_LostFocus()
    If TxtMonto.Text = "" Or TxtMonto.Text = "0" Or TxtMonto.Text = "0.00" Then
        TxtMontoDol = "0"
    Else
        'TxtMontoDol = Round(CDbl(TxtMonto.Text) / GlTipoCambioMercado, 2)
        TxtMontoDol = Round(CDbl(TxtMonto.Text) / CDbl(Txt_tdc), 2)
    End If
End Sub

Private Sub TxtPlazo_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]" Or KeyAscii = 8, KeyAscii, 0)
End Sub
'adelante
Private Function CodigoControl(NAuto As String, NFactura As String, Nit As String, Fecha As String, Monto As String, Key As String) As String
Dim Suma As Currency
Dim CodControl As String, Cadena As String, NroVer As String
Dim Pos As Integer, i As Integer, Nro As Integer, j As Integer
Dim SumTot As Long, SumPar(1 To 5) As Currency

  Suma = 0
  Cadena = NFactura
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  NFactura = Cadena
  Suma = Suma + CDbl(Cadena)
  'MsgBox NFactura
  'Para el Nit o CI del Cliente.
  Cadena = Nit
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  Nit = Cadena
  Suma = Suma + CDbl(Cadena)
  'MsgBox Nit
  'Para la Fecha de transaccion.
  Cadena = Fecha
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  Fecha = Cadena
  Suma = Suma + CDbl(Cadena)
  'MsgBox Fecha
  'Para el monto de transaccion.
  Cadena = Monto
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  Monto = Cadena
  'MsgBox Monto
  Suma = Suma + CDbl(Cadena)
  'MsgBox Suma
  
  'Para Obtener los 5 numeros Verhoeff.
  Cadena = Str(Suma)
  For i = 1 To 5
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  NroVer = Right(Cadena, 5)
  'MsgBox NroVer
  
  'Para obtener las nuevas cadenas.
  Cadena = ""
  Pos = 1
  For i = 1 To 5
    Nro = (Val(Mid(NroVer, i, 1)) + 1)
    Select Case i
      Case 1: Cadena = Cadena & NAuto & Mid(Key, Pos, Nro)
      Case 2: Cadena = Cadena & NFactura & Mid(Key, Pos, Nro)
      Case 3: Cadena = Cadena & Nit & Mid(Key, Pos, Nro)
      Case 4: Cadena = Cadena & Fecha & Mid(Key, Pos, Nro)
      Case 5: Cadena = Cadena & Monto & Mid(Key, Pos, Nro)
    End Select
    Pos = Pos + Nro
  Next i

  Cadena = AllegedRC4(Cadena, (Key & NroVer))

  
  SumTot = 0
  i = 0
  Do While i < Len(Cadena)
    i = i + 1
    SumTot = SumTot + Asc(Mid(Trim(Cadena), i, 1))
  Loop
 
  
  For i = 1 To 5
    SumPar(i) = 0
    j = i
    Do While j <= Len(Cadena)
      SumPar(i) = SumPar(i) + Asc(Mid(Cadena, j, 1))
      j = j + 5
    Loop
  
  Next i
  
  Suma = 0
  For i = 1 To 5
    SumPar(i) = Int((SumTot * SumPar(i)) / (Val(Mid(NroVer, i, 1)) + 1))
    Suma = Suma + SumPar(i)
  Next i
  Cadena = Base64(Str(Suma))
  
  Cadena = AllegedRC4(Cadena, (Key & NroVer))
  

  CodigoControl = ""
  i = 0
  j = 1
  
  Do While i < Len(Cadena)
    i = i + 1
    If i Mod 2 = 0 Then
      CodigoControl = CodigoControl & Mid(Cadena, j, 2) & "-"
      j = i + 1
    End If
  Loop
  
  CodigoControl = Mid(CodigoControl, 1, (Len(CodigoControl) - 1))
End Function
Public Function Redondear(dNumero As Double, iDecimales As Integer) As Double
    Dim lMultiplicador As Long
    Dim dRetorno As Double
    
    If iDecimales > 9 Then iDecimales = 9
    lMultiplicador = 10 ^ iDecimales
    dRetorno = CDbl(CLng(dNumero * lMultiplicador)) / lMultiplicador
    
    Redondear = dRetorno
End Function
Private Function Redondeo(ByVal Numero, ByVal Decimales)
      Redondeo = Int(Numero * 10 ^ Decimales + 1 / 2) / 10 ^ Decimales
End Function

Private Sub TxtMonto02_LostFocus()
    TxtMonto02D.Text = Round(CDbl(TxtMonto02.Text) / Ado_datos02.Recordset!cobranza_tdc, 2)
End Sub

Private Sub TxtMonto02D_LostFocus()
    TxtMonto02.Text = Round(CDbl(TxtMonto02D.Text) * Ado_datos02.Recordset!cobranza_tdc, 2)
End Sub

Private Sub TxtMontoDol_Change()
    'TxtMonto.Text = CDbl(TxtMontoDol.Text) * CDbl(txt_tdc.Text)
End Sub

Private Sub TxtMontoDol_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtMontoDol_LostFocus()
    TxtMonto.Text = CDbl(TxtMontoDol.Text) * CDbl(Txt_tdc.Text)
End Sub

Private Sub BtnImprimir5_Click()
    'IMPRIMIR FACTURA con QR
  'RE-IMPRIME FACTURA
  If Ado_datos01.Recordset.RecordCount > 0 And (dtc_aux5.Text <> "") Then
        gestion0 = Ado_datos01.Recordset!ges_gestion
        nroventa = Ado_datos01.Recordset!venta_codigo
        NRO_COBR = Ado_datos01.Recordset!cobranza_codigo
        VAR_COD1 = Ado_datos01.Recordset!cobranza_nro_factura
'        'Dim Exel As Object
'        'Set Exel = CreateObject("Excel.Application")
'        'Exel.Workbooks.Open "c:\tmp\Factura.xlt", , , , "123", "123"
'        'Exel.Visible = True
'        Call CmdFoto_Click
'        ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"
'
'        Picture2.AutoRedraw = True
'        Picture2.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
'
'        ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"
'        ' MsgBox CadenaQr
'        FastQRCode CadenaQr, ImagenQr
'        Picture1.AutoRedraw = True
'        Picture1.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
'        Clipboard.Clear
'        Clipboard.SetData Picture2.Image
'    '    Exel.Application.Range("a2").Select
'    '    Exel.Application.ActiveSheet.Paste
    
        Dim iResult As Variant  ', i%, y%
        sino = MsgBox("Imprimir� con el detalle de Bienes ? ", vbYesNo, "Confirmando")
        If sino = vbYes Then
            CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_rep.rpt"
        Else
            CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura.rpt"
        End If
        CryQ01.WindowShowRefreshBtn = True
        CryQ01.StoredProcParam(0) = gestion0       'Me.Ado_datos.Recordset!ges_gestion
        CryQ01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
        CryQ01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
        var_literal = Ado_datos01.Recordset!Literal
        CryQ01.Formulas(1) = "literalcobro = '" & var_literal & "' "
        CryQ01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
        iResult = CryQ01.PrintReport
        If iResult <> 0 Then MsgBox CryQ01.LastErrorNumber & " : " & CryQ01.LastErrorString, vbCritical, "Error de impresi�n"
  Else
      MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
  End If
End Sub

'Private Sub BtnImprimir5_Click()
'  'RE-IMPRIME FACTURA
'  If Ado_datos01.Recordset.RecordCount > 0 And (dtc_aux5.Text <> "") Then
'        gestion0 = Ado_datos01.Recordset!ges_gestion
'        nroventa = Ado_datos01.Recordset!venta_codigo
'        NRO_COBR = Ado_datos01.Recordset!cobranza_codigo
'
'        Dim iResult As Variant  ', i%, y%
'        sino = MsgBox("Imprimir� con el detalle de Bienes ? ", vbYesNo, "Confirmando")
'        If sino = vbYes Then
'            CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_anterior_rep.rpt"
'        Else
'            CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_anterior.rpt"
'        End If
'        CryF01.WindowShowRefreshBtn = True
'        CryF01.StoredProcParam(0) = gestion0       'Me.Ado_datos.Recordset!ges_gestion
'        CryF01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
'        CryF01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
'        var_literal = Ado_datos01.Recordset!Literal
'        CryF01.Formulas(1) = "literalcobro = '" & var_literal & "' "
'        CryF01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
'        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
'        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'        iResult = CryF01.PrintReport
'        If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresi�n"
'  Else
'      MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
'  End If
'End Sub

'Private Sub BtnImprimir5_Click()
''RE-IMPRIME FACTURA
'  If Ado_datos02.Recordset.RecordCount > 0 And (dtc_aux5.Text <> "") Then
'    If (Ado_datos02.Recordset!factura_impresa = "N") And (Ado_datos02.Recordset!cobranza_deuda_bs <> "0.00") Then
'      If Ado_datos02.Recordset!doc_codigo_fac = "R-101" Then
'        '===== ini GENERA EL CODIGO DE FACTURA ====
'        Set rs_aux1 = New ADODB.Recordset
'        rs_aux1.CursorLocation = adUseClient
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "select * from fc_dosificacion_docs where doc_codigo = 'R-101' AND estado_codigo = 'APR' ", db, adOpenDynamic, adLockOptimistic
'        'rs_aux1.Open "select * from fc_dosificacion_docs  where doc_codigo = 'R-101'  ", db, adOpenDynamic, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
'            gestion0 = glGestion        'Ado_datos02.Recordset("ges_gestion")
'            correlv = Ado_datos02.Recordset("venta_codigo")
'            nroventa = Ado_datos02.Recordset("venta_codigo")
'            NRO_COBR = Me.Ado_datos02.Recordset!cobranza_codigo
'            VAR_BENEF = Ado_datos02.Recordset!beneficiario_codigo
'            VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
'            'VAR_GLOSA = Ado_datos02.Recordset!cobranza_observaciones
'            VAR_GLOSA = Trim(Ado_datos02.Recordset!cobranza_observaciones) + " - Tram.: " + Trim(VAR_CITE)
'            'VAR_DOL2 = Round(Ado_datos02.Recordset!cobranza_deuda_dol, 2)
'            'VAR_BS2 = Round(Ado_datos02.Recordset!cobranza_deuda_bs, 2)
'            VAR_DOL2 = Round(Ado_datos02.Recordset!cobranza_total_dol, 2)
'            VAR_BS2 = Round(Ado_datos02.Recordset!cobranza_total_bs, 2)
'            'VAR_CTA = IIf(Ado_datos02.Recordset!Cta_Codigo = "", "NN", Ado_datos02.Recordset!Cta_Codigo)
'            var_literal = Ado_datos02.Recordset!Literal
'            VAR_FFAC = Format((Date), "DD/MM/YYYY")
'            VAR_CODTIPO = "REF"     'Tipo Comprobante (paralelo VAR_DOC)
'            VAR_DOC = "R-112"       'Doc. Respaldo
'            VAR_ETAPA = "FIN-01-02"
'            VAR_TCOMP = "RECAUDADO (FACTURACION)"
'            Llave = Trim(rs_aux1!dosifica_llave)
'            If dtc_aux5.Text Like " " Then
'                MsgBox "Error en el NIT del Cliente, Contactese con el Administrador y vuelva a intentar ...", , "Atenci�n"
'                Exit Sub
'            Else
'                NitCi = IIf(dtc_aux5.Text = "", Ado_datos02.Recordset!beneficiario_codigo_fac, dtc_aux5.Text)    'VAR_BENEF
'            End If
'            Autorizacion = rs_aux1!dosifica_autorizacion
'            'Fecha = Val(Format((Date), "YYYYMMDD"))
'            'Monto = Redondeo((VAR_BS2), 0)
'            'CodigoContro = CodigoControl(Autorizacion, NroFactura, NitCi, Fecha, Monto, Llave)
'            VAR_PROY2 = Ado_datos16.Recordset!edif_codigo
'            VAR_COD4 = Ado_datos16.Recordset!unidad_codigo
'            VAR_TIPOV = Ado_datos16.Recordset!venta_tipo
'            VAR_SOL = Ado_datos16.Recordset!solicitud_codigo
'            VAR_MONEDA = Ado_datos02.Recordset!tipo_moneda
'            'CodigoContro = CodigoControl(NroFactura)
'            If Autorizacion <> "" And NitCi <> "" And Llave <> "" And VAR_BS2 <> "0" And rs_aux1!CORREL >= 0 Then
'                VAR_SW = 1
'            Else
'                VAR_SW = 0
'                MsgBox "Error en Autorizacion, NIT o Llave, Contactese con el Administrador y vuelva a intentar ...", , "Atenci�n"
'                Exit Sub
'            End If
'            VAR_COD1 = CDbl(rs_aux1!CORREL) + 1
'            sino = MsgBox("Esta seguro(a) de IMPRIMIR la Factura Nro. " + Str(VAR_COD1) + " ?", vbYesNo, "Confirmando")
'            If sino = vbYes Then
'                rs_aux1!CORREL = Trim(Str(VAR_COD1))
'                rs_aux1.Update
'                'GENERA CORREL NOTA DEBITO POR DEPTO INI
'                Set rs_aux5 = New ADODB.Recordset
'                If rs_aux5.State = 1 Then rs_aux5.Close
'                'rs_aux5.Open "Select correl_contab as Codigo from gc_departamento where depto_codigo = '" & Left(VAR_PROY3, 1) & "'    ", db, adOpenStatic
'                rs_aux5.Open "Select * from fc_correl where tipo_tramite  = 'NDEBITO '    ", db, adOpenStatic
'                If Not rs_aux5.EOF Then
'                    VAR_CONTAB = IIf(IsNull(rs_aux5!numero_correlativo), 1, CDbl(rs_aux5!numero_correlativo) + 1)
'                End If
'                'rs_aux5!Codigo = VAR_CONTAB
'                'rs_aux5.Update
'
'                VAR_COD2 = rs_aux1!dosifica_autorizacion
'                NroFactura = Trim(Str(VAR_COD1))
'                Fecha = Val(Format((Date), "YYYYMMDD"))
'                Monto = Redondeo((VAR_BS2), 0)
'
'                CodigoContro = CodigoControl(Autorizacion, NroFactura, NitCi, Fecha, Monto, Llave)
'                If CodigoContro = "" Or CodigoContro = "0" Then
'                    VAR_SW = 0
'                    MsgBox "Error en Codigo de Control, Contactese con el Administrador o vuelva a intentar ...", , "Atenci�n"
'                    Exit Sub
'                Else
'                    VAR_SW = 1
'                End If
'                db.Execute "update ao_ventas_cobranza set correl_contab = " & VAR_CONTAB & " Where ao_ventas_cobranza.venta_codigo = " & Ado_datos02.Recordset!venta_codigo & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                db.Execute "update fc_correl set numero_correlativo = " & VAR_CONTAB & " Where tipo_tramite = 'NDEBITO' "
'                'Ado_datos02.Recordset!correl_contab = VAR_CONTAB
'                If VAR_CONTAB < 10 Then
'                    'Ado_datos02.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
'                    VAR_GLOSA = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
'                End If
'                If VAR_CONTAB > 9 And VAR_CONTAB < 100 Then
'                   'Ado_datos02.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
'                   VAR_GLOSA = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
'                End If
'                If VAR_CONTAB > 99 Then
''                    If VAR_CONTAB > 1200 Then
''                        MsgBox "El ND Finaliza en 6564 ... ", , "Atenci�n"
''                    End If
'                   'Ado_datos02.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
'                   VAR_GLOSA = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
'                End If
'                db.Execute "update ao_ventas_cobranza set cobranza_observaciones = '" & VAR_GLOSA & "' Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
''               'GENERA CORREL NOTA DEBITO POR DEPTO FIN
'
'                '===== ini nombre archivo de la FACTURA ====
'                'db.Execute "update ao_ventas_cobranza set archivo_foto = '" & doc_codigo & "' + '-' + '" & Str(VAR_COD1) & "' + '.JPG' Where venta_codigo = " & Ado_datos02.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set archivo_foto = 'R101-' + '" & Str(VAR_COD1) & "' + '.JPG' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set archivo_foto_cargado = 'N' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                '===== fin nombre archivo de la FACTURA ====
'                ' ACTUALIZA NRO FAC. EN ao_ventas_cobranza
'                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac = '" & Date & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set cobranza_nro_factura = " & VAR_COD1 & " Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion = " & VAR_COD2 & " Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                'IMPRIMIR FACTURA
''                VAR_ANIO = CStr(glGestion)
''                VAR_MES = CStr(Month(Date))
''                VAR_DIA = CStr(Day(Date))
''                VAR_FECHA = VAR_ANIO & VAR_MES & VAR_DIA
'                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac2 = '" & Fecha & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & NRO_COBR & " "
'                'Dim F1
'                'FI = Ado_datos02.Recordset!cobranza_fecha_cobro
'                'frm_qr.txt_texto = GlParametro + "|" + GlParametroDes + "|" + Trim(str(VAR_COD1)) + "|" + TrimSTR((VAR_COD2)) + "|" + '" & Ado_datos02.Recordset!cobranza_fecha_cobro & "' + "|" + " & Ado_datos02.Recordset!cobranza_deuda_bs & " + "|" + '" & rs_aux1!dosifica_codigo_control & "' + "|" + '" & rs_aux1!dosifica_fecha_limite & "' + "|" + "0" + "|" + "0" + "|" + '" & Ado_datos02.Recordset!beneficiario_codigo & "' + "|" + '" & dtc_desc2A.Text & "'
'                'frm_qr.Show vbModal
'                'NIT del emisor, Nombre o Raz�n Social del emisor, N�mero correlativo de Factura, N�mero de Autorizaci�n, Fecha de emisi�n, Importe de la compra, C�digo de Control, Fecha L�mite de Emisi�n, 0, 0, NIT / NDI Comprador, Nombre o Raz�n Social del comprador
'
'                'MsgBox "Se est� Imprimiendo la Factura Nro. " + Str(VAR_COD1), , "Atenci�n"
'                db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where venta_codigo = " & Ado_datos02.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'APR' Where cobranza_codigo = " & Ado_datos02.Recordset("cobranza_codigo") & " "
'                db.Execute "update ao_ventas_cobranza set estado_codigo_bco = 'REG' Where cobranza_codigo = " & Ado_datos02.Recordset("cobranza_codigo") & " "
'
'                db.Execute "update ao_ventas_cobranza set cobranza_codigo_control = '" & CodigoContro & "' Where cobranza_codigo = " & Ado_datos02.Recordset("cobranza_codigo") & " "
'
'                sFile = App.Path & "\CLIENTES\QRCode.bmp"
'                CadenaQ = Trim("1018533029") _
'                & "|" & Trim(VAR_COD1) _
'                & "|" & Trim(VAR_COD2) _
'                & "|" & Format(Trim(Date), "DD/MM/YYYY") _
'                & "|" & Format(Trim(VAR_BS2), "###0.00") _
'                & "|" & Format(Trim(VAR_BS2), "###0.00") _
'                & "|" & Trim(CodigoContro) _
'                & "|" & Trim(dtc_aux5.Text) _
'                & "|" & Trim("0") _
'                & "|" & Trim("0") _
'                & "|" & Trim("0") _
'                & "|" & Trim("0")
'
'                FastQRCode CadenaQ, sFile
'                Set Picture1.Picture = LoadPicture(sFile)
'                'FIN QR
'                'Call IMPRIME_FACTURA
'                Call IMPRIME_QR
'                'MsgBox CadenaQ
'                'If VAR_TIPOV = "C" Then
'                    Call Contabiliza_venta
'                'End If
'            Else
'                VAR_COD1 = "0"
'                If rs_aux1.State = 1 Then rs_aux1.Close
'                Exit Sub
'            End If
'        End If
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        '===== fin TERMINA GENERACION DE FACTURA =====
'
'
'        TxtCmpbte = VAR_COD1
'        If (Ado_datos02.Recordset("estado_codigo_sol") = "APR" And Ado_datos02.Recordset("estado_codigo_fac") = "REG") Then          'REG
'          Call OptFilGral1_Click
'        Else
'          Call OptFilGral2_Click
'        End If
'      Else
'        Call generarRepRecibo
'      End If
'      If Ado_datos02.Recordset!doc_codigo_fac = "R-103" Then
'      'WWWWWWWWWWWWWWWWWWWWWWWWW
'        '===== ini GENERA EL CODIGO DE RECIBO ====
'        Set rs_aux1 = New ADODB.Recordset
'        rs_aux1.CursorLocation = adUseClient
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "select * from gc_documentos_respaldo where doc_codigo = 'R-103' AND estado_codigo = 'APR' ", db, adOpenDynamic, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
'            gestion0 = glGestion        'Ado_datos02.Recordset("ges_gestion")
'            correlv = Ado_datos02.Recordset("venta_codigo")
'            nroventa = Ado_datos02.Recordset("venta_codigo")
'            NRO_COBR = Me.Ado_datos02.Recordset!cobranza_codigo
'            VAR_BENEF = Ado_datos02.Recordset!beneficiario_codigo
'            VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
'            'VAR_GLOSA = Ado_datos02.Recordset!cobranza_observaciones
'            VAR_GLOSA = Trim(Ado_datos02.Recordset!cobranza_observaciones) + " - Tram.: " + Trim(VAR_CITE)
'            VAR_DOL2 = Round(Ado_datos02.Recordset!cobranza_deuda_dol, 2)
'            VAR_BS2 = Round(Ado_datos02.Recordset!cobranza_deuda_bs, 2)
'            'VAR_CTA = IIf(Ado_datos02.Recordset!Cta_Codigo = "", "NN", Ado_datos02.Recordset!Cta_Codigo)
'            var_literal = Ado_datos02.Recordset!Literal
'            'Llave = Trim(rs_aux1!dosifica_llave)
'            NitCi = IIf(dtc_aux5.Text = "", Ado_datos02.Recordset!beneficiario_codigo_fac, dtc_aux5.Text)    'VAR_BENEF
'            'Autorizacion = rs_aux1!dosifica_autorizacion
'            VAR_PROY2 = Ado_datos16.Recordset!edif_codigo
'            VAR_COD4 = Ado_datos16.Recordset!unidad_codigo
'            VAR_TIPOV = Ado_datos16.Recordset!venta_tipo
'            VAR_SOL = Ado_datos16.Recordset!solicitud_codigo
'            VAR_MONEDA = Ado_datos02.Recordset!tipo_moneda
'
'            VAR_COD1 = CDbl(rs_aux1!correl_doc) + 1
'            sino = MsgBox("Esta seguro(a) de IMPRIMIR la Recibo Nro. " + Str(VAR_COD1) + " ?", vbYesNo, "Confirmando")
'            If sino = vbYes Then
'                rs_aux1!correl_doc = Trim(Str(VAR_COD1))
'                rs_aux1.Update
'                'GENERA CORREL NOTA DEBITO POR DEPTO INI
'                VAR_GLOSA = TxtObs.Text
'                db.Execute "update ao_ventas_cobranza set cobranza_observaciones = '" & VAR_GLOSA & "' Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
''                'GENERA CORREL NOTA DEBITO POR DEPTO FIN
'
'                VAR_COD2 = "0"  'rs_aux1!dosifica_autorizacion
'                NroFactura = Trim(Str(VAR_COD1))
'                '===== ini nombre archivo de la FACTURA ====
'                db.Execute "update ao_ventas_cobranza set archivo_foto = 'R103-' + '" & Str(VAR_COD1) & "' + '.JPG' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set archivo_foto_cargado = 'N' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                '===== fin nombre archivo de la FACTURA ====
'                ' ACTUALIZA NRO FAC. EN ao_ventas_cobranza
'                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac = '" & Date & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set cobranza_nro_factura = " & VAR_COD1 & " Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion = " & VAR_COD2 & " Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                'IMPRIMIR FACTURA
'                Fecha = Val(Format((Date), "YYYYMMDD"))
'                Monto = Redondeo((VAR_BS2), 0)
'                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac2 = '" & Fecha & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & NRO_COBR & " "
'                'Dim F1
'                'FI = Ado_datos02.Recordset!cobranza_fecha_cobro
'                'frm_qr.txt_texto = GlParametro + "|" + GlParametroDes + "|" + Trim(str(VAR_COD1)) + "|" + TrimSTR((VAR_COD2)) + "|" + '" & Ado_datos02.Recordset!cobranza_fecha_cobro & "' + "|" + " & Ado_datos02.Recordset!cobranza_deuda_bs & " + "|" + '" & rs_aux1!dosifica_codigo_control & "' + "|" + '" & rs_aux1!dosifica_fecha_limite & "' + "|" + "0" + "|" + "0" + "|" + '" & Ado_datos02.Recordset!beneficiario_codigo & "' + "|" + '" & dtc_desc2A.Text & "'
'                'frm_qr.Show vbModal
'                'NIT del emisor, Nombre o Raz�n Social del emisor, N�mero correlativo de Factura, N�mero de Autorizaci�n, Fecha de emisi�n, Importe de la compra, C�digo de Control, Fecha L�mite de Emisi�n, 0, 0, NIT / NDI Comprador, Nombre o Raz�n Social del comprador
'
'                'MsgBox "Se est� Imprimiendo la Factura Nro. " + Str(VAR_COD1), , "Atenci�n"
'                db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where venta_codigo = " & Ado_datos02.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos02.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'APR' Where cobranza_codigo = " & Ado_datos02.Recordset("cobranza_codigo") & " "
'                db.Execute "update ao_ventas_cobranza set estado_codigo_bco = 'REG' Where cobranza_codigo = " & Ado_datos02.Recordset("cobranza_codigo") & " "
'
'                VAR_SW = 1
'                'CodigoContro = CodigoControl(Autorizacion, NroFactura, NitCi, Fecha, Monto, Llave)
'                'db.Execute "update ao_ventas_cobranza set cobranza_codigo_control = '" & CodigoContro & "' Where cobranza_codigo = " & Ado_datos02.Recordset("cobranza_codigo") & " "
'                Call IMPRIME_RECIBO
'                'If VAR_TIPOV = "C" Then
'                    Call Contabiliza_venta
'                'End If
'            Else
'                VAR_COD1 = "0"
'                If rs_aux1.State = 1 Then rs_aux1.Close
'                Exit Sub
'            End If
'        End If
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        '===== fin TERMINA GENERACION DE FACTURA =====
'        TxtCmpbte = VAR_COD1
'        If (Ado_datos02.Recordset("estado_codigo_sol") = "APR" And Ado_datos02.Recordset("estado_codigo_fac") = "REG") Then          'REG
'          Call OptFilGral1_Click
'        Else
'          Call OptFilGral2_Click
'        End If
'      'WWWWWWWWWWWWWWWWWWWWWWWWW
'      End If
'    Else
'        MsgBox "La Factura Nro. " + Ado_datos02.Recordset!cobranza_nro_factura + " ya fue Impresa", , "Atenci�n"
'        'Call IMPRIME_FACTURA
'        If (Ado_datos02.Recordset("estado_codigo_sol") = "APR" And Ado_datos02.Recordset("estado_codigo_fac") = "REG") Then          'REG
'          Call OptFilGral1_Click
'        Else
'          Call OptFilGral2_Click
'        End If
'    End If
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atenci�n"
'  End If
'
'End Sub



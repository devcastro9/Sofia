VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmfo_FA_formulario 
   Caption         =   "REGISTRO Y CONTROL DE PERSONAL"
   ClientHeight    =   10230
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame fraOpciones 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   20
      TabIndex        =   16
      Top             =   600
      Width           =   6030
      Begin VB.CommandButton CmdVerDisco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Ver Disco"
         Height          =   720
         Left            =   4440
         Picture         =   "frmfo_FA_formulario.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   240
         Width           =   730
      End
      Begin VB.CommandButton CmdDel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Anu&Lar"
         Height          =   720
         Left            =   1560
         Picture         =   "frmfo_FA_formulario.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Anula Registro"
         Top             =   240
         Width           =   730
      End
      Begin VB.CommandButton cmdAprueba 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Aprobar"
         Height          =   720
         Left            =   2280
         Picture         =   "frmfo_FA_formulario.frx":0B2C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Aprueba Registro"
         Top             =   240
         Width           =   730
      End
      Begin VB.CommandButton CmdObs 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Rechazar"
         Height          =   720
         Left            =   1560
         Picture         =   "frmfo_FA_formulario.frx":0D36
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Rechaza Registro"
         Top             =   240
         Width           =   730
      End
      Begin VB.CommandButton CmdCopiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Volver a &Copiar"
         Height          =   720
         Left            =   4440
         Picture         =   "frmfo_FA_formulario.frx":1178
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   240
         Visible         =   0   'False
         Width           =   730
      End
      Begin VB.CommandButton cmdDesaprueba 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2280
         Picture         =   "frmfo_FA_formulario.frx":15BA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   730
      End
      Begin VB.CommandButton CmdBuscar 
         BackColor       =   &H8000000B&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3720
         Picture         =   "frmfo_FA_formulario.frx":17C4
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   730
      End
      Begin VB.CommandButton CmdSal 
         BackColor       =   &H8000000B&
         Caption         =   "Salir"
         Height          =   720
         Left            =   5160
         Picture         =   "frmfo_FA_formulario.frx":208E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Sale del Formulario"
         Top             =   240
         Width           =   730
      End
      Begin VB.CommandButton CmdAdd 
         BackColor       =   &H8000000B&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         MousePointer    =   4  'Icon
         Picture         =   "frmfo_FA_formulario.frx":2298
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Nuevo Registro"
         Top             =   240
         Width           =   730
      End
      Begin VB.CommandButton CmdMod 
         BackColor       =   &H8000000B&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   840
         Picture         =   "frmfo_FA_formulario.frx":25A2
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Edita Registro"
         Top             =   240
         Width           =   730
      End
      Begin VB.CommandButton CmdImprimir 
         BackColor       =   &H8000000B&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   3000
         Picture         =   "frmfo_FA_formulario.frx":27AC
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Imprime Registro"
         Top             =   240
         Width           =   730
      End
      Begin Crystal.CrystalReport Cry_M1 
         Left            =   600
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame FraGrabarCancelar 
      BackColor       =   &H00C0E0FF&
      Height          =   1035
      Left            =   20
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   6030
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reno&var"
         Height          =   540
         Left            =   2640
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdGrabar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Grabar"
         Height          =   720
         Left            =   1560
         Picture         =   "frmfo_FA_formulario.frx":2E96
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   200
         Width           =   770
      End
      Begin VB.CommandButton CmdCancelar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancelar"
         Height          =   720
         Left            =   3720
         Picture         =   "frmfo_FA_formulario.frx":30A0
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   200
         Width           =   770
      End
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   3480
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfo_FA_formulario.frx":32AA
            Key             =   "Raiz"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfo_FA_formulario.frx":35C4
            Key             =   "Cerrado"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfo_FA_formulario.frx":3E9E
            Key             =   "Abierto"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfo_FA_formulario.frx":4778
            Key             =   "Detalle"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trv 
      Height          =   2265
      Left            =   0
      TabIndex        =   39
      Top             =   6360
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   3995
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "iml"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   6045
      TabIndex        =   24
      Top             =   600
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   7
      Tab             =   3
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   -2147483637
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
      TabCaption(0)   =   "1. Inf. General"
      TabPicture(0)   =   "frmfo_FA_formulario.frx":5052
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FraCabecera"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "2. Asistencia"
      TabPicture(1)   =   "frmfo_FA_formulario.frx":506E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command19"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command17"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "FraEmpresa"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "3. Permisos"
      TabPicture(2)   =   "frmfo_FA_formulario.frx":508A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FraProyecto"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdMod3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command18"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command20"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "4. Vacaciones"
      TabPicture(3)   =   "frmfo_FA_formulario.frx":50A6
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command21"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Command22"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Command23"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Command24"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Command25"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Command26"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "5. Memos"
      TabPicture(4)   =   "frmfo_FA_formulario.frx":50C2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame2"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Command2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Command27"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Command28"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "6. Movilidad Pers"
      TabPicture(5)   =   "frmfo_FA_formulario.frx":50DE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label9"
      Tab(5).Control(1)=   "FraDeclarJur"
      Tab(5).Control(2)=   "Frame4"
      Tab(5).Control(3)=   "TxtDireccion2"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Tab 6"
      TabPicture(6)   =   "frmfo_FA_formulario.frx":50FA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ImgMemo"
      Tab(6).Control(1)=   "ImgVacacion"
      Tab(6).Control(2)=   "Image2"
      Tab(6).Control(3)=   "ImgEvaluacion"
      Tab(6).Control(4)=   "DataGrid5"
      Tab(6).Control(5)=   "DataGrid4"
      Tab(6).Control(6)=   "DataGrid3"
      Tab(6).Control(7)=   "DataGrid2"
      Tab(6).ControlCount=   8
      Begin VB.TextBox TxtDireccion2 
         DataField       =   "domicilio_legal"
         DataSource      =   "adoLista"
         Height          =   285
         Left            =   -73440
         TabIndex        =   56
         Top             =   7500
         Width           =   7155
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000B&
         Caption         =   "Lugar (DOMICILIO ACTUAL) ----------------------------- Lugar Desigacion (de Oficina)"
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
         Height          =   1695
         Left            =   -74760
         TabIndex        =   153
         Top             =   5100
         Width           =   8475
         Begin MSDataListLib.DataCombo Dtc_prov_cod2 
            Bindings        =   "frmfo_FA_formulario.frx":5116
            DataField       =   "prov_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   5640
            TabIndex        =   154
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "prov_codigo"
            BoundColumn     =   "prov_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_munic_cod2 
            Bindings        =   "frmfo_FA_formulario.frx":512E
            DataField       =   "munic_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   5640
            TabIndex        =   155
            Top             =   795
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "munic_codigo"
            BoundColumn     =   "munic_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_local_cod2 
            Bindings        =   "frmfo_FA_formulario.frx":5146
            DataField       =   "comun_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   5640
            TabIndex        =   156
            Top             =   1140
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "comun_codigo"
            BoundColumn     =   "comun_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_depto_cod2 
            Bindings        =   "frmfo_FA_formulario.frx":5161
            DataField       =   "depto_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   5640
            TabIndex        =   157
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "depto_codigo"
            BoundColumn     =   "depto_codigo"
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
         Begin MSDataListLib.DataCombo DataCombo9 
            Bindings        =   "frmfo_FA_formulario.frx":517A
            DataField       =   "prov_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   2400
            TabIndex        =   158
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "prov_codigo"
            BoundColumn     =   "prov_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo10 
            Bindings        =   "frmfo_FA_formulario.frx":5191
            DataField       =   "munic_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   2400
            TabIndex        =   159
            Top             =   795
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "munic_codigo"
            BoundColumn     =   "munic_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_local_cod 
            Bindings        =   "frmfo_FA_formulario.frx":51A8
            DataField       =   "comun_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   2400
            TabIndex        =   160
            Top             =   1140
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "comun_codigo"
            BoundColumn     =   "comun_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo11 
            Bindings        =   "frmfo_FA_formulario.frx":51C2
            DataField       =   "prov_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   720
            TabIndex        =   161
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "prov_descripcion"
            BoundColumn     =   "prov_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo12 
            Bindings        =   "frmfo_FA_formulario.frx":51D9
            DataField       =   "munic_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   720
            TabIndex        =   162
            Top             =   960
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "munic_descripcion"
            BoundColumn     =   "munic_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_local 
            Bindings        =   "frmfo_FA_formulario.frx":51F0
            DataField       =   "comun_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   720
            TabIndex        =   163
            Top             =   1320
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "comun_descripcion"
            BoundColumn     =   "comun_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo13 
            Bindings        =   "frmfo_FA_formulario.frx":520A
            DataField       =   "depto_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   2400
            TabIndex        =   164
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483624
            ListField       =   "depto_codigo"
            BoundColumn     =   "depto_codigo"
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
         Begin MSDataListLib.DataCombo DataCombo14 
            Bindings        =   "frmfo_FA_formulario.frx":5222
            DataField       =   "depto_codigo"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   720
            TabIndex        =   165
            Top             =   240
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "depto_descripcion"
            BoundColumn     =   "depto_codigo"
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
         Begin MSDataListLib.DataCombo Dtc_depto2 
            Bindings        =   "frmfo_FA_formulario.frx":523A
            DataField       =   "depto_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   4920
            TabIndex        =   166
            Top             =   240
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "depto_descripcion"
            BoundColumn     =   "depto_codigo"
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
         Begin MSDataListLib.DataCombo Dtc_prov2 
            Bindings        =   "frmfo_FA_formulario.frx":5253
            DataField       =   "prov_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   4920
            TabIndex        =   167
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "prov_descripcion"
            BoundColumn     =   "prov_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_munic2 
            Bindings        =   "frmfo_FA_formulario.frx":526B
            DataField       =   "munic_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   4920
            TabIndex        =   168
            Top             =   960
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "munic_descripcion"
            BoundColumn     =   "munic_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_local2 
            Bindings        =   "frmfo_FA_formulario.frx":5283
            DataField       =   "comun_codigo2"
            DataSource      =   "adoLista"
            Height          =   315
            Left            =   4920
            TabIndex        =   169
            Top             =   1320
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "comun_descripcion"
            BoundColumn     =   "comun_codigo"
            Text            =   ""
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H8000000B&
            Caption         =   "Depto."
            Height          =   255
            Index           =   26
            Left            =   4320
            TabIndex        =   60
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H8000000B&
            Caption         =   "Comuni."
            Height          =   255
            Index           =   25
            Left            =   4320
            TabIndex        =   176
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H8000000B&
            Caption         =   "Munic."
            Height          =   255
            Index           =   24
            Left            =   4320
            TabIndex        =   175
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H8000000B&
            Caption         =   "Prov."
            Height          =   255
            Index           =   23
            Left            =   4320
            TabIndex        =   174
            Top             =   585
            Width           =   615
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H8000000B&
            Caption         =   "Prov."
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   173
            Top             =   645
            Width           =   615
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H8000000B&
            Caption         =   "Munic."
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   172
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H8000000B&
            Caption         =   "Comuni."
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   171
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H8000000B&
            Caption         =   "Depto."
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   170
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.Frame FraDeclarJur 
         BackColor       =   &H8000000B&
         Caption         =   "18. DECLARACION JURADA"
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
         Height          =   4215
         Left            =   -74760
         TabIndex        =   142
         Top             =   900
         Visible         =   0   'False
         Width           =   9135
         Begin VB.CheckBox ChkPromotor 
            BackColor       =   &H8000000B&
            Caption         =   "Tiene la Firma del PROMOTOR"
            ForeColor       =   &H00400040&
            Height          =   255
            Left            =   240
            TabIndex        =   151
            Top             =   240
            Width           =   3255
         End
         Begin VB.CheckBox ChkReprLegal 
            BackColor       =   &H8000000B&
            Caption         =   "Tiene la Firma del Representante Legal de la Institución o Empresa"
            ForeColor       =   &H00400040&
            Height          =   255
            Left            =   240
            TabIndex        =   150
            Top             =   1140
            Width           =   6135
         End
         Begin VB.CheckBox ChkConsultor 
            BackColor       =   &H8000000B&
            Caption         =   "Tiene la Firma del CONSULTOR que Elaboró la Ficha Ambiental"
            ForeColor       =   &H00400040&
            Height          =   255
            Left            =   240
            TabIndex        =   149
            Top             =   660
            Width           =   5415
         End
         Begin VB.CheckBox ChkOtraPers 
            BackColor       =   &H8000000B&
            Caption         =   "Tiene la Firma de otra Persona que se Menciona en el Documento"
            ForeColor       =   &H00400040&
            Height          =   255
            Left            =   240
            TabIndex        =   148
            Top             =   1560
            Width           =   6135
         End
         Begin VB.CheckBox ChkProyPlano 
            BackColor       =   &H8000000B&
            Caption         =   "Tiene Plano de Ubicación del Predio"
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   240
            TabIndex        =   147
            Top             =   2640
            Width           =   3375
         End
         Begin VB.CheckBox ChkProyCertUS 
            BackColor       =   &H8000000B&
            Caption         =   "Tiene Certificado de Uso de Suelo"
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   240
            TabIndex        =   146
            Top             =   3420
            Width           =   3135
         End
         Begin VB.CheckBox ChkProyFoto 
            BackColor       =   &H8000000B&
            Caption         =   "Tiene Fotografias Panorámicas del Lugar"
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   240
            TabIndex        =   145
            Top             =   3780
            Width           =   3495
         End
         Begin VB.CheckBox ChkProyProp 
            BackColor       =   &H8000000B&
            Caption         =   "Tiene Derecho Propietario del Inmueble"
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   240
            TabIndex        =   144
            Top             =   3060
            Width           =   3615
         End
         Begin VB.CheckBox ChkTestimonio 
            BackColor       =   &H8000000B&
            Caption         =   "Tiene Testimonio de Costitución de la Institución o empresa"
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   240
            TabIndex        =   143
            Top             =   2280
            Width           =   4695
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "---------------------------------------------------- DOCUMENTACION DE RESPALDO ----------------------------------------"
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
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   152
            Top             =   1920
            Width           =   8640
         End
      End
      Begin VB.CommandButton Command28 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Modif."
         Enabled         =   0   'False
         Height          =   720
         Left            =   -66360
         Picture         =   "frmfo_FA_formulario.frx":529E
         Style           =   1  'Graphical
         TabIndex        =   141
         ToolTipText     =   "Registro Ubicacion Fisica"
         Top             =   1980
         Width           =   645
      End
      Begin VB.CommandButton Command27 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   -66360
         MousePointer    =   4  'Icon
         Picture         =   "frmfo_FA_formulario.frx":56E0
         Style           =   1  'Graphical
         TabIndex        =   140
         ToolTipText     =   "Nuevo Registro"
         Top             =   1260
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Grabar"
         Height          =   720
         Left            =   -66360
         Picture         =   "frmfo_FA_formulario.frx":59EA
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   2820
         Width           =   645
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Caption         =   "5. REGISTRO DE MEMORANDUMS"
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
         ForeColor       =   &H00008000&
         Height          =   3585
         Left            =   -74880
         TabIndex        =   125
         Top             =   4500
         Width           =   9255
         Begin VB.TextBox Txt_InsumoAlm 
            Alignment       =   2  'Center
            DataField       =   "Insumos_almacenaje"
            DataSource      =   "ADO_M1"
            Height          =   405
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   138
            Text            =   "frmfo_FA_formulario.frx":5BF4
            Top             =   2160
            Width           =   8655
         End
         Begin VB.TextBox Txt_AlternativaOtro 
            Alignment       =   2  'Center
            DataField       =   "Alternativa_localizacion_Desc"
            DataSource      =   "ADO_M1"
            Height          =   405
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   137
            Text            =   "frmfo_FA_formulario.frx":5BF7
            Top             =   1320
            Width           =   8655
         End
         Begin VB.ComboBox Combo13 
            DataField       =   "DDRR_Anio"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":5BFA
            Left            =   7080
            List            =   "frmfo_FA_formulario.frx":5C04
            TabIndex        =   128
            Text            =   "NO"
            Top             =   3000
            Width           =   735
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            DataField       =   "Vida_Util_Proy_Anio"
            DataSource      =   "ADO_M1"
            Height          =   285
            Left            =   3840
            ScrollBars      =   2  'Vertical
            TabIndex        =   127
            Text            =   "0"
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            DataField       =   "Vida_Util_Proy_Mes"
            DataSource      =   "ADO_M1"
            Height          =   285
            Left            =   240
            ScrollBars      =   2  'Vertical
            TabIndex        =   126
            Text            =   "0"
            Top             =   3000
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DTPicker11 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   6720
            TabIndex        =   129
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301697
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker18 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   3720
            TabIndex        =   130
            Top             =   3000
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301697
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo DataCombo7 
            Bindings        =   "frmfo_FA_formulario.frx":5C10
            DataField       =   "Sitio_Etop_codigo"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   2760
            TabIndex        =   131
            Top             =   600
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483637
            ListField       =   "Etop_codigo"
            BoundColumn     =   "Etop_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo8 
            Bindings        =   "frmfo_FA_formulario.frx":5C2D
            DataField       =   "Sitio_Etop_codigo"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   240
            TabIndex        =   132
            Top             =   600
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "Etop_descripcion"
            BoundColumn     =   "Etop_codigo"
            Text            =   ""
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Sancion:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   136
            Top             =   1920
            Width           =   630
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   $"frmfo_FA_formulario.frx":5C4A
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   135
            Top             =   2760
            Width           =   8250
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Motivo:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   240
            TabIndex        =   134
            Top             =   1080
            Width           =   525
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   $"frmfo_FA_formulario.frx":5CD7
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   133
            Top             =   360
            Width           =   7740
         End
      End
      Begin VB.CommandButton Command26 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Modif."
         Enabled         =   0   'False
         Height          =   600
         Left            =   8640
         Picture         =   "frmfo_FA_formulario.frx":5D69
         Style           =   1  'Graphical
         TabIndex        =   120
         ToolTipText     =   "Registro Ubicacion Fisica"
         Top             =   1380
         Width           =   645
      End
      Begin VB.CommandButton Command25 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Nuevo"
         Height          =   600
         Left            =   8640
         MousePointer    =   4  'Icon
         Picture         =   "frmfo_FA_formulario.frx":61AB
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "Nuevo Registro"
         Top             =   780
         Width           =   645
      End
      Begin VB.CommandButton Command24 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Grabar"
         Height          =   600
         Left            =   8640
         Picture         =   "frmfo_FA_formulario.frx":64B5
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   1980
         Width           =   645
      End
      Begin VB.CommandButton Command23 
         BackColor       =   &H80000013&
         Caption         =   "Modif."
         Enabled         =   0   'False
         Height          =   600
         Left            =   8640
         Picture         =   "frmfo_FA_formulario.frx":66BF
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Registro Ubicacion Fisica"
         Top             =   3300
         Width           =   645
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H80000013&
         Caption         =   "Nuevo"
         Height          =   600
         Left            =   8640
         MousePointer    =   4  'Icon
         Picture         =   "frmfo_FA_formulario.frx":6B01
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Nuevo Registro"
         Top             =   2700
         Width           =   645
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H80000013&
         Caption         =   "Grabar"
         Height          =   600
         Left            =   8640
         Picture         =   "frmfo_FA_formulario.frx":6E0B
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   3900
         Width           =   645
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Caption         =   "4. REGISTRO DE VACACIONES"
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
         ForeColor       =   &H00008000&
         Height          =   3585
         Left            =   120
         TabIndex        =   101
         Top             =   3480
         Width           =   9255
         Begin VB.TextBox Txt_InvBs 
            Alignment       =   2  'Center
            DataField       =   "Inversion_Total_Bs"
            DataSource      =   "ADO_M1"
            Height          =   285
            Left            =   3600
            ScrollBars      =   2  'Vertical
            TabIndex        =   123
            Text            =   "0"
            Top             =   3120
            Width           =   1935
         End
         Begin VB.TextBox Txt_ProyVUMes 
            Alignment       =   2  'Center
            DataField       =   "Vida_Util_Proy_Mes"
            DataSource      =   "ADO_M1"
            Height          =   285
            Left            =   6720
            ScrollBars      =   2  'Vertical
            TabIndex        =   122
            Text            =   "0"
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox Txt_ProyVUAnio 
            Alignment       =   2  'Center
            DataField       =   "Vida_Util_Proy_Anio"
            DataSource      =   "ADO_M1"
            Height          =   285
            Left            =   3600
            ScrollBars      =   2  'Vertical
            TabIndex        =   121
            Text            =   "0"
            Top             =   600
            Width           =   1455
         End
         Begin VB.ComboBox Combo12 
            DataField       =   "DDRR_Anio"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":7015
            Left            =   240
            List            =   "frmfo_FA_formulario.frx":703D
            TabIndex        =   103
            Text            =   "2010"
            Top             =   600
            Width           =   1575
         End
         Begin VB.ComboBox Combo10 
            DataField       =   "DDRR_Anio"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":7089
            Left            =   7200
            List            =   "frmfo_FA_formulario.frx":7093
            TabIndex        =   102
            Text            =   "NO"
            Top             =   2280
            Width           =   735
         End
         Begin MSComCtl2.DTPicker DTPicker12 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   240
            TabIndex        =   104
            Top             =   1440
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301697
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker13 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   3600
            TabIndex        =   105
            Top             =   1440
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301698
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker14 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   240
            TabIndex        =   106
            Top             =   2280
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301697
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker15 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   3600
            TabIndex        =   107
            Top             =   2280
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301698
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker16 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   240
            TabIndex        =   108
            Top             =   3120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301697
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo DataCombo5 
            Bindings        =   "frmfo_FA_formulario.frx":709F
            DataField       =   "Sitio_Etop_codigo"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   5520
            TabIndex        =   109
            Top             =   1440
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483637
            ListField       =   "Etop_codigo"
            BoundColumn     =   "Etop_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo6 
            Bindings        =   "frmfo_FA_formulario.frx":70BC
            DataField       =   "Sitio_Etop_codigo"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   6480
            TabIndex        =   110
            Top             =   1440
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "Etop_descripcion"
            BoundColumn     =   "Etop_codigo"
            Text            =   ""
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   $"frmfo_FA_formulario.frx":70D9
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   114
            Top             =   360
            Width           =   7890
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   $"frmfo_FA_formulario.frx":7176
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   113
            Top             =   1200
            Width           =   8370
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   $"frmfo_FA_formulario.frx":720F
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   112
            Top             =   2040
            Width           =   8220
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Fecha Reincorporacion:                                              Nro. Memo:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   111
            Top             =   2880
            Width           =   4605
         End
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H80000018&
         Caption         =   "Grabar"
         Height          =   720
         Left            =   -66360
         Picture         =   "frmfo_FA_formulario.frx":72AA
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   2700
         Width           =   645
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H80000018&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   -66360
         MousePointer    =   4  'Icon
         Picture         =   "frmfo_FA_formulario.frx":74B4
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Nuevo Registro"
         Top             =   1140
         Width           =   645
      End
      Begin VB.CommandButton cmdMod3 
         BackColor       =   &H80000018&
         Caption         =   "Modif."
         Enabled         =   0   'False
         Height          =   720
         Left            =   -66360
         Picture         =   "frmfo_FA_formulario.frx":77BE
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Registro Ubicacion Fisica"
         Top             =   1860
         Width           =   645
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   -66360
         MousePointer    =   4  'Icon
         Picture         =   "frmfo_FA_formulario.frx":7C00
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Nuevo Registro"
         Top             =   900
         Width           =   645
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Grabar"
         Height          =   720
         Left            =   -66360
         Picture         =   "frmfo_FA_formulario.frx":7F0A
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   2460
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Modif."
         Enabled         =   0   'False
         Height          =   720
         Left            =   -66360
         Picture         =   "frmfo_FA_formulario.frx":8114
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Registro Ubicacion Fisica"
         Top             =   1620
         Width           =   645
      End
      Begin VB.Frame FraEmpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Caption         =   "2. CONTROL DE ASISTENCIA"
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
         ForeColor       =   &H00008000&
         Height          =   3780
         Left            =   -74880
         TabIndex        =   42
         Top             =   4260
         Width           =   9360
         Begin VB.ComboBox Combo5 
            DataField       =   "DDRR_Anio"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":8556
            Left            =   8160
            List            =   "frmfo_FA_formulario.frx":8560
            TabIndex        =   72
            Text            =   "NO"
            Top             =   2640
            Width           =   735
         End
         Begin VB.ComboBox Combo4 
            DataField       =   "DDRR_Anio"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":856C
            Left            =   6840
            List            =   "frmfo_FA_formulario.frx":8576
            TabIndex        =   71
            Text            =   "NO"
            Top             =   2640
            Width           =   735
         End
         Begin VB.ComboBox Combo3 
            DataField       =   "DDRR_Anio"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":8582
            Left            =   8160
            List            =   "frmfo_FA_formulario.frx":858C
            TabIndex        =   69
            Text            =   "NO"
            Top             =   1680
            Width           =   735
         End
         Begin VB.ComboBox Combo2 
            DataField       =   "DDRR_Anio"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":8598
            Left            =   6840
            List            =   "frmfo_FA_formulario.frx":85A2
            TabIndex        =   68
            Text            =   "NO"
            Top             =   1680
            Width           =   735
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "DDRR_Anio"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":85AE
            Left            =   7080
            List            =   "frmfo_FA_formulario.frx":85C7
            TabIndex        =   67
            Text            =   "LUN"
            Top             =   720
            Width           =   1815
         End
         Begin VB.ComboBox TxtProyDDRRAnio 
            DataField       =   "DDRR_Anio"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":85EE
            Left            =   240
            List            =   "frmfo_FA_formulario.frx":8616
            TabIndex        =   66
            Text            =   "ENERO"
            Top             =   720
            Width           =   2535
         End
         Begin VB.CommandButton Cmd_comuNew 
            BackColor       =   &H008080FF&
            DragIcon        =   "frmfo_FA_formulario.frx":867F
            Height          =   675
            Left            =   7920
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmfo_FA_formulario.frx":8989
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Datos de la Empresa"
            Top             =   840
            Width           =   560
         End
         Begin MSComCtl2.DTPicker DtcProyUsoSFech 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   4080
            TabIndex        =   61
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301697
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   240
            TabIndex        =   62
            Top             =   1680
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301698
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   4080
            TabIndex        =   63
            Top             =   1680
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301698
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   240
            TabIndex        =   64
            Top             =   2640
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301698
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   4080
            TabIndex        =   65
            Top             =   2640
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301698
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   $"frmfo_FA_formulario.frx":91CB
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   70
            Top             =   2400
            Width           =   8475
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   $"frmfo_FA_formulario.frx":926F
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   45
            Left            =   240
            TabIndex        =   45
            Top             =   480
            Width           =   7950
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   $"frmfo_FA_formulario.frx":9306
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   40
            Left            =   240
            TabIndex        =   44
            Top             =   1440
            Width           =   8475
         End
      End
      Begin VB.Frame FraProyecto 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Caption         =   "3. CONTROL DE PERMISOS"
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
         ForeColor       =   &H00008000&
         Height          =   3705
         Left            =   -74880
         TabIndex        =   41
         Top             =   4380
         Width           =   9255
         Begin VB.ComboBox Combo9 
            DataField       =   "DDRR_Anio"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":93AA
            Left            =   7200
            List            =   "frmfo_FA_formulario.frx":93B4
            TabIndex        =   93
            Text            =   "NO"
            Top             =   2400
            Width           =   735
         End
         Begin VB.ComboBox Combo8 
            DataField       =   "DDRR_Anio"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":93C0
            Left            =   3840
            List            =   "frmfo_FA_formulario.frx":93D9
            TabIndex        =   89
            Text            =   "LUN"
            Top             =   3240
            Width           =   1815
         End
         Begin VB.ComboBox Combo7 
            DataField       =   "DDRR_Anio"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":9400
            Left            =   240
            List            =   "frmfo_FA_formulario.frx":9428
            TabIndex        =   81
            Text            =   "ENERO"
            Top             =   720
            Width           =   2535
         End
         Begin VB.ComboBox Combo6 
            DataField       =   "DDRR_Anio"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":9491
            Left            =   6840
            List            =   "frmfo_FA_formulario.frx":94AA
            TabIndex        =   80
            Text            =   "LUN"
            Top             =   720
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DTPicker5 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   3840
            TabIndex        =   79
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301697
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker6 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   240
            TabIndex        =   83
            Top             =   1560
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301697
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker7 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   3840
            TabIndex        =   84
            Top             =   1560
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301698
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker8 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   240
            TabIndex        =   85
            Top             =   2400
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301697
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker9 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   3840
            TabIndex        =   86
            Top             =   2400
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301698
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker DTPicker10 
            DataField       =   "uso_Suelo_Certif_Fecha"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   240
            TabIndex        =   88
            Top             =   3240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102301697
            CurrentDate     =   40179
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo Dtc_SitioTopo 
            Bindings        =   "frmfo_FA_formulario.frx":94D1
            DataField       =   "Sitio_Etop_codigo"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   5640
            TabIndex        =   91
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483637
            ListField       =   "Etop_codigo"
            BoundColumn     =   "Etop_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_SitioTopoDes 
            Bindings        =   "frmfo_FA_formulario.frx":94EE
            DataField       =   "Sitio_Etop_codigo"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   6480
            TabIndex        =   92
            Top             =   1560
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "Etop_descripcion"
            BoundColumn     =   "Etop_codigo"
            Text            =   ""
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Fecha Reincorporacion:                                              Dias de Permiso:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   90
            Top             =   3000
            Width           =   4965
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   $"frmfo_FA_formulario.frx":950B
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   87
            Top             =   2160
            Width           =   8220
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   $"frmfo_FA_formulario.frx":95A6
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   82
            Top             =   1320
            Width           =   7815
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   $"frmfo_FA_formulario.frx":963A
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   78
            Top             =   480
            Width           =   7815
         End
      End
      Begin VB.Frame FraCabecera 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Caption         =   "1. INFORMACION GENERAL"
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
         ForeColor       =   &H00008000&
         Height          =   7260
         Left            =   -74880
         TabIndex        =   25
         Top             =   900
         Width           =   9135
         Begin VB.TextBox txtCorrel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            DataField       =   "Item"
            DataSource      =   "ADO_M1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   420
            Index           =   2
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   945
            Width           =   1335
         End
         Begin VB.TextBox TxtObs 
            BackColor       =   &H00FFC0C0&
            DataField       =   "Posibles_Contingencias"
            DataSource      =   "ADO_M1"
            Height          =   525
            Left            =   165
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   48
            Text            =   "frmfo_FA_formulario.frx":96CE
            Top             =   6600
            Width           =   8655
         End
         Begin VB.ComboBox txtParam 
            DataField       =   "Gestion"
            DataSource      =   "ADO_M1"
            Height          =   315
            ItemData        =   "frmfo_FA_formulario.frx":96D1
            Left            =   990
            List            =   "frmfo_FA_formulario.frx":9726
            TabIndex        =   1
            Text            =   "2010"
            Top             =   945
            Width           =   1095
         End
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Caption         =   "LUGAR DEL TRABAJO"
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
            Height          =   1275
            Left            =   165
            TabIndex        =   32
            Top             =   3240
            Width           =   8775
            Begin MSDataListLib.DataCombo DataCombo2 
               Bindings        =   "frmfo_FA_formulario.frx":97CC
               DataField       =   "munic_codigo"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   5280
               TabIndex        =   54
               Top             =   840
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483628
               ListField       =   "munic_descripcion"
               BoundColumn     =   "munic_codigo"
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
            Begin MSDataListLib.DataCombo Dtc_munic 
               Bindings        =   "frmfo_FA_formulario.frx":97E3
               DataField       =   "munic_codigo"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   1200
               TabIndex        =   5
               Top             =   840
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483628
               ListField       =   "munic_descripcion"
               BoundColumn     =   "munic_codigo"
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
            Begin MSDataListLib.DataCombo Dtc_prov 
               Bindings        =   "frmfo_FA_formulario.frx":97FA
               DataField       =   "prov_codigo"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   5280
               TabIndex        =   4
               Top             =   360
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483628
               ListField       =   "prov_descripcion"
               BoundColumn     =   "prov_codigo"
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
            Begin MSDataListLib.DataCombo Dtc_depto 
               Bindings        =   "frmfo_FA_formulario.frx":9811
               DataField       =   "depto_codigo"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   1200
               TabIndex        =   3
               Top             =   345
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ListField       =   "depto_descripcion"
               BoundColumn     =   "depto_codigo"
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
            Begin MSDataListLib.DataCombo Dtc_prov_cod 
               Bindings        =   "frmfo_FA_formulario.frx":9829
               DataField       =   "prov_codigo"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   5280
               TabIndex        =   10
               Top             =   240
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483637
               ListField       =   "prov_codigo"
               BoundColumn     =   "prov_codigo"
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
            Begin MSDataListLib.DataCombo Dtc_munic_cod 
               Bindings        =   "frmfo_FA_formulario.frx":9840
               DataField       =   "munic_codigo"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   1440
               TabIndex        =   11
               Top             =   600
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483637
               ListField       =   "munic_codigo"
               BoundColumn     =   "munic_codigo"
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
            Begin MSDataListLib.DataCombo Dtc_depto_cod 
               Bindings        =   "frmfo_FA_formulario.frx":9857
               DataField       =   "depto_codigo"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   1440
               TabIndex        =   9
               Top             =   240
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483637
               ListField       =   "depto_codigo"
               BoundColumn     =   "depto_codigo"
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
            Begin MSDataListLib.DataCombo DataCombo1 
               Bindings        =   "frmfo_FA_formulario.frx":986F
               DataField       =   "munic_codigo"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   5280
               TabIndex        =   53
               Top             =   720
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483637
               ListField       =   "munic_codigo"
               BoundColumn     =   "munic_codigo"
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
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Municipio                                                                                   Localidad"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   44
               Left            =   120
               TabIndex        =   52
               Top             =   840
               Width           =   5100
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Departamento                                                                            Provincia"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   33
               Top             =   360
               Width           =   5085
            End
         End
         Begin VB.Frame FraResp 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Caption         =   "NOMBRE DEL JEFE INMEDIATO SUPERIOR"
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
            Height          =   1500
            Left            =   165
            TabIndex        =   30
            Top             =   4740
            Width           =   8775
            Begin VB.CommandButton Cmd_PersNuevo 
               BackColor       =   &H00FFFFC0&
               Height          =   675
               Left            =   8040
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmfo_FA_formulario.frx":9886
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Datos del Consultor"
               Top             =   180
               Width           =   570
            End
            Begin MSDataListLib.DataCombo DtcRespNom 
               Bindings        =   "frmfo_FA_formulario.frx":9B90
               DataField       =   "codigo_responsable"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   1800
               TabIndex        =   6
               Top             =   300
               Width           =   6135
               _ExtentX        =   10821
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483628
               ListField       =   "denominacion_beneficiario"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DtcRespId 
               Bindings        =   "frmfo_FA_formulario.frx":9BAC
               DataField       =   "codigo_responsable"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   165
               TabIndex        =   12
               Top             =   300
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483637
               ListField       =   "codigo_beneficiario"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DtcRespRenca 
               Bindings        =   "frmfo_FA_formulario.frx":9BC8
               DataField       =   "codigo_responsable"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   4440
               TabIndex        =   13
               Top             =   975
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483637
               ListField       =   "Reg_Profesional"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DtcRespCargo 
               Bindings        =   "frmfo_FA_formulario.frx":9BE4
               DataField       =   "codigo_responsable"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   165
               TabIndex        =   35
               Top             =   975
               Width           =   4575
               _ExtentX        =   8070
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483637
               ListField       =   "cargo"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Puesto del Jefe:                                                                        Unidad/Oficina:"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   22
               Left            =   165
               TabIndex        =   31
               Top             =   765
               Width           =   5505
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Caption         =   "NOMBRE Y PUESTO DE LA PERSONA"
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
            Height          =   1485
            Left            =   165
            TabIndex        =   36
            Top             =   1560
            Width           =   8775
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               DataField       =   "cargo"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   240
               TabIndex        =   96
               Text            =   "NS"
               Top             =   1080
               Width           =   4215
            End
            Begin VB.TextBox TxtSupOcu 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               DataField       =   "denominacion_beneficiario"
               DataSource      =   "ADO_M1"
               Height          =   285
               Left            =   1920
               TabIndex        =   95
               Text            =   "NS"
               Top             =   360
               Width           =   6015
            End
            Begin VB.TextBox TxtCI 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               DataField       =   "codigo_beneficiario"
               DataSource      =   "ADO_M1"
               Height          =   285
               Left            =   165
               ScrollBars      =   2  'Vertical
               TabIndex        =   94
               Text            =   "NS"
               Top             =   360
               Width           =   1695
            End
            Begin VB.CommandButton Cmd_PersNuevo2 
               BackColor       =   &H00C0FFC0&
               Height          =   675
               Left            =   8040
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmfo_FA_formulario.frx":9C00
               Style           =   1  'Graphical
               TabIndex        =   40
               ToolTipText     =   "Datos del Consultor"
               Top             =   360
               Width           =   570
            End
            Begin MSDataListLib.DataCombo DataCombo3 
               Bindings        =   "frmfo_FA_formulario.frx":9F0A
               DataField       =   "depto_codigo"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   4440
               TabIndex        =   97
               Top             =   1080
               Visible         =   0   'False
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483637
               ListField       =   "depto_codigo"
               BoundColumn     =   "depto_codigo"
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
            Begin MSDataListLib.DataCombo DataCombo4 
               Bindings        =   "frmfo_FA_formulario.frx":9F22
               DataField       =   "depto_codigo"
               DataSource      =   "ADO_M1"
               Height          =   315
               Left            =   6000
               TabIndex        =   98
               Top             =   720
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483637
               ListField       =   "depto_codigo"
               BoundColumn     =   "depto_codigo"
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
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Puesto de la Persona:                                                              Unidad/Oficina:"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   6
               Left            =   240
               TabIndex        =   55
               Top             =   840
               Width           =   5475
            End
         End
         Begin MSComCtl2.DTPicker DTP_FechaLectura 
            DataField       =   "Fecha_Ingreso"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   7320
            TabIndex        =   2
            Top             =   945
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   -2147483624
            CheckBox        =   -1  'True
            Format          =   102301697
            CurrentDate     =   39097
            MinDate         =   36526
         End
         Begin MSDataListLib.DataCombo DtCForm 
            Bindings        =   "frmfo_FA_formulario.frx":9F3A
            DataField       =   "Tipo_Formulario"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   990
            TabIndex        =   0
            Top             =   405
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "Tipo_Formulario"
            BoundColumn     =   "Tipo_Formulario"
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
         Begin MSDataListLib.DataCombo DtCFormDes 
            Bindings        =   "frmfo_FA_formulario.frx":9F51
            DataField       =   "Tipo_Formulario"
            DataSource      =   "ADO_M1"
            Height          =   315
            Left            =   2160
            TabIndex        =   38
            Top             =   405
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            MatchEntry      =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483637
            ListField       =   "Denominacion_Tipo"
            BoundColumn     =   "Tipo_Formulario"
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
         Begin VB.Label LblObs 
            BackColor       =   &H8000000B&
            Caption         =   "OBSERVACIONES:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   6360
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   $"frmfo_FA_formulario.frx":9F68
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   37
            Top             =   975
            Width           =   7110
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Tipo Form.:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   27
            Top             =   435
            Width           =   795
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmfo_FA_formulario.frx":9FF2
         Height          =   1215
         Left            =   -74880
         TabIndex        =   73
         Top             =   720
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
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
         Caption         =   "RESUMEN EVALUACION DE DESEMPEÑO"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Tipo"
            Caption         =   "Tipo Evaluacion"
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
            DataField       =   "Ocupacion"
            Caption         =   "Objetivo Evaluacion"
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
            DataField       =   "nro_personas"
            Caption         =   "% Calificacion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cargo_funcion"
            Caption         =   "Recomendaciones"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2594.835
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3614.74
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmfo_FA_formulario.frx":A009
         Height          =   1215
         Left            =   -74880
         TabIndex        =   75
         Top             =   2010
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
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
         Caption         =   "RESUMEN DE PERMISOS"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Tipo"
            Caption         =   "Tipo Permiso"
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
            DataField       =   "Ocupacion"
            Caption         =   "Justificacion / Motivo"
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
            DataField       =   "nro_personas"
            Caption         =   "Fecha Permiso"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cargo_funcion"
            Caption         =   "Permiso con Cargo a:"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   ""
            Caption         =   "Tiempo Min."
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
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3300.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1560.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   945.071
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmfo_FA_formulario.frx":A020
         Height          =   1215
         Left            =   -74880
         TabIndex        =   99
         Top             =   3285
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
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
         Caption         =   "RESUMEN DE VACACIONES"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Tipo"
            Caption         =   "Programado"
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
            DataField       =   "Ocupacion"
            Caption         =   "Dias Progs."
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
            DataField       =   ""
            Caption         =   "Horas Progs."
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
            DataField       =   ""
            Caption         =   "Dias Utilizados"
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
            DataField       =   ""
            Caption         =   "Horas Utilizadas"
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
            DataField       =   "nro_personas"
            Caption         =   "Fecha Prog.Inicio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "cargo_funcion"
            Caption         =   "Fecha Prog.Fin"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1214.929
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "frmfo_FA_formulario.frx":A037
         Height          =   1215
         Left            =   -74880
         TabIndex        =   100
         Top             =   4560
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
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
         Caption         =   "RESUMEN DE MEMORANDUMS"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Tipo"
            Caption         =   "Tipo Memo"
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
            DataField       =   "Ocupacion"
            Caption         =   "Motivo"
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
            DataField       =   "nro_personas"
            Caption         =   "Fecha Inicio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cargo_funcion"
            Caption         =   "Emitido por:"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3525.166
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2369.764
            EndProperty
         EndProperty
      End
      Begin VB.Image ImgEvaluacion 
         Height          =   540
         Left            =   -66480
         Picture         =   "frmfo_FA_formulario.frx":A04E
         Top             =   1005
         Width           =   555
      End
      Begin VB.Image Image2 
         Height          =   540
         Left            =   -66480
         Picture         =   "frmfo_FA_formulario.frx":A3D6
         Top             =   2325
         Width           =   555
      End
      Begin VB.Image ImgVacacion 
         Height          =   540
         Left            =   -66480
         Picture         =   "frmfo_FA_formulario.frx":A75E
         Top             =   3525
         Width           =   555
      End
      Begin VB.Image ImgMemo 
         Height          =   540
         Left            =   -66480
         Picture         =   "frmfo_FA_formulario.frx":AAE6
         Top             =   4845
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Dirección Oficina:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   124
         Top             =   7545
         Width           =   1260
      End
   End
   Begin MSAdodcLib.Adodc ADO_M1 
      Height          =   330
      Left            =   60
      Top             =   6000
      Width           =   6000
      _ExtentX        =   10583
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
   Begin MSAdodcLib.Adodc Ado_Form 
      Height          =   330
      Left            =   0
      Top             =   9240
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
      Caption         =   "Ado_Form"
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
   Begin MSAdodcLib.Adodc Ado_Depto 
      Height          =   330
      Left            =   6120
      Top             =   9240
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
      Caption         =   "Ado_Depto"
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
   Begin MSAdodcLib.Adodc Ado_consultor 
      Height          =   330
      Left            =   10200
      Top             =   9600
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
      Caption         =   "Ado_consultor"
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
   Begin MSAdodcLib.Adodc Ado_prov 
      Height          =   330
      Left            =   6120
      Top             =   9600
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
      Caption         =   "Ado_Prov"
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
   Begin MSAdodcLib.Adodc Ado_Muni 
      Height          =   330
      Left            =   6120
      Top             =   9960
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
      Caption         =   "Ado_Muni"
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
   Begin MSAdodcLib.Adodc Ado_Comunid 
      Height          =   330
      Left            =   8160
      Top             =   9960
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
      Caption         =   "Ado_Comunid"
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
   Begin MSAdodcLib.Adodc Ado_persona 
      Height          =   330
      Left            =   10200
      Top             =   9240
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
      Caption         =   "Ado_persona"
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
   Begin MSAdodcLib.Adodc Ado_Empresa 
      Height          =   330
      Left            =   8160
      Top             =   9240
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
      Caption         =   "Ado_Empresa"
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
   Begin MSAdodcLib.Adodc Ado_Establ 
      Height          =   330
      Left            =   8160
      Top             =   9600
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
      Caption         =   "Ado_Establ"
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
   Begin MSAdodcLib.Adodc Ado_Ocupac 
      Height          =   330
      Left            =   10200
      Top             =   9960
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
      Caption         =   "Ado_Ocupacion"
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
   Begin MSAdodcLib.Adodc Ado_ProyUbic 
      Height          =   330
      Left            =   0
      Top             =   9600
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
      Caption         =   "Ado_ProyUbic"
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
   Begin MSAdodcLib.Adodc Ado_tipo_prueba 
      Height          =   330
      Left            =   4080
      Top             =   9240
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
      Caption         =   "Ado_Tipo_prueba"
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
   Begin MSAdodcLib.Adodc Ado_estado_caso 
      Height          =   330
      Left            =   4080
      Top             =   9600
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
      Caption         =   "Ado_Estado_caso"
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
   Begin MSAdodcLib.Adodc Ado_tipo_caso 
      Height          =   330
      Left            =   4080
      Top             =   9960
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
      Caption         =   "Ado_Tipo_caso"
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
   Begin MSAdodcLib.Adodc Ado_Paracito 
      Height          =   330
      Left            =   2040
      Top             =   9240
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
      Caption         =   "Ado_Paracito"
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
   Begin MSAdodcLib.Adodc Ado_Tipo_Infec 
      Height          =   330
      Left            =   2040
      Top             =   9600
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
      Caption         =   "Ado_Tipo_Infec"
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
   Begin MSAdodcLib.Adodc Ado_sem_epidem 
      Height          =   330
      Left            =   2040
      Top             =   9960
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
      Caption         =   "Ado_sem_epidem"
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
   Begin MSAdodcLib.Adodc Ado_Actividad 
      Height          =   330
      Left            =   0
      Top             =   9960
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
      Caption         =   "Ado_Actividad"
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
   Begin MSAdodcLib.Adodc Ado_Depto2 
      Height          =   330
      Left            =   0
      Top             =   8880
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
      Caption         =   "Ado_Depto"
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
   Begin MSAdodcLib.Adodc Ado_prov2 
      Height          =   330
      Left            =   2040
      Top             =   8880
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
      Caption         =   "Ado_Prov"
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
   Begin MSAdodcLib.Adodc Ado_Muni2 
      Height          =   330
      Left            =   4080
      Top             =   8880
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
      Caption         =   "Ado_Muni"
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
   Begin MSAdodcLib.Adodc Ado_Depto3 
      Height          =   330
      Left            =   6120
      Top             =   8880
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
      Caption         =   "Ado_Depto3"
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
   Begin MSAdodcLib.Adodc Ado_CtrlAsistencia 
      Height          =   330
      Left            =   8160
      Top             =   8880
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
      Caption         =   "Ado_CtrlAsistencia"
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
   Begin MSAdodcLib.Adodc Ado_Topografia 
      Height          =   330
      Left            =   0
      Top             =   10320
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
      Caption         =   "Ado_Topografia"
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
   Begin MSAdodcLib.Adodc ado_napa_freatica 
      Height          =   330
      Left            =   2040
      Top             =   10320
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
      Caption         =   "ado_napa_freatica"
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
   Begin MSAdodcLib.Adodc Ado_Agua_Calidad 
      Height          =   330
      Left            =   4080
      Top             =   10320
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
      Caption         =   "Ado_Agua_Calidad"
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
   Begin MSAdodcLib.Adodc ado_vegetacion 
      Height          =   330
      Left            =   6120
      Top             =   10320
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
      Caption         =   "ado_vegetacion"
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
   Begin MSAdodcLib.Adodc Ado_Red_Drenaje 
      Height          =   330
      Left            =   8160
      Top             =   10320
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
      Caption         =   "Ado_Red_Drenaje"
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
   Begin MSAdodcLib.Adodc Ado_SectorAct 
      Height          =   330
      Left            =   12240
      Top             =   9600
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
      Caption         =   "Ado_SectorAct"
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
   Begin MSAdodcLib.Adodc Ado_SubSector 
      Height          =   330
      Left            =   12240
      Top             =   9240
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
      Caption         =   "Ado_SubSector"
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
   Begin MSAdodcLib.Adodc Ado_ProyNat 
      Height          =   330
      Left            =   12240
      Top             =   9960
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
      Caption         =   "Ado_ProyNat"
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
   Begin MSAdodcLib.Adodc Ado_CtrlPermiso 
      Height          =   330
      Left            =   10200
      Top             =   8880
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
      Caption         =   "Ado_CtrlPermiso"
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
   Begin MSAdodcLib.Adodc Ado_ProyAmbito 
      Height          =   330
      Left            =   10200
      Top             =   10320
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
      Caption         =   "Ado_ProyAmbito"
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
   Begin MSAdodcLib.Adodc Ado_Proy_tipo 
      Height          =   330
      Left            =   12240
      Top             =   10320
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
      Caption         =   "Ado_Proy_tipo"
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
   Begin MSAdodcLib.Adodc Ado_Proyecto 
      Height          =   330
      Left            =   0
      Top             =   10680
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
      Caption         =   "Ado_Proyecto"
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
   Begin MSAdodcLib.Adodc Ado_ProyFase 
      Height          =   330
      Left            =   2040
      Top             =   10680
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
      Caption         =   "Ado_ProyFase"
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
   Begin MSAdodcLib.Adodc Ado_ProyEtapa 
      Height          =   330
      Left            =   4080
      Top             =   10680
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
      Caption         =   "Ado_ProyEtapa"
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
   Begin MSAdodcLib.Adodc AdoUniMed 
      Height          =   330
      Left            =   6120
      Top             =   10680
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
      Caption         =   "AdoUniMed"
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
   Begin MSAdodcLib.Adodc Ado_MaqEq 
      Height          =   330
      Left            =   8160
      Top             =   10680
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
      Caption         =   "Ado_MaqEq"
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
   Begin MSAdodcLib.Adodc Ado_RRHH 
      Height          =   330
      Left            =   10200
      Top             =   10680
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
      Caption         =   "Ado_RRH"
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
   Begin MSAdodcLib.Adodc Ado_ServMant 
      Height          =   330
      Left            =   12240
      Top             =   10680
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
      Caption         =   "Ado_ServMant"
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
   Begin MSAdodcLib.Adodc Ado_FteFin 
      Height          =   330
      Left            =   0
      Top             =   11040
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
      Caption         =   "Ado_FteFin"
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
   Begin MSAdodcLib.Adodc Ado_RRNN 
      Height          =   330
      Left            =   2040
      Top             =   11040
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
      Caption         =   "Ado_RRNN"
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
   Begin MSAdodcLib.Adodc Ado_MatPrima 
      Height          =   330
      Left            =   4080
      Top             =   11040
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
      Caption         =   "AdoMatPrima"
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
   Begin MSAdodcLib.Adodc Ado_Energia 
      Height          =   330
      Left            =   6120
      Top             =   11040
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
      Caption         =   "Ado_Energia"
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
   Begin MSAdodcLib.Adodc Ado_Desecho 
      Height          =   330
      Left            =   8160
      Top             =   11040
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
      Caption         =   "Ado_Desecho"
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
   Begin MSAdodcLib.Adodc Ado_Ruido 
      Height          =   330
      Left            =   14280
      Top             =   8880
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
      Caption         =   "Ado_Ruido"
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
   Begin MSAdodcLib.Adodc Ado_IA_MM 
      Height          =   330
      Left            =   14280
      Top             =   9240
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
      Caption         =   "Ado_IA_MM"
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
   Begin MSAdodcLib.Adodc Ado_ComunidV 
      Height          =   330
      Left            =   14280
      Top             =   9600
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
      Caption         =   "Ado_ComunidV"
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
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRO DE CONTROL DE PERSONAL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   405
      Left            =   8370
      TabIndex        =   46
      Top             =   120
      Width           =   6690
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   12720
      TabIndex        =   34
      Top             =   210
      Width           =   75
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   0
      Picture         =   "frmfo_FA_formulario.frx":AE6E
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "frmfo_FA_formulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declaracion de variables
Dim rs_CtrlAsistencia As New ADODB.Recordset

Dim rs_M1, RSNADA As New ADODB.Recordset
Dim rs_Form As New ADODB.Recordset
Dim rs_Depto, rs_Depto2, rs_Depto3 As New ADODB.Recordset
Dim rs_Region As New ADODB.Recordset
Dim rs_Prov, rs_Prov2 As New ADODB.Recordset
Dim rs_Muni, rs_Muni2 As New ADODB.Recordset
Dim rs_comunid, rs_comunidV As New ADODB.Recordset
Dim rs_persona As New ADODB.Recordset
Dim rs_consultor As New ADODB.Recordset
Dim rs_Empresa As New ADODB.Recordset

Dim rs_correlativo As New ADODB.Recordset
Dim rs_ocupac As New ADODB.Recordset
Dim rs_ProyUbic, rs_topografia As New ADODB.Recordset
Dim rs_Agua_Calidad, rc_napa_freatica As New ADODB.Recordset
Dim rs_Red_Drenaje, rs_vegetacion As New ADODB.Recordset
Dim rs_CtrlPermiso, rs_subsector, rs_Sector_Act As New ADODB.Recordset
Dim rs_ProyNat, rs_Proy_tipo, rs_Proyecto As New ADODB.Recordset
Dim rs_ProyAmbito, rs_ProyFase, RS_ProyEtapa As New ADODB.Recordset
Dim rs_MaqEq, rs_RRHH, rs_ServMant As New ADODB.Recordset
Dim rs_UniMed, rs_sem_aux As New ADODB.Recordset
Dim rs_Actividad, rs_FteFin As New ADODB.Recordset
Dim rs_RRNN, rs_Ruido As New ADODB.Recordset
Dim RS_MatPrima, rs_Energia, rs_Desecho As New ADODB.Recordset
Dim rs_TramiteC, rs_U_Fisisca, rs_U_FIS2 As New ADODB.Recordset
'Buscador
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim queryinicial As String
'Carpeta
Dim e As String
Dim NombreCarpeta As String
Dim Mensaje As String

Dim mvBookMark As Variant
Dim marca1 As BookmarkEnum

Dim var_form As String
Dim var_correl As String
Dim var_param, var_gest As String
Dim var_DocId As String
Dim var_FA, swgraba3 As Integer
'Public var_nom As String
Dim sino As String
Dim VAR_VAL As String

' Variables para grabar
Dim VARB, VARBD, VARG, VARS, VARU, VARP, varCat, VAR10, VAR11, VAR12, VAR13, VAR14, VAR15 As String
Dim VARPU, VARCAN, VARPT As Double

Private Sub ADO_M1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

  'If (Not ADO_M1.Recordset.BOF) And (Not ADO_M1.Recordset.EOF) Then
  If (Not ADO_M1.Recordset.EOF) Then
  
       Set rs_CtrlAsistencia = New ADODB.Recordset
       If GlSW <> "ADD" Then
            'rs_CtrlAsistencia.Open "select * from rc_ControlAsistencia where codigo_beneficiario = '" & ADO_M1.Recordset!codigo_beneficiario & "' and Item = " & ADO_M1.Recordset!Item & " ", DB, adOpenKeyset, adLockOptimistic
            rs_CtrlAsistencia.Open "select * from rc_ControlAsistencia ", DB, adOpenKeyset, adLockOptimistic
            Set Ado_CtrlAsistencia.Recordset = rs_CtrlAsistencia
       Else
            'Set Ado_ProyUbic.Recordset = RSNADA
            Set Dtg_CtrlAsistencia.DataSource = RSNADA
'            rs_ProyUbic.Open "select * from mo_proy_Id_Ubicacion  ", db, adOpenKeyset, adLockOptimistic
       End If
'        Set TDBGProyUbic.DataSource = Ado_ProyUbic.Recordset

       Set rs_CtrlPermiso = New ADODB.Recordset
       If GlSW <> "ADD" Then
            'rs_CtrlPermiso.Open "select * from rc_Permisos where codigo_beneficiario = '" & ADO_M1.Recordset!codigo_beneficiario & "' and Item = " & ADO_M1.Recordset!Item & "  ", DB, adOpenKeyset, adLockOptimistic
            rs_CtrlPermiso.Open "select * from rc_Permisos ", DB, adOpenKeyset, adLockOptimistic
            Set Ado_CtrlPermiso.Recordset = rs_CtrlPermiso
       Else
            Set Dtg_CtrlPermiso.DataSource = RSNADA
'            RS_ProyEtapa.Open "select * from mo_Proy_etapas   ", db, adOpenKeyset, adLockOptimistic
       End If

'       Set rs_MaqEq = New ADODB.Recordset
'       If GlSW <> "ADD" Then
'        rs_MaqEq.Open "select * from mo_tecnologia_maq where Gestion = '" & ADO_M1.Recordset!gestion & "' and Numero_FA = " & ADO_M1.Recordset!Numero_FA & "  AND CodGrupo='01' ", db, adOpenKeyset, adLockOptimistic
'       Set Ado_MaqEq.Recordset = rs_MaqEq
'       Set DgtMaqEquipo.DataSource = Ado_MaqEq.Recordset
''       Else
''        rs_MaqEq.Open "select * from mo_tecnologia_maq  ", db, adOpenKeyset, adLockOptimistic
'       End If
'
'       Set rs_RRHH = New ADODB.Recordset
'       If GlSW <> "ADD" Then
'        rs_RRHH.Open "select * from mo_Proy_RRHH where Gestion = '" & ADO_M1.Recordset!gestion & "' and Numero_FA = " & ADO_M1.Recordset!Numero_FA & "  ", db, adOpenKeyset, adLockOptimistic
'       Set Ado_RRHH.Recordset = rs_RRHH
''       Else
''        rs_RRHH.Open "select * from mo_Proy_RRHH  ", db, adOpenKeyset, adLockOptimistic
'       End If
 
'       Set rs_ServMant = New ADODB.Recordset
'       If GlSW <> "ADD" Then
'        rs_ServMant.Open "select * from mo_tecnologia_maq where Gestion = '" & ADO_M1.Recordset!gestion & "' and Numero_FA = " & ADO_M1.Recordset!Numero_FA & "  AND Cod_montador = '102' ", db, adOpenKeyset, adLockOptimistic
'        Set Ado_ServMant.Recordset = rs_ServMant
'       Set DgtServMant.DataSource = Ado_ServMant.Recordset
''       Else
''        rs_ServMant.Open "select * from mo_tecnologia_maq  ", db, adOpenKeyset, adLockOptimistic
'       End If
'
'       Set rs_Actividad = New ADODB.Recordset
'       If GlSW <> "ADD" Then
'        rs_Actividad.Open "select * from mo_Proy_Etapa_Act where Gestion = '" & ADO_M1.Recordset!gestion & "' and Numero_FA = " & ADO_M1.Recordset!Numero_FA & "  ", db, adOpenKeyset, adLockOptimistic
'        Set Ado_Actividad.Recordset = rs_Actividad
'        Set Ado_IA_MM.Recordset = rs_Actividad       ' Impactio Ambiental
''       Else
''        rs_Actividad.Open "select * from mo_Proy_Etapa_Act  ", db, adOpenKeyset, adLockOptimistic
'       End If
'       'Dtc_Act_Des.BoundText = Dtc_Act.BoundText
'
'       Set rs_FteFin = New ADODB.Recordset
'       If GlSW <> "ADD" Then
'        rs_FteFin.Open "select * from mo_Proy_FteFin where Gestion = '" & ADO_M1.Recordset!gestion & "' and Numero_FA = " & ADO_M1.Recordset!Numero_FA & "  ", db, adOpenKeyset, adLockOptimistic
'       Set Ado_FteFin.Recordset = rs_FteFin
''       Else
''        rs_FteFin.Open "select * from mo_Proy_FteFin  ", db, adOpenKeyset, adLockOptimistic
'       End If
'
'       Set rs_RRNN = New ADODB.Recordset
'       If rs_RRNN.State = 1 Then rs_RRNN.Close
'       If GlSW <> "ADD" Then
'        rs_RRNN.Open "select * from mo_Proy_RRNN WHERE Gestion = '" & ADO_M1.Recordset!gestion & "' and Numero_FA = " & ADO_M1.Recordset!Numero_FA & "  ", db, adOpenKeyset, adLockReadOnly
'       Set Ado_RRNN.Recordset = rs_RRNN
''       Else
''        rs_RRNN.Open "select * from mo_Proy_RRNN  ", db, adOpenKeyset, adLockReadOnly
'       End If
'
'       Set RS_MatPrima = New ADODB.Recordset
'       If GlSW <> "ADD" Then
'        RS_MatPrima.Open "select * from mo_tecnologia_maq where Gestion = '" & ADO_M1.Recordset!gestion & "' and Numero_FA = " & ADO_M1.Recordset!Numero_FA & "  AND Cod_montador = '21' ", db, adOpenKeyset, adLockOptimistic
'       Set Ado_MatPrima.Recordset = RS_MatPrima
'       Set DtgMatPrima.DataSource = Ado_MatPrima.Recordset
''       Else
''        RS_MatPrima.Open "select * from mo_tecnologia_maq  ", db, adOpenKeyset, adLockOptimistic
'       End If
'
'       Set rs_Energia = New ADODB.Recordset
'       If GlSW <> "ADD" Then
'        rs_Energia.Open "select * from mo_tecnologia_maq where Gestion = '" & ADO_M1.Recordset!gestion & "' and Numero_FA = " & ADO_M1.Recordset!Numero_FA & "  AND Cod_montador = '26' ", db, adOpenKeyset, adLockOptimistic
'       Set Ado_Energia.Recordset = rs_Energia
'       Set DtgEnergia.DataSource = Ado_Energia.Recordset
''      Else
''        rs_Energia.Open "select * from mo_tecnologia_maq  ", db, adOpenKeyset, adLockOptimistic
'       End If
'
'       Set rs_Desecho = New ADODB.Recordset
'       If GlSW <> "ADD" Then
'        rs_Desecho.Open "select * from mo_desecho_produccion where Gestion = '" & ADO_M1.Recordset!gestion & "' and Numero_FA = " & ADO_M1.Recordset!Numero_FA & "  ", db, adOpenKeyset, adLockOptimistic
'       Set Ado_Desecho.Recordset = rs_Desecho
'       Set DtgDesechos.DataSource = Ado_Desecho.Recordset
''      Else
''        rs_Desecho.Open "select * from mo_desecho_produccion  ", db, adOpenKeyset, adLockOptimistic
'       End If
'
'       Set rs_Ruido = New ADODB.Recordset
'       If GlSW <> "ADD" Then
'        rs_Ruido.Open "select * from mo_ruido_produccion where Gestion = '" & ADO_M1.Recordset!gestion & "' and Numero_FA = " & ADO_M1.Recordset!Numero_FA & "  ", db, adOpenKeyset, adLockOptimistic
'       Set Ado_Ruido.Recordset = rs_Ruido
'       Set DtgRuido.DataSource = Ado_Ruido.Recordset
''      Else
''        rs_Ruido.Open "select * from mo_ruido_produccion  ", db, adOpenKeyset, adLockOptimistic
'       End If
       '
       If ADO_M1.Recordset!estado_registro = "L" Then
            cmdAprueba.Visible = False
            CmdCopiar.Visible = False
            CmdDel.Visible = False
            CmdMod.Visible = False
            LblObs.Visible = True
            TxtObs.Visible = True
            cmdDesaprueba.Visible = False
            CmdObs.Visible = False
            Mensaje = "anuLado"
       End If
       If ADO_M1.Recordset!estado_registro = "R" Or ADO_M1.Recordset!estado_registro = "O" Then
            cmdAprueba.Visible = False
            CmdCopiar.Visible = True
            CmdDel.Visible = False
            CmdMod.Visible = True
            LblObs.Visible = True
            TxtObs.Visible = True
            cmdDesaprueba.Visible = False
            CmdObs.Visible = False
            Mensaje = "Rechazado/Observado"
       End If
       If ADO_M1.Recordset!estado_registro = "A" Or ADO_M1.Recordset!estado_registro = "S" Then
            cmdAprueba.Visible = False
            CmdCopiar.Visible = False
            CmdDel.Visible = False
            CmdMod.Visible = False
            LblObs.Visible = False
            TxtObs.Visible = False
            cmdDesaprueba.Visible = True
            CmdObs.Visible = False
            Mensaje = "Aprobado"
       End If
       If ADO_M1.Recordset!estado_registro = "N" Or ADO_M1.Recordset!estado_registro = "" Then
            cmdAprueba.Visible = True
            CmdCopiar.Visible = True
            CmdDel.Visible = True
            CmdMod.Visible = True
            LblObs.Visible = False
            TxtObs.Visible = False
            cmdDesaprueba.Visible = False
            CmdObs.Visible = True
            Mensaje = "No revisado"
       End If
    
    End If
'    If GlSW = "MOD" Then
'        var_DocId = rs_M1!pac_doc_id
'        rs_consultor.Find "pac_doc_id = '" & Trim(var_DocId) & "' ", , adSearchForward
'        If Not rs_consultor.EOF Then
'            txttimeEmb.Enabled = IIf(IsNull(adopuestosol.Recordset("materno")) = True, " ", adopuestosol.Recordset("materno"))
'        End If
'        Ado_consultor.Recordset.Find "pac_doc_id = '" & Trim(var_DocId) & "' ", , adSearchForward
'        If Not Ado_consultor.Recordset.EOF Then
'            txttimeEmb.Enabled = IIf(Ado_consultor.Recordset("genero_codigo") = "F", True, False)
'
'        End If
'        If Dtc_Genero.Text = "F" Then
'            txttimeEmb.Enabled = False
'        Else
'            txttimeEmb.Enabled = True
'        End If
'    End If
  If ADO_M1.Recordset.RecordCount > 0 Then
    ADO_M1.Caption = ADO_M1.Recordset.AbsolutePosition & " / " & ADO_M1.Recordset.RecordCount & " --> " & Mensaje
  End If
End Sub

Private Sub Chk_iniT_no_Click()
    Txt_ini_tratam = "NO"
    Chk_iniT_si.Value = 0
    Chk_iniT_ns.Value = 0
End Sub

Private Sub Chk_iniT_ns_Click()
    Txt_ini_tratam = "NS"
    Chk_iniT_no.Value = 0
    Chk_iniT_si.Value = 0
End Sub

Private Sub Chk_iniT_si_Click()
    Txt_ini_tratam = "SI"
    Chk_iniT_no.Value = 0
    Chk_iniT_ns.Value = 0
End Sub

Private Sub Chk_finT_no_Click()
    Txt_fin_tratam = "NO"
    'Chk_finT_no.Value = 1
    Chk_finT_si.Value = 0
    Chk_finT_ns.Value = 0
End Sub

Private Sub Chk_finT_ns_Click()
    Txt_fin_tratam = "NS"
    'Chk_finT_ns.Value = 1
    Chk_finT_si.Value = 0
    Chk_finT_no.Value = 0
End Sub

Private Sub Chk_finT_si_Click()
    Txt_fin_tratam = "SI"
    'Chk_finT_si.Value = 1
    Chk_finT_no.Value = 0
    Chk_finT_ns.Value = 0
End Sub

Private Sub Chk_sanoT_no_Click()
    Txt_sano_pac.Text = "NO"
    Chk_sanoT_si.Value = 0
    Chk_sanoT_ns.Value = 0
End Sub

Private Sub Chk_sanoT_ns_Click()
    Txt_sano_pac.Text = "NS"
    Chk_sanoT_no.Value = 0
    Chk_sanoT_si.Value = 0
End Sub

Private Sub Chk_sanoT_si_Click()
    Txt_sano_pac.Text = "SI"
    Chk_sanoT_no.Value = 0
    Chk_sanoT_ns.Value = 0
End Sub

Private Sub Chk_AT_Click()
    If Chk_AT.Value = 1 Then
        Txt_AlternativaOtro.Visible = True
    Else
        Txt_AlternativaOtro.Visible = False
    End If
End Sub

Private Sub Cmd_comuNew_Click()
  If DtcEmpId.Text = "" Then
    glPersOtro = DtcEmpId.Text
    GlTP = "6"
    frmBeneficiario.Show
  Else
    glPersOtro = DtcEmpId.Text
    GlTP = "6"
    frmBeneficiario_Consulta.Caption = "DATOS DE LA UNIDAD PRODUCTIVA"
    frmBeneficiario_Consulta.Show vbModal
  End If
End Sub

Private Sub Cmd_comuNew_LostFocus()
    rs_comunid.Requery
End Sub

Private Sub Cmd_LabNuevo_LostFocus()
    rs_persona.Requery
End Sub

Private Sub ChkEmpTest_Click(Area As Integer)
'    DtcEmpId.BoundText = ChkEmpTest.BoundText
'    DtcEmpDes.BoundText = ChkEmpTest.BoundText
'    DtcEmpLegPat.BoundText = ChkEmpTest.BoundText
'    DtcEmpLegNom.BoundText = ChkEmpTest.BoundText
'    DtcEmpLeg.BoundText = ChkEmpTest.BoundText
'    DtcEmpAct.BoundText = ChkEmpTest.BoundText
'    DtcEmpAsoc.BoundText = ChkEmpTest.BoundText
'    DtcEmpRenca.BoundText = ChkEmpTest.BoundText
'    DtcEmpFech.BoundText = ChkEmpTest.BoundText
'    DtcEmpNit.BoundText = ChkEmpTest.BoundText
'    DtcEmpDepto.BoundText = ChkEmpTest.BoundText
'    DtcEmpProv.BoundText = ChkEmpTest.BoundText
'    DtcEmpMunic.BoundText = ChkEmpTest.BoundText
'    DtcEmpComun.BoundText = ChkEmpTest.BoundText
'    DtcEmpDir.BoundText = ChkEmpTest.BoundText
'    DtcEmpZona.BoundText = ChkEmpTest.BoundText
'    DtcEmpTelef.BoundText = ChkEmpTest.BoundText
'    DtcEmpEmail.BoundText = ChkEmpTest.BoundText
'    DtcEmpDepto2.BoundText = ChkEmpTest.BoundText
'    DtcEmpProv2.BoundText = ChkEmpTest.BoundText
'    DtcEmpMunic2.BoundText = ChkEmpTest.BoundText
'    DtcEmpComun2.BoundText = ChkEmpTest.BoundText
'    DtcEmpDir2.BoundText = ChkEmpTest.BoundText
End Sub

Private Sub Cmd_Nuevo_Pac_Click()
    marca1 = adosolicitud.Recordset.Bookmark
    Frmmo_proyUbic.Lblformulario = "F_A"
    Frmmo_proyUbic.lblges_gestion = ADO_M1.Recordset!gestion
    Frmmo_proyUbic.LblFA = adosolicitud.Recordset!Numero_FA
    'Frmmo_proyUbic.lblcodigo_solicitud = adosolicitud.Recordset!correl_ubicacion
    'Frmmo_proyUbic.lbltipo_beneficiario = "-"       'adosolicitud.Recordset!tipo_beneficiario
    Frmmo_proyUbic.Show vbModal
    'Call OptFilGral1_Click
End Sub

Private Sub Cmd_PersNuevo_Click()
  If DtcRespId.Text = "" Then
    glPersOtro = DtcRespId.Text
    GlTP = "2"
    frmBeneficiario.Show
  Else
    glPersOtro = DtcRespId.Text
    GlTP = "2"
    frmBeneficiario_Consulta.Caption = "DATOS DEL CONSULTOR AMBIENTAL"
    frmBeneficiario_Consulta.Show vbModal
  End If
End Sub

Private Sub Cmd_PersNuevo2_Click()
  If DtcRespId.Text = "" Then
    glPersOtro = DtcPromId.Text
    GlTP = "1"
    frmBeneficiario.Show
  Else
    glPersOtro = DtcPromId.Text
    GlTP = "1"
    frmBeneficiario_Consulta.Caption = "DATOS DEL PROMOTOR"
    frmBeneficiario_Consulta.Show vbModal
  End If
End Sub

Private Sub cmdAdd3_Click()
 sino = MsgBox("Desea Adicionar un NUEVO Registro ?", vbQuestion + vbYesNo, "Confirmando...")
  If sino = vbYes Then
'    Call Abre_Sol_Bien
   If ADO_M1.Recordset.RecordCount > 0 Then
      If ADO_M1.Recordset!estado_registro = "N" Then
        swgraba3 = 0
        DtgProyUbica.AllowAddNew = False
        DtgProyUbica.AllowDelete = False
        DtgProyUbica.AllowUpdate = True
'        marca1 = ADO_M1.Recordset.Bookmark
'        rs_ProyUbic.AddNew
'        rs_ProyUbic!gestion = ADO_M1.Recordset!gestion
'        rs_ProyUbic!Numero_FA = ADO_M1.Recordset!Numero_FA
'        rs_ProyUbic!comun_codigo = "-"
'        rs_ProyUbic!depto_codigo = "-"
'        rs_ProyUbic!prov_codigo = "-"
'        rs_ProyUbic!munic_codigo = "-"
'        rs_ProyUbic.Update
'        rs_ProyUbic.MoveLast
'        Call OptFilGral1_Click
        'adosolicitud.Recordset.BookMark = marca1
        'adosolicitud.Refresh
        Dim rs_U_Fisisca As New ADODB.Recordset
        Set rs_U_Fisisca = New ADODB.Recordset
        If rs_U_Fisisca.State = 1 Then rs_U_Fisisca.Close
        rs_U_Fisisca.Open "select * from mo_proy_Id_Ubicacion where Gestion = '" & ADO_M1.Recordset!gestion & "' and Numero_FA = " & ADO_M1.Recordset!Numero_FA & " and comun_codigo = '-'  ", DB, adOpenDynamic, adLockOptimistic
        If rs_U_Fisisca.RecordCount > 0 Then
            MsgBox "No se puede Adicionar un NUEVO registro, mientras exista otro PENDIENTE !!...", vbInformation, "Formulario 04"
            Exit Sub
        Else
            DB.Execute "INSERT INTO mo_proy_Id_Ubicacion (gestion, Numero_FA, comun_codigo, depto_codigo, prov_codigo, munic_codigo, depto, prov, munic, comun, latitud, longitud, altitud_snm, estado_registro, usr_codigo) VALUES ('" & ADO_M1.Recordset!gestion & "', " & ADO_M1.Recordset!Numero_FA & ", '-', '-', '-', '-','-', '-', '-', '-', '0', '0', '0', 'S', '" & GlUsuario & "') "
            
            Call abrirtabla_maestra
        End If
        If rs_U_Fisisca.State = 1 Then rs_U_Fisisca.Close
        
'        cmdAdd3.Enabled = False
        CmdMod3.Enabled = False
        Command1.Enabled = False
        Command2.Enabled = False
'        Command3.Enabled = False
'        Command4.Enabled = False
'        Command5.Enabled = False
'        Command6.Enabled = False
'        Command7.Enabled = False
'        Command8.Enabled = False
'        Command9.Enabled = False
'        Command10.Enabled = False
'        Command11.Enabled = False
'        Command12.Enabled = False
'        Command13.Enabled = False
'        Command14.Enabled = False
'        Command15.Enabled = False
'        Command16.Enabled = False
'        cmdDel3.Enabled = False
'        cmdGraba3.Enabled = True
      Else
         MsgBox "No se puede editar un registro APROBADO o ANULADO ", vbInformation, "Formulario 1"
      End If
   Else
          MsgBox "No Existen Registros habilitados ", vbInformation, "Formulario 1"
   End If

  End If
        
End Sub

Private Sub cmdAprueba_Click()
  On Error GoTo UpdateErr

   If ADO_M1.Recordset!estado_registro = "N" Or ADO_M1.Recordset!estado_registro = "" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        ADO_M1.Recordset!estado_registro = "A"
        ADO_M1.Recordset!fecha_aprueba = Date
        ADO_M1.Recordset!usr_codigo_apr = GlUsuario
        ADO_M1.Recordset.UpdateBatch adAffectAll
      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub CmdBuscar_Click()
    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexión = DB
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = DtG_M1
    ClBuscaGrid.QueryUtilizado = queryinicial
    Set ClBuscaGrid.RecordsetTrabajo = rs_M1
    'ClBuscaGrid.CamposVisibles = "11010011"
    ClBuscaGrid.Ejecutar
End Sub

Private Sub cmdCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
      rs_M1.CancelUpdate
      If GlSW = "ADD" Then
        Call abrirtabla_maestra
      Else
        If mvBookMark > 0 Then
          rs_M1.Bookmark = mvBookMark
        Else
          rs_M1.MoveFirst
        End If
      End If
    '  mbDataChanged = False
      FraCabecera.Enabled = False
      FraEmpresa.Enabled = False
      FraProyecto.Enabled = False
'      FraSitio.Enabled = False
'      FraDescripProy.Enabled = False
'      FraAlterTec.Enabled = False
'      FraInversion.Enabled = False
'      FraActividades.Enabled = False
'      FraRRHH.Enabled = False
'      FraRRNN.Enabled = False
'      FraMatPrima.Enabled = False
'      FraDesecho.Enabled = False
'      FraRuido.Enabled = False
'      FraAlmInsumo.Enabled = False
'      FraTranspInsumo.Enabled = False
'      FraContingencia.Enabled = False
'      FraImpactos.Enabled = False
      FraDeclarJur.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      DtG_M1.Enabled = True
      lblStatus.Caption = "."
'      cmdAdd3.Enabled = False
      CmdMod3.Enabled = False
      Command1.Enabled = False
    Command2.Enabled = False
'    Command3.Enabled = False
'    Command4.Enabled = False
'    Command5.Enabled = False
'    Command6.Enabled = False
'    Command7.Enabled = False
'    Command8.Enabled = False
'    Command9.Enabled = False
'    Command10.Enabled = False
'    Command11.Enabled = False
'    Command12.Enabled = False
'    Command13.Enabled = False
'    Command14.Enabled = False
'    Command15.Enabled = False
'    Command16.Enabled = False
'      cmdDel3.Enabled = False
'      cmdGraba3.Enabled = False
'      trv.Enabled = True
   End If
End Sub

Private Sub CmdCopiar_Click()
 On Error GoTo Error_Sub
    If ADO_M1.Recordset!Numero_FA < 10 Then
       NombreCarpeta = App.Path & "\FA\FA-" & ADO_M1.Recordset!gestion & "-00" & ADO_M1.Recordset!Numero_FA
       'e = "\\SERVIDOR\users\public\documents\FA\FA-" & ADO_M1.Recordset!gestion & "-00" & ADO_M1.Recordset!Numero_FA
       e = "\\SERVIDOR\Sistema\FA\FA-" & ADO_M1.Recordset!gestion & "-00" & ADO_M1.Recordset!Numero_FA
       'e = "\\DMA-196\Sistema\FA\FA-" & ADO_M1.Recordset!gestion & "-00" & ADO_M1.Recordset!Numero_FA
    End If
    If ADO_M1.Recordset!Numero_FA > 9 And ADO_M1.Recordset!Numero_FA < 100 Then
       NombreCarpeta = App.Path & "\FA\FA-" & ADO_M1.Recordset!gestion & "-0" & ADO_M1.Recordset!Numero_FA
       'e = "\\SERVIDOR\users\public\documents\FA\FA-" & ADO_M1.Recordset!gestion & "-0" & ADO_M1.Recordset!Numero_FA
       e = "\\SERVIDOR\Sistema\FA\FA-" & ADO_M1.Recordset!gestion & "-0" & ADO_M1.Recordset!Numero_FA
       'e = "\\DMA-196\Sistema\FA\FA-" & ADO_M1.Recordset!gestion & "-0" & ADO_M1.Recordset!Numero_FA
    End If
    If ADO_M1.Recordset!Numero_FA > 99 And ADO_M1.Recordset!Numero_FA < 1000 Then
       NombreCarpeta = App.Path & "\FA\FA-" & ADO_M1.Recordset!gestion & "-" & ADO_M1.Recordset!Numero_FA
       'e = "\\SERVIDOR\users\public\documents\FA\FA-" & ADO_M1.Recordset!gestion & "-" & ADO_M1.Recordset!Numero_FA
       e = "\\SERVIDOR\Sistema\FA\FA-" & ADO_M1.Recordset!gestion & "-" & ADO_M1.Recordset!Numero_FA
       'e = "\\DMA-196\Sistema\FA\FA-" & ADO_M1.Recordset!gestion & "-" & ADO_M1.Recordset!Numero_FA
    End If
'    Mensaje = NombreCarpeta
'    Call Eliminar_Directorio(NombreCarpeta)
'    Mensaje = e
'    Call Eliminar_Directorio(e)
    Frmexporta.DirDestino.Path = NombreCarpeta
    Frmexporta.DirDestino2.Path = e
    Frmexporta.Show vbModal
    'MsgBox "Coloque el CD, para volver a COPIAR su contenido ... ", vbCritical + vbExclamation, "Realiza la Copia de CD"
    'sino = MsgBox("Desea Borrar los datos copiados anteriormente en su computadora ? ", vbYesNo + vbQuestion, "Atención")
    'If sino = vbYes Then
    '    Kill NombreCarpeta & "\*.*"
    '    Kill e & "\*.*"
    '    My.Computer.FileSystem.DeleteFile (NombreCarpeta & "\*.*")
    '    'My.Computer.FileSystem.DeleteFile(NombreCarpeta & "\*.*", FileIO.UIOption.AllDialogs, FileIO.RecycleOption.DeletePermanently, FileIO.UICancelOption.DoNothing)

    '    'MkDir NombreCarpeta
    '    'MkDir e
    'End If
    'Set fs = CreateObject("Scripting.FileSystemObject")
    'fs.CopyFile "G:\*.*", NombreCarpeta
    'fs.CopyFile "G:\*.*", e
    'fs.CopyFile "F:\WIN\*.*", NombreCarpeta
    'fs.CopyFile "F:\COPIA\*.*", e
 Exit Sub
Error_Sub:
 MsgBox Err.Description, vbCritical
End Sub

Function Eliminar_Directorio(Path As String) As Boolean
' función que borra la carpeta
 On Error GoTo Error_Sub
   sino = MsgBox("Desea Borrar los datos copiados anteriormente en -->" & Mensaje, vbYesNo + vbQuestion, "Atención")
   If sino = vbYes And Err.Number = 0 Then
        Dim fso As FileSystemObject      'Variable de tipo file System Object
        Set fso = New FileSystemObject   'Creamos la Nueva referencia Fso
        'fso.DeleteFolder Path, True      'Le pasamos a DeleTeFolder el Path a eliminar
        fso.DeleteFile "*.*", True          'Le pasamos a DeleTeFolder el Path a eliminar
        ' Ok
        Eliminar_Directorio = True
        Set fso = Nothing
   End If
 Exit Function
Error_Sub:
 MsgBox Err.Description, vbCritical
End Function

'Private Sub Command1_Click()
'
'If Text1 <> "" Then
'        ' Msgbox de Confirmación de eliminación
'        If MsgBox("Seguro que se quiere borrar el directorio " & _
'                  "indicado ??", vbQuestion + vbYesNo) = vbYes Then
'
'             ' elimina la carpeta
'If Eliminar_Directorio(Trim(Text1)) Then
'                 MsgBox "Directorio eliminado", vbInformation
'End If
'End If
'End If
'End Sub

'Private Sub Form_Load()
'     Command1.Caption = " Eliminar "
'	Call SeguridadSet(Me)
End Sub

Private Sub CmdDel_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de RECHAZAR el Registro? ", vbYesNo + vbQuestion, "Atención")
   If ADO_M1.Recordset!estado_registro = "N" Or ADO_M1.Recordset!estado_registro = "" Then
      If sino = vbYes Then
        ADO_M1.Recordset!estado_registro = "L"
        ADO_M1.Recordset!fecha_modifica = Date
        ADO_M1.Recordset!usr_codigo_mod = GlUsuario
        ADO_M1.Recordset.Update     'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede RECHAZAR un registro APROBADO ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDesaprueba_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If ADO_M1.Recordset!estado_registro = "S" Then
      If sino = vbYes Then
        ADO_M1.Recordset!estado_registro = "N"
        ADO_M1.Recordset!fecha_aprueba = Date
        ADO_M1.Recordset!usr_codigo_apr = GlUsuario
        ADO_M1.Recordset.UpdateBatch adAffectAll
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Elaborado o Anulado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdGraba3_Click()
'--------------
  If ADO_M1.Recordset!estado_registro <> "S" And swgraba3 <> 2 Then
   'If adosolicitud1.Recordset.RecordCount > 0 And Not IsNull(DataGrid3.Columns("NombreCta").Value) And (DataGrid3.Columns("NombreCta").Value) <> "" Then
    If Not IsNull(DtgProyUbica.Columns("munic").Value) And (DataGrid3.Columns("munic").Value) <> "" Then
        VARC = DtgProyUbica.Columns("Gestion").Value
        VARS1 = DtgProyUbica.Columns("Numero_FA").Value
        VARS2 = DtgProyUbica.Columns("comun_codigo").Value
        Dim rs_U_FIS2 As New ADODB.Recordset
        Set rs_U_FIS2 = New ADODB.Recordset
        If rs_U_FIS2.State = 1 Then rs_U_FIS2.Close
        tot_form = 0
        rs_U_FIS2.Open "select COUNT(*) AS tot_form from mo_proy_Id_Ubicacion WHERE Gestion = '" & VARC & "' and Numero_FA = '" & VARS1 & "' and comun_codigo = '" & VARS2 & "' ", DB, adOpenDynamic, adLockOptimistic
        If rs_U_FIS2!tot_form > swgraba3 Then
            MsgBox "No se puede Guardar un registro ya EXISTENTE, verifique por favor !!...", vbInformation, "Formulario"
            DataGrid3.SetFocus
            Exit Sub
        Else
            If rs_U_FIS2.State = 1 Then rs_U_FIS2.Close
            'marca1 = adosolicitud1.Recordset.Bookmark
            VARS2 = DtgProyUbica.Columns("comun_codigo").Value
            VARA1 = DtgProyUbica.Columns("depto_codigo").Value
            VARA2 = DtgProyUbica.Columns("prov_codigo").Value
            VARA3 = DtgProyUbica.Columns("munic_codigo").Value
            VARAA1 = DtgProyUbica.Columns("depto").Value
            VARAA2 = DtgProyUbica.Columns("prov").Value
            VARAA3 = DtgProyUbica.Columns("munic").Value
            VARNC = DtgProyUbica.Columns("comun").Value
            VARDB = DtgProyUbica.Columns("latitud").Value
            VARHB = DtgProyUbica.Columns("longitud").Value
            VARCA = DtgProyUbica.Columns("altitud_snm").Value
'            Call Abre_Balance
    '        'MarcaB = rstAo_solicitud1.Bookmark
            'rstAo_solicitud1.Bookmark = marca1
            DB.Execute "UPDATE mo_proy_Id_Ubicacion SET Gestion = '" & VARC & "' , Numero_FA = '" & VARS1 & "' , comun_codigo = '" & VARS2 & "' , comun_codigo = '" & VARS2 & "' , depto_codigo = '" & VARA1 & "' , prov_codigo = '" & VARA2 & "' , munic_codigo = '" & VARA3 & "' , depto = '" & VARAA1 & "' , prov = '" & VARAA2 & "' , munic = '" & VARAA3 & "' , comun = '" & VARNC & "' , latitud = '" & VARDB & "' , longitud = '" & VARHB & "' , altitud_snm = '" & VARCA & "' "
            'db.Execute "UPDATE ao_Solicitud_detalle SET ao_Solicitud_detalle.monto_bolivianos = (SELECT SUM(Total_venta) FROM ao_solicitud_bien WHERE ao_solicitud_bien.CODIGO_UNIDAD = ao_Solicitud_detalle.CODIGO_UNIDAD AND ao_solicitud_bien.codigo_solicitud = ao_Solicitud_detalle.codigo_solicitud AND ao_solicitud_bien.cod_univ = ao_Solicitud_detalle.par_codigo) FROM ao_Solicitud_detalle, ao_solicitud_bien"
            swgraba3 = 2
            marca1 = rs_M1.Bookmark
            Call abrirtabla_maestra
            rs_M1.Bookmark = marca1
            'rstAo_solicitud1.MoveLast
'            Frame1.Enabled = False
'            Frame4.Visible = False
'            frmabm.Visible = True
'            FrmGraba.Visible = False
'            CmdEnviar.Visible = False
'            CmdGraCabeza.Visible = True
'
'            Call Limpia_combos
                    
            DtgProyUbica.AllowAddNew = False
            DtgProyUbica.AllowDelete = False
            DtgProyUbica.AllowUpdate = False
        End If
      'Else
      '   MsgBox "No se puede modificar un registro APROBADO ", vbInformation, "Formulario 04"
      'End If
    Else
         MsgBox "Verifique los datos para continuar ... ", vbInformation, "Formulario"
    End If
 Else
    MsgBox "ERROR, NO se puede modificar un registro aprobado..."
 End If

'--------------
'     If swnuevo = 1 Then
'        Set rstdestino = New ADODB.Recordset
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from mo_proy_Id_Ubicacion where Gestion = '" & lblges_gestion & "' and Numero_FA = " & lblFA & " and comun_codigo = " & lblCorrel & " ", db, adOpenKeyset, adLockOptimistic
'        If rstdestino.RecordCount > 0 Then
''        Dim rstid_beneficiario As New ADODB.Recordset
''        Set rstid_beneficiario = New ADODB.Recordset
''        If rstid_beneficiario.State = 1 Then rstid_beneficiario.Close
''        rstid_beneficiario.Open "select Max(id_beneficiario) as max_id from ao_solicitud_lista where codigo_unidad = '" & lblFA & "' and codigo_solicitud = " & lblcodigo_solicitud, db, adOpenKeyset, adLockOptimistic
''        id_beneficiario1 = rstid_beneficiario!max_id + 1
''        If rstid_beneficiario.State = 1 Then rstid_beneficiario.Close
'            MsgBox "Error el registro ya existe... Vuelva a intentar ..."
'            Exit Sub
'        'Else
''        id_beneficiario1 = 1
'        End If
''      rstdestino.AddNew
'        ado_proy_ubicacion.Recordset!gestion = lblges_gestion.Caption      'Year(Date)
'        ado_proy_ubicacion.Recordset!Numero_FA = Trim(lblFA.Caption)
''        ado_proy_ubicacion.Recordset!correl_ubicacion = CDbl(Dtc_local_cod.Text)
'        ado_proy_ubicacion.Recordset!depto_codigo = IIf(Dtc_depto_cod.Text = "", "-", Dtc_depto_cod.Text)
'        ado_proy_ubicacion.Recordset!prov_codigo = IIf(Dtc_prov_cod.Text = "", "-", Dtc_prov_cod.Text)
'        ado_proy_ubicacion.Recordset!munic_codigo = IIf(Dtc_munic_cod.Text = "", "-", Dtc_munic_cod.Text)
'        ado_proy_ubicacion.Recordset!comun_codigo = IIf(Dtc_local_cod.Text = "", "-", Dtc_local_cod.Text)
'        ado_proy_ubicacion.Recordset!depto = IIf(Dtc_depto.Text = "", "NO TIENE", Dtc_depto.Text)
'        ado_proy_ubicacion.Recordset!prov = IIf(Dtc_prov.Text = "", "NO TIENE", Dtc_prov.Text)
'        ado_proy_ubicacion.Recordset!munic = IIf(Dtc_munic.Text = "", "NO TIENE", Dtc_munic.Text)
'        ado_proy_ubicacion.Recordset!comun = IIf(Dtc_local.Text = "", "NO TIENE", Dtc_local.Text)
'
''        rs_proy_Ubicacion!ges_gestion = lblges_gestion.Caption      'Year(Date)
''        rs_proy_Ubicacion!Numero_FA = Trim(lblFA.Caption)
''        rs_proy_Ubicacion!depto_codigo = Dtc_depto_cod.Text
''        rs_proy_Ubicacion!prov_codigo = Dtc_prov_cod.Text
''        rs_proy_Ubicacion!munic_codigo = Dtc_munic_cod.Text
''        rs_proy_Ubicacion!comun_codigo = Dtc_local_cod.Text
''        rs_proy_Ubicacion!depto = Dtc_depto.Text
''        rs_proy_Ubicacion!prov = Dtc_prov.Text
''        rs_proy_Ubicacion!munic = Dtc_munic.Text
''        rs_proy_Ubicacion!comun = Dtc_local.Text
'     End If
'     If swnuevo = 2 Then
'        marcaL = ado_proy_ubicacion.Recordset.Bookmark
'     End If
''     rs_proy_Ubicacion!num_habitantes = txtHabitantes.Text
''     rs_proy_Ubicacion!num_flias = txtFlias.Text
''     rs_proy_Ubicacion!latitud = TxtLatitud.Text
''     rs_proy_Ubicacion!longitud = TxtLongitud.Text
''     rs_proy_Ubicacion!altitud_snm = TxtAltitud.Text
''     rs_proy_Ubicacion!usr_codigo = GlUsuario
''     rs_proy_Ubicacion!fecha_registro = Format(Date, "dd/mm/yyyy")
''     rs_proy_Ubicacion!hora_registro = Format(Time, "HH:mm:ss")
''     rs_proy_Ubicacion.Update
'
'     ado_proy_ubicacion.Recordset!num_habitantes = IIf(txtHabitantes.Text = "", 0, txtHabitantes.Text)
'     ado_proy_ubicacion.Recordset!num_flias = IIf(txtFlias.Text = "", 0, txtFlias.Text)
'     ado_proy_ubicacion.Recordset!latitud = IIf(TxtLatitud.Text = "", 0, TxtLatitud.Text)
'     ado_proy_ubicacion.Recordset!longitud = IIf(TxtLongitud.Text = "", 0, TxtLongitud.Text)
'     ado_proy_ubicacion.Recordset!altitud_snm = IIf(TxtAltitud.Text = "", 0, TxtAltitud.Text)
'     ado_proy_ubicacion.Recordset!usr_codigo = GlUsuario
'     ado_proy_ubicacion.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
'     ado_proy_ubicacion.Recordset!hora_registro = Format(Time, "HH:mm:ss")
'     ado_proy_ubicacion.Recordset.Update
'     db.CommitTrans
'  End If

End Sub

Private Sub CmdGrabar_Click()
  On Error GoTo UpdateErr
   VAR_VAL = "OK"
   Call valida_campos
   If VAR_VAL = "OK" Then
        'GlParametro = "CVM00001"
      If GlSW = "ADD" Then
         var_form = IIf(IsNull(DtCForm.Text), "FA", DtCForm.Text)
         Set rs_correlativo = New ADODB.Recordset
         rs_correlativo.Open "select * from gc_correlativos WHERE PARAM_CODIGO = '" & txtParam & "' and form_codigo = '" & var_form & "' ", DB, adOpenKeyset, adLockOptimistic
         If rs_correlativo.RecordCount > 0 Then
            rs_correlativo!correlativo = rs_correlativo!correlativo + 1
            rs_correlativo.Update
            rs_M1!Numero_FA = rs_correlativo!correlativo
         Else
            MsgBox "Error en Parametros de Correlativo .... "
         End If
         rs_M1!gestion = txtParam.Text
        'Dim e As Long
        'e = Shell(App.Path & "\Reportes Tesoreria\cnsPagados.exe", 1)
        
        '    Dim NombreCarpeta As String
        '    NombreCarpeta = "c:/prueba"
        '    Call Shell("mk " & NombreCarpeta)
        
        'FrmExplora.lblges_gestion = ADO_M1.Recordset!gestion
        'FrmExplora.lblFA = ADO_M1.Recordset!Numero_FA
        ''Dir1.Path = App.Path & "\FA-" & lblges_gestion & "-00" & LblFA
        'FrmExplora.Label1 = App.Path & "\FA-" & FrmExplora.lblges_gestion & "-00" & FrmExplora.lblFA
        
'        Dim fs, f
'        Set fs = CreateObject("Scripting.FileSystemObject")
'        Set f = fs.CreateFolder("X:\ruta_carpeta\")

'        Try
'        If Directory.Exists(Path & "\" & TextBox1.Text) Then
'
'        a = MsgBox("El directorio ya existe")
'        Else
'        di = Directory.CreateDirectory(Path & "\" & TextBox1.Text)
'        b = MsgBox("El directorio fue creado exitosamente")
'        End If
'        Finally
'        End Try

'        Public Function CreaDir(ValDir As String)
'        Dim AttrDev%
'        On Error Resume Next
'        AttrDev = GetAttr(ValDir)
'          If Err.Number Then
'          Err.Clear
'          MkDir ValDir
'          End If
'        End Function
         
         If rs_M1!Numero_FA < 10 Then
            NombreCarpeta = App.Path & "\FA\FA-" & rs_M1!gestion & "-00" & rs_M1!Numero_FA
            'e = "\\SERVIDOR\users\public\documents\FA\FA-" & rs_M1!gestion & "-00" & rs_M1!Numero_FA
            e = "\\SERVIDOR\Sistema\FA\FA-" & rs_M1!gestion & "-00" & rs_M1!Numero_FA
            'e = "\\DMA-196\Sistema\FA\FA-" & rs_M1!gestion & "-00" & rs_M1!Numero_FA
         End If
         If rs_M1!Numero_FA > 9 And rs_M1!Numero_FA < 100 Then
            NombreCarpeta = App.Path & "\FA\FA-" & rs_M1!gestion & "-0" & rs_M1!Numero_FA
            'e = "\\SERVIDOR\users\public\documents\FA\FA-" & rs_M1!gestion & "-0" & rs_M1!Numero_FA
            e = "\\SERVIDOR\Sistema\FA\FA-" & rs_M1!gestion & "-0" & rs_M1!Numero_FA
            'e = "\\DMA-196\Sistema\FA\FA-" & rs_M1!gestion & "-0" & rs_M1!Numero_FA
         End If
         If rs_M1!Numero_FA > 99 And rs_M1!Numero_FA < 1000 Then
            NombreCarpeta = App.Path & "\FA\FA-" & rs_M1!gestion & "-" & rs_M1!Numero_FA
            'e = "\\SERVIDOR\users\public\documents\FA\FA-" & rs_M1!gestion & "-" & rs_M1!Numero_FA
            e = "\\SERVIDOR\Sistema\FA\FA-" & rs_M1!gestion & "-" & rs_M1!Numero_FA
            'e = "\\DMA-196\Sistema\FA\FA-" & rs_M1!gestion & "-" & rs_M1!Numero_FA
         End If
         MkDir NombreCarpeta
         MkDir e
         ''Kill "E:\*.*"
         'MsgBox "Coloque el CD, para copiar su contenido ... ", vbCritical + vbExclamation, "Realiza la Copia de CD"
         'Set fs = CreateObject("Scripting.FileSystemObject")
         'fs.CopyFile "G:\*.*", NombreCarpeta
         'fs.CopyFile "G:\*.*", e
         Frmexporta.DirDestino.Path = NombreCarpeta
         'Frmexporta.DirDestino2.Path = NombreCarpeta
         Frmexporta.DirDestino2.Path = e
         Frmexporta.Show vbModal

         'FileCopy "E:", NombreCarpeta
         'FileCopy "E:", e
         'FileCopy NombreCarpeta, e
         'Shell ("mkdir " & NombreCarpeta)
         'e = NombreCarpeta
      End If
   rs_M1!tipo_formulario = DtCForm.Text
   rs_M1!fecha_Llenado = DTP_FechaLectura.Value
   rs_M1!depto_codigo = Dtc_depto_cod.Text
   rs_M1!prov_codigo = Dtc_prov_cod.Text
   rs_M1!munic_codigo = Dtc_munic_cod.Text
   rs_M1!codigo_Promotor = DtcPromId.Text
   rs_M1!codigo_responsable = DtcRespId.Text
   rs_M1!codigo_empresa = DtcEmpId.Text
   rs_M1!Proy_Descripcion = TxtProyNom.Text
   rs_M1!catastro_codigo_predio = IIf(TxtProyCodCat.Text = "", "", TxtProyCodCat.Text)
   rs_M1!catastro_No_Reg = IIf(TxtProyRegCat.Text = "", 0, TxtProyRegCat.Text)
   rs_M1!DDRR_Partida = TxtProyDDRRPar.Text
   rs_M1!DDRR_Fojas = TxtProyDDRRFoj.Text
   rs_M1!DDRR_Libro = TxtProyDDRRLib.Text
   rs_M1!DDRR_Anio = TxtProyDDRRAnio.Text
   rs_M1!DDRR_depto_codigo = IIf(DtcProyDptoCod.Text = "", "-", DtcProyDptoCod.Text)
   rs_M1!Colindante_Norte = TxtProyColiN.Text
   rs_M1!Colindante_Sur = TxtProyColiS.Text
   rs_M1!Colindante_Este = TxtProyColiE.Text
   rs_M1!Colindante_Oeste = TxtProyColiO.Text
   rs_M1!Act_Colind_Norte = TxtProyActN.Text
   rs_M1!Act_Colind_Sur = TxtProyActS.Text
   rs_M1!Act_Colind_Este = TxtProyActE.Text
   rs_M1!Act_Colind_Oeste = TxtProyActO.Text
   rs_M1!uso_suelo_codigo_Actual = IIf(DtcProyUsoSC.Text = "", 0, DtcProyUsoSC.Text)
   rs_M1!uso_suelo_codigo_Potencial = IIf(DtcProyUsoSC2.Text = "", 0, DtcProyUsoSC2.Text)
   rs_M1!uso_Suelo_Certificado_Nro = IIf(DtcProyUsoSNo.Text = "", 0, DtcProyUsoSNo.Text)
   rs_M1!uso_Suelo_Certif_Expedido_en = DtcProyUsoSExp.Text
   rs_M1!uso_Suelo_Certif_Fecha = DtcProyUsoSFech.Value
   rs_M1!sitio_Superf_Total = IIf(TxtSupTot.Text = "", 0, TxtSupTot.Text)
   rs_M1!sitio_Superf_Ocupada = IIf(TxtSupOcu.Text = "", 0, TxtSupOcu.Text)
   rs_M1!Sitio_UniMed = IIf(DtcUniMed.Text = "", "-", DtcUniMed.Text)
   rs_M1!Sitio_Etop_codigo = IIf(Dtc_SitioTopo.Text = "", 0, Dtc_SitioTopo.Text)
   rs_M1!Sitio_ETop_Obs = Txt_SitioTopoObs.Text
   rs_M1!Sitio_Napa_codigo = IIf(Dtc_SitioNapa.Text = "", 0, Dtc_SitioNapa.Text)
   rs_M1!Sitio_Napa_Obs = Txt_SitioNapaObs.Text
   rs_M1!Sitio_Agua_calid_codigo = IIf(Dtc_SitioAgua.Text = "", 0, Dtc_SitioAgua.Text)
   rs_M1!Sitio_Agua_calid_Obs = Txt_SitioAguaObs.Text
   rs_M1!Sitio_Eveg_codigo = IIf(Dtc_SitioVeg.Text = "", 0, Dtc_SitioVeg.Text)
   rs_M1!Sitio_EVeg_Obs = Txt_SitioVegObs.Text
   rs_M1!Sitio_red_drenaje_codigo = IIf(Dtc_SitioDren.Text = "", 0, Dtc_SitioDren.Text)
   rs_M1!Sitio_red_drenaje_Obs = Txt_SitioDrenObs.Text
   rs_M1!tot_poblacion_beneficiaria = IIf(Txt_SitioPobl.Text = "", 0, Txt_SitioPobl.Text)
   rs_M1!sector_codigo = IIf(Dtc_Sector.Text = "", "-", Dtc_Sector.Text)
   rs_M1!subsector_codigo = IIf(Dtc_SubSector.Text = "", "-", Dtc_SubSector.Text)
   rs_M1!sector_act_codigo = IIf(Dtc_SectorAct.Text = "", "-", Dtc_SectorAct.Text)
   rs_M1!CIIU = IIf(Dtc_SectorAct.Text = "", "-", Dtc_SectorAct.Text)
   rs_M1!Naturaleza_Codigo = IIf(Dtc_ProyNat.Text = "", 0, Dtc_ProyNat.Text)
   rs_M1!Naturaleza_Obs = "-"
   rs_M1!Ambito_codigo = IIf(Dtc_ProyAmb.Text = "", 0, Dtc_ProyAmb.Text)
   rs_M1!Objetivo_Gral_Proy = Txt_ProyObjGral.Text
   rs_M1!Objetivos_Especificos = Txt_ProyObjEsp.Text
   rs_M1!proy_tipo_codigo = IIf(Dtc_ProyTipo.Text = "", 0, Dtc_ProyTipo.Text)
   rs_M1!Proy_otro_codigo = IIf(Dtc_Proy.Text = "", 0, Dtc_Proy.Text)
   rs_M1!Proy_Otro_Obs = Dtc_ProyDes.Text
   rs_M1!Vida_Util_Proy_Anio = IIf(Txt_ProyVUAnio.Text = "", 0, Txt_ProyVUAnio.Text)
   rs_M1!Vida_Util_Proy_Mes = IIf(Txt_ProyVUMes.Text = "", 0, Txt_ProyVUMes.Text)
   rs_M1!Produccion_Anual_Estimada = Txt_ProyProdAnual.Text
   rs_M1!Alternativa_localizacion = "-"
   rs_M1!Alternativa_localizacion_Desc = Txt_AlternativaOtro.Text
   rs_M1!Tecnologias_Descripcion = Txt_AlternativaOtro.Text
   rs_M1!fase_codigo = IIf(Dtc_ProyFase.Text = "", 0, Dtc_ProyFase.Text)
   rs_M1!Inversion_Total_Bs = IIf(Txt_InvBs.Text = "", 0, Txt_InvBs.Text)
   rs_M1!Inversion_Total_Sus = IIf(Txt_InvDol.Text = "", 0, Txt_InvDol.Text)
   rs_M1!Insumos_almacenaje = Txt_InsumoAlm.Text
   rs_M1!Insumos_Trasp_Manipulacion = Txt_InsumoManip.Text
   rs_M1!Posibles_Contingencias = Txt_Conting.Text
   If ChkPromotor.Value = 0 Then
        rs_M1!DJ_Promotor = "N"
   Else
        rs_M1!DJ_Promotor = "S"
   End If
   If ChkConsultor.Value = 0 Then
        rs_M1!DJ_Responsable = "N"
   Else
        rs_M1!DJ_Responsable = "S"
   End If
   If ChkReprLegal.Value = 0 Then
        rs_M1!DJ_RepLegal = "N"
   Else
        rs_M1!DJ_RepLegal = "S"
   End If
   If ChkOtraPers.Value = 0 Then
        rs_M1!DJ_Otro = "N"
   Else
        rs_M1!DJ_Otro = "S"
   End If
   If ChkTestimonio.Value = 0 Then
        rs_M1!Emp_Testimonio_Constitucion = "N"
   Else
        rs_M1!Emp_Testimonio_Constitucion = "S"
   End If
   If ChkProyPlano.Value = 0 Then
        rs_M1!Anexo3_Plano_Ubicacion = "N"
   Else
        rs_M1!Anexo3_Plano_Ubicacion = "S"
   End If
   If ChkProyCertUS.Value = 0 Then
        rs_M1!Anexo3_Certificado_Uso_Suel = "N"
   Else
        rs_M1!Anexo3_Certificado_Uso_Suel = "S"
   End If
   If ChkProyProp.Value = 0 Then
        rs_M1!Anexo3_Derecho_Propietario = "N"
   Else
        rs_M1!Anexo3_Derecho_Propietario = "S"
   End If
   If ChkProyFoto.Value = 0 Then
        rs_M1!Anexo3_Foto_Panoramica_Lugar = "N"
   Else
        rs_M1!Anexo3_Foto_Panoramica_Lugar = "S"
   End If
   If GlSW = "ADD" Then
        rs_M1!estado_registro = "N"
   End If
   rs_M1!fecha_registro = Date
   'rs_M1!hora_registro
   rs_M1!usr_codigo = GlUsuario
   rs_M1.Update
      'rs_M1.UpdateBatch adAffectAll
      FraCabecera.Enabled = False
      FraEmpresa.Enabled = False
      FraProyecto.Enabled = False
'      FraSitio.Enabled = False
'      FraDescripProy.Enabled = False
'      FraAlterTec.Enabled = False
'      FraInversion.Enabled = False
'      FraActividades.Enabled = False
'      FraRRHH.Enabled = False
'      FraRRNN.Enabled = False
'      FraMatPrima.Enabled = False
'      FraDesecho.Enabled = False
'      FraRuido.Enabled = False
'      FraAlmInsumo.Enabled = False
'      FraTranspInsumo.Enabled = False
'      FraContingencia.Enabled = False
'      FraImpactos.Enabled = False
      FraDeclarJur.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      DtG_M1.Enabled = True
      lblStatus.Caption = "."
'      cmdAdd3.Enabled = False
      CmdMod3.Enabled = False
      Command1.Enabled = False
        Command2.Enabled = False
'        Command3.Enabled = False
'        Command4.Enabled = False
'        Command5.Enabled = False
'        Command6.Enabled = False
'        Command7.Enabled = False
'        Command8.Enabled = False
'        Command9.Enabled = False
'        Command10.Enabled = False
'        Command11.Enabled = False
'        Command12.Enabled = False
'        Command13.Enabled = False
'        Command14.Enabled = False
'        Command15.Enabled = False
'        Command16.Enabled = False
'      cmdDel3.Enabled = False
'      cmdGraba3.Enabled = False
'      trv.Enabled = True
   End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
  GlSW = "NADA"
End Sub

Private Sub valida_campos()
  If DtCForm.Text = "" Then
    MsgBox "Debe registrar el Formulario ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If

  If DtcRespId.Text = "" Then
    MsgBox "Debe registrar los datos del Responsable del Llenado del Formulario ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If

  If DtcPromId.Text = "" Then
    MsgBox "Debe registrar los datos del Promotor del proyecto ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If

  If DtcEmpId.Text = "" Then
    MsgBox "Debe registrar los datos de la Unidad Productiva (Institución o Empresa) ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If

'  If Dtc_Comu_Infec.Text = "" Then
'    MsgBox "Debe registrar el Lugar Probable de la Infección (Comunidad o Localidad)...", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'
'  If Dtc_Establ.Text = "" Then
'    MsgBox "Debe registrar el Lugar Probable de la Infección (Establecimiento de Salud) ...", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'
' Dim var_edad_pac As Integer
' If DTP_fecha_muestra >= Dtc_fecha_nac Then
''        var_edad_pac = DTP_fecha_muestra.Value - Dtc_fecha_nac.Text
''        Select Case var_edad_pac
''            Case Is < 7
''                var = "bebe"
''            Case 2
''                var
''            Case Else
''                var = "**no identificado**"
''        End Select
''    MsgBox "Debe registrar una fecha correcta de la Toma de la Muestra ...", vbCritical + vbExclamation, "Validación de datos"
''    Exit Sub
'
' End If
''    rs_M1!fni_edad = var_edad_pac
'
'  If Dtc_doc_id_lab.Text = "" Then
'    MsgBox "Debe registrar los datos del Responsable de la Prueba de Laboratorio ...", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'
'  If Dtc_paracitoCod_lab.Text = "" Then
'    MsgBox "Debe registrar el Resultado (Paracito Encontrado) de la Prueba de Laboratorio ...", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If

End Sub


Private Sub CmdImprimir_Click()
  Dim IResult As Integer
'  var_param = rs_M1!param_codigo
'  var_correl = rs_M1!fni_correlativo
  With Cry_M1
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .WindowShowPrintSetupBtn = True
    .WindowShowGroupTree = True
    .WindowShowExportBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowShowSearchBtn = True
'    .StoredProcParam(0) = var_param
'    .StoredProcParam(1) = var_correl
'    .StoredProcParam(2) = "%"
'    .StoredProcParam(3) = "%"
    .ReportFileName = App.Path & "\Reportes\Ficha Ambiental\FichaAmbiental.rpt"
    IResult = .PrintReport
    If IResult <> 0 Then
        MsgBox .LastErrorNumber & " : " & .LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
  End With
End Sub

Private Sub CmdMod_Click()
  On Error GoTo EditErr
  
  If ADO_M1.Recordset!estado_registro <> "A" And ADO_M1.Recordset!estado_registro <> "S" Then
     mvBookMark = rs_M1.Bookmark
     lblStatus.Caption = "Modificando Registro ..."
     FraCabecera.Enabled = True
     FraEmpresa.Enabled = True
     FraProyecto.Enabled = True
'     FraSitio.Enabled = True
'     FraDescripProy.Enabled = True
'      FraAlterTec.Enabled = True
'      FraInversion.Enabled = True
'      FraActividades.Enabled = True
'      FraRRHH.Enabled = True
'      FraRRNN.Enabled = True
'      FraMatPrima.Enabled = True
'      FraDesecho.Enabled = True
'      FraRuido.Enabled = True
'      FraAlmInsumo.Enabled = True
'      FraTranspInsumo.Enabled = True
'      FraContingencia.Enabled = True
'      FraImpactos.Enabled = True
      FraDeclarJur.Enabled = True
     fraOpciones.Visible = False
     FraGrabarCancelar.Visible = True
     If ADO_M1.Recordset!estado_registro = "N" Or ADO_M1.Recordset!estado_registro = "" Then
        TxtObs.Visible = False
        LblObs.Visible = False
     Else
        TxtObs.Visible = True
        LblObs.Visible = True
     End If
     DtG_M1.Enabled = False
     GlSW = "MOD"
'    cmdAdd3.Enabled = True
      CmdMod3.Enabled = True
      Command1.Enabled = True
        Command2.Enabled = True
'        Command3.Enabled = True
'        Command4.Enabled = True
'        Command5.Enabled = True
'        Command6.Enabled = True
'        Command7.Enabled = True
'        Command8.Enabled = True
'        Command9.Enabled = True
'        Command10.Enabled = True
'        Command11.Enabled = True
'        Command12.Enabled = True
'        Command13.Enabled = True
'        Command14.Enabled = True
'        Command15.Enabled = True
'        Command16.Enabled = True
'      cmdDel3.Enabled = False
'      cmdGraba3.Enabled = True
'    trv.Enabled = False
    'CTRL ESTADO GENERO
'    var_DocId = rs_M1!codigo_responsable
'    Ado_consultor.Recordset.Find "codigo_responsable = '" & Trim(var_DocId) & "' ", , adSearchForward
'    If Not Ado_consultor.Recordset.EOF Then
''        If rs_M1!fni_edad > 10 And Ado_consultor.Recordset("genero_codigo") = "F" Then
''            txttimeEmb.Enabled = True
''        Else
''            txttimeEmb.Enabled = False
''        'txttimeEmb.Enabled = IIf(Ado_consultor.Recordset("genero_codigo") = "F", True, False)
''        End If
'    End If
  Else
    MsgBox "No se puede EDITAR un registro Aprobado o Anulado ...", vbExclamation, "Validación de Registro"
  End If

  Exit Sub

EditErr:
  MsgBox Err.Description

End Sub

Private Sub CmdMod3_Click()
    marca1 = ADO_M1.Recordset.Bookmark
    Frmmo_proyUbic.Lblformulario = ADO_M1.Recordset!tipo_formulario     '"FA"
    Frmmo_proyUbic.lblges_gestion = ADO_M1.Recordset!gestion
    Frmmo_proyUbic.LblFA = ADO_M1.Recordset!Numero_FA
    Frmmo_proyUbic.lblcodigo_solicitud = "A"
    Frmmo_proyUbic.Caption = " Ubicación Física del Proyecto"
    'Frmmo_proyUbic.lbltipo_beneficiario = "-"       'ADO_M1.Recordset!tipo_beneficiario
    Frmmo_proyUbic.Show vbModal
    'Call OptFilGral1_Click

''MODIFICA UBICACION FISICA DEL PROYECTO
' sino = MsgBox("Desea Modificar el Registro Elegido ?", vbQuestion + vbYesNo, "Confirmando...")
'  If sino = vbYes Then
'   If ADO_M1.Recordset.RecordCount > 0 Then
'      If ADO_M1.Recordset!estado_registro = "N" Then
'        swgraba3 = 1
'        DtgProyUbica.AllowAddNew = False
'        DtgProyUbica.AllowDelete = False
'        DtgProyUbica.AllowUpdate = True
'
'        cmdAdd3.Enabled = False
'        cmdMod3.Enabled = False
'        cmdDel3.Enabled = False
'        cmdGraba3.Enabled = True
'      Else
'         MsgBox "No se puede editar un registro APROBADO o ANULADO ", vbInformation, "Formulario "
'      End If
'   Else
'          MsgBox "No Existen Registros habilitados ", vbInformation, "Formulario "
'   End If
'  End If
''-----------
End Sub

Private Sub CmdObs_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de RECHAZAR el Registro? ", vbYesNo + vbQuestion, "Atención")
   If ADO_M1.Recordset!estado_registro = "N" Or ADO_M1.Recordset!estado_registro = "" Then
      If sino = vbYes Then
        ADO_M1.Recordset!estado_registro = "O"
        ADO_M1.Recordset!fecha_modifica = Date
        ADO_M1.Recordset!usr_codigo_mod = GlUsuario
        ADO_M1.Recordset.Update    'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede RECHAZAR un registro APROBADO ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub CmdSal_Click()
    FraCabecera.Enabled = False
    FraEmpresa.Enabled = False
      FraProyecto.Enabled = False
'      FraSitio.Enabled = False
'      FraDescripProy.Enabled = False
'      FraAlterTec.Enabled = False
'      FraInversion.Enabled = False
'      FraActividades.Enabled = False
'      FraRRHH.Enabled = False
'      FraRRNN.Enabled = False
'      FraMatPrima.Enabled = False
'      FraDesecho.Enabled = False
'      FraRuido.Enabled = False
'      FraAlmInsumo.Enabled = False
'      FraTranspInsumo.Enabled = False
'      FraContingencia.Enabled = False
'      FraImpactos.Enabled = False
      FraDeclarJur.Enabled = False
    Unload Me
End Sub

Private Sub CmdVerDisco_Click()
  On Error GoTo Error_Sub
    marca1 = ADO_M1.Recordset.Bookmark
    FrmExplora.lblges_gestion = ADO_M1.Recordset!gestion
    FrmExplora.LblFA = ADO_M1.Recordset!Numero_FA
    FrmExplora.LblForm = ADO_M1.Recordset!tipo_formulario
    'Dir1.Path = App.Path & "\FA-" & lblges_gestion & "-00" & LblFA
    'FrmExplora.Label1 = App.Path & "\FA-" & FrmExplora.lblges_gestion & "-00" & FrmExplora.LblFA
    'Frmexporta.Show vbModal
    If ADO_M1.Recordset!Numero_FA < 10 Then
       NombreCarpeta = App.Path & "\FA\FA-" & ADO_M1.Recordset!gestion & "-00" & ADO_M1.Recordset!Numero_FA
       'e = "\\DMA-196\Sistema\FA\FA-" & ADO_M1.Recordset!gestion & "-00" & ADO_M1.Recordset!Numero_FA
       'e = "\\SERVIDOR\users\public\documents\FA\FA-" & ADO_M1.Recordset!gestion & "-00" & ADO_M1.Recordset!Numero_FA
       e = "\\SERVIDOR\Sistema\FA\FA-" & ADO_M1.Recordset!gestion & "-00" & ADO_M1.Recordset!Numero_FA
    End If
    If ADO_M1.Recordset!Numero_FA > 9 And ADO_M1.Recordset!Numero_FA < 100 Then
       NombreCarpeta = App.Path & "\FA\FA-" & ADO_M1.Recordset!gestion & "-0" & ADO_M1.Recordset!Numero_FA
       'e = "\\DMA-196\Sistema\FA\FA-" & ADO_M1.Recordset!gestion & "-0" & ADO_M1.Recordset!Numero_FA
       e = "\\SERVIDOR\Sistema\FA\FA-" & ADO_M1.Recordset!gestion & "-0" & ADO_M1.Recordset!Numero_FA
       'e = "\\SERVIDOR\users\public\documents\FA\FA-" & ADO_M1.Recordset!gestion & "-0" & ADO_M1.Recordset!Numero_FA
    End If
    If ADO_M1.Recordset!Numero_FA > 99 And ADO_M1.Recordset!Numero_FA < 1000 Then
       NombreCarpeta = App.Path & "\FA\FA-" & ADO_M1.Recordset!gestion & "-" & ADO_M1.Recordset!Numero_FA
       'e = "\\DMA-196\Sistema\FA\FA-" & ADO_M1.Recordset!gestion & "-" & ADO_M1.Recordset!Numero_FA
       'e = "\\SERVIDOR\users\public\documents\FA\FA-" & ADO_M1.Recordset!gestion & "-" & ADO_M1.Recordset!Numero_FA
       e = "\\SERVIDOR\Sistema\FA\FA-" & ADO_M1.Recordset!gestion & "-" & ADO_M1.Recordset!Numero_FA
    End If
'    sino = MsgBox("Elija <SI> para ver la Información de su Disco Local. , o del Servidor <NO> ", vbQuestion + vbYesNo, "Confirmando...")
'    If sino = vbYes Then
    If MsgBox("- Elija 'Si' para ver la Información de su Disco Local ..." & vbCrLf & _
             "- Elija 'No' para ver la Información del SERVIDOR ... ", vbQuestion + vbYesNo, "Confirmar") = vbYes Then
        FrmExplora.Dir1.Path = NombreCarpeta
        FrmExplora.Label1 = NombreCarpeta
    Else
        FrmExplora.Dir1.Path = e
        FrmExplora.Label1 = e
    End If
    FrmExplora.Show vdmodal
Exit Sub
Error_Sub:
 MsgBox Err.Description, vbCritical
End Sub

Private Sub Command1_Click()
    marca1 = ADO_M1.Recordset.Bookmark
    Frmmo_proyUbic.Lblformulario = ADO_M1.Recordset!tipo_formulario     '"FA"
    Frmmo_proyUbic.lblges_gestion = ADO_M1.Recordset!gestion
    Frmmo_proyUbic.LblFA = ADO_M1.Recordset!Numero_FA
    Frmmo_proyUbic.lblcodigo_solicitud = "B"
    Frmmo_proyUbic.Caption = " Medio Humano"
    Frmmo_proyUbic.FrmABMDet.BackColor = &HC0FFFF
    Frmmo_proyUbic.frmgrabDet.BackColor = &HC0FFFF
    Frmmo_proyUbic.ado_proy_ubicacion.BackColor = &H80FFFF
    Frmmo_proyUbic.Show vbModal
End Sub


Private Sub Dtc_depto_cod_Click(Area As Integer)
    Dtc_depto.BoundText = Dtc_depto_cod.BoundText
    Call pProvincia(Dtc_depto.BoundText)
'    Call pRedSalud(Dtc_Depto.BoundText)
End Sub

Private Sub Dtc_depto_Click(Area As Integer)
    Dtc_depto_cod.BoundText = Dtc_depto.BoundText
    Call pProvincia(Dtc_depto_cod.BoundText)
    'Call pRedSalud(Dtc_depto_cod.BoundText)
End Sub

Private Sub pProvincia(CodDepto As String)
   Dim strConsultaP As String

   strConsultaP = "select * from gc_Provincia where depto_codigo='" & CodDepto & "'"

   Set Dtc_prov_cod.RowSource = Nothing
   Set Dtc_prov_cod.RowSource = DB.Execute(strConsultaP, , adCmdText)
   Dtc_prov_cod.ReFill
   Dtc_prov_cod.BoundText = Empty

   Set Dtc_prov.RowSource = Nothing
   Set Dtc_prov.RowSource = DB.Execute(strConsultaP, , adCmdText)
   Dtc_prov.ReFill
   Dtc_prov.BoundText = Empty
End Sub

'Private Sub pProvincia2(CodDepto As String)
'   Dim strConsultaP As String
'
'   strConsultaP = "select * from gc_Provincia where depto_codigo='" & CodDepto & "'"
'
'   Set Dtc_prov_cod2.RowSource = Nothing
'   Set Dtc_prov_cod2.RowSource = DB.Execute(strConsultaP, , adCmdText)
'   Dtc_prov_cod2.ReFill
'   Dtc_prov_cod2.BoundText = Empty
'
'   Set Dtc_prov2.RowSource = Nothing
'   Set Dtc_prov2.RowSource = DB.Execute(strConsultaP, , adCmdText)
'   Dtc_prov2.ReFill
'   Dtc_prov2.BoundText = Empty
'End Sub

Private Sub Dtc_depto_cod2_Click(Area As Integer)
    Dtc_depto2.BoundText = Dtc_depto_cod2.BoundText
    Call pProvincia2(Dtc_depto2.BoundText)
End Sub

Private Sub Dtc_depto2_Click(Area As Integer)
    Dtc_depto_cod2.BoundText = Dtc_depto2.BoundText
    Call pProvincia2(Dtc_depto_cod2.BoundText)
End Sub

Private Sub pProvincia2(depto_codigo2 As String)
   Dim strConsultaP2 As String

   strConsultaP2 = "select * from GC_Provincia where depto_codigo='" & depto_codigo2 & "'"

   Set Dtc_prov_cod2.RowSource = Nothing
   Set Dtc_prov_cod2.RowSource = DB.Execute(strConsultaP2, , adCmdText)
   Dtc_prov_cod2.ReFill
   Dtc_prov_cod2.BoundText = Empty

   Set Dtc_prov2.RowSource = Nothing
   Set Dtc_prov2.RowSource = DB.Execute(strConsultaP2, , adCmdText)
   Dtc_prov2.ReFill
   Dtc_prov2.BoundText = Empty
End Sub

Private Sub Dtc_local_cod2_Click(Area As Integer)
    Dtc_local2.BoundText = Dtc_local_cod2.BoundText
End Sub

Private Sub Dtc_local2_Click(Area As Integer)
    Dtc_local_cod2.BoundText = Dtc_local2.BoundText
End Sub

Private Sub Dtc_munic_cod2_Click(Area As Integer)
    Dtc_munic2.BoundText = Dtc_munic_cod2.BoundText
    Call pComunidad2(Dtc_munic_cod2.BoundText)
End Sub

Private Sub Dtc_munic2_Click(Area As Integer)
    Dtc_munic_cod2.BoundText = Dtc_munic2.BoundText
    Call pComunidad2(Dtc_munic_cod2.BoundText)
End Sub

Private Sub pComunidad2(CodMunic2 As String)
   Dim strConsultaC2 As String

   strConsultaC2 = "select * from GC_comunidad where munic_codigo='" & CodMunic2 & "'"

   Set Dtc_local_cod2.RowSource = Nothing
   Set Dtc_local_cod2.RowSource = DB.Execute(strConsultaC2, , adCmdText)
   Dtc_local_cod2.ReFill
   Dtc_local_cod2.BoundText = Empty

   Set Dtc_local2.RowSource = Nothing
   Set Dtc_local2.RowSource = DB.Execute(strConsultaC2, , adCmdText)
   Dtc_local2.ReFill
   Dtc_local2.BoundText = Empty
End Sub

Private Sub Dtc_prov_Click(Area As Integer)
    Dtc_prov_cod.BoundText = Dtc_prov.BoundText
    Call pMunicipio(Dtc_prov_cod.BoundText)
End Sub

Private Sub Dtc_prov_cod_Click(Area As Integer)
    Dtc_prov.BoundText = Dtc_prov_cod.BoundText
    Call pMunicipio(Dtc_prov.BoundText)
End Sub

Private Sub pMunicipio(CodProv As String)
   Dim strConsultaM As String

   strConsultaM = "select * from gc_Municipio where prov_codigo='" & CodProv & "'"

   Set Dtc_munic_cod.RowSource = Nothing
   Set Dtc_munic_cod.RowSource = DB.Execute(strConsultaM, , adCmdText)
   Dtc_munic_cod.ReFill
   Dtc_munic_cod.BoundText = Empty

   Set Dtc_munic.RowSource = Nothing
   Set Dtc_munic.RowSource = DB.Execute(strConsultaM, , adCmdText)
   Dtc_munic.ReFill
   Dtc_munic.BoundText = Empty
End Sub

Private Sub Dtc_munic_Click(Area As Integer)
    Dtc_munic_cod.BoundText = Dtc_munic.BoundText
    'Call pComunidad(Dtc_munic_cod.BoundText)
End Sub

Private Sub Dtc_munic_cod_Click(Area As Integer)
    Dtc_munic.BoundText = Dtc_munic_cod.BoundText
    'Call pComunidad(Dtc_munic.BoundText)
End Sub

Private Sub pComunidad(CodMunic As String)
'   Dim strConsultaC As String
'
'   strConsultaC = "select * from t_comunidad where codmunicip='" & CodMunic & "'"
'
'   Set Dtc_local_cod.RowSource = Nothing
'   Set Dtc_local_cod.RowSource = db.Execute(strConsultaC, , adCmdText)
'   Dtc_local_cod.ReFill
'   Dtc_local_cod.BoundText = Empty
'
'   Set Dtc_local.RowSource = Nothing
'   Set Dtc_local.RowSource = db.Execute(strConsultaC, , adCmdText)
'   Dtc_local.ReFill
'   Dtc_local.BoundText = Empty
End Sub

Private Sub Dtc_local_Click(Area As Integer)
    Dtc_local_cod.BoundText = Dtc_local.BoundText
End Sub

Private Sub Dtc_local_cod_Click(Area As Integer)
    Dtc_local.BoundText = Dtc_local_cod.BoundText
End Sub

'Private Sub Dtc_Pac_Id_LostFocus()
'    If Dtc_Genero = "F" Then
'        'lbl_genero.Caption = "Femenino"
'        txttimeEmb.Enabled = True
'    Else
'        'lbl_genero.Caption = "Masculino"
'        txttimeEmb.Enabled = False
'    End If
'End Sub

'Private Sub Dtc_pac_1apell_LostFocus()
'    If Dtc_Genero = "F" Then
'        'lbl_genero.Caption = "Femenino"
'        txttimeEmb.Enabled = True
'    Else
'        'lbl_genero.Caption = "Masculino"
'        txttimeEmb.Enabled = False
'    End If
'End Sub

Private Sub Dtc_pac_nombre_LostFocus()
    If Dtc_Genero = "F" Then
        'lbl_genero.Caption = "Femenino"
        txttimeEmb.Enabled = True
    Else
        'lbl_genero.Caption = "Masculino"
        txttimeEmb.Enabled = False
    End If
End Sub

Private Sub Dtc_Genero_Click(Area As Integer)
    Dtc_Pac_Id.BoundText = Dtc_Genero.BoundText
    Dtc_pac_1apell.BoundText = Dtc_Genero.BoundText
    Dtc_pac_2apell.BoundText = Dtc_Genero.BoundText
    Dtc_pac_nombre.BoundText = Dtc_Genero.BoundText
    Dtc_fecha_nac.BoundText = Dtc_Genero.BoundText
    'Dtc_ocup.BoundText = Dtc_Genero.BoundText
    Dtc_ocup2.BoundText = Dtc_Genero.BoundText
End Sub

Private Sub Dtc_Act_Click(Area As Integer)
    Dtc_Act_Des.BoundText = Dtc_Act.BoundText
End Sub

Private Sub Dtc_Act_Des_Click(Area As Integer)
    Dtc_Act.BoundText = Dtc_Act_Des.BoundText
End Sub


Private Sub pProvincia_infec(CodDepto As String)
   Dim strConsultaP As String

   strConsultaP = "select * from gc_Provincia where depto_codigo='" & CodDepto & "'"

   Set Dtc_Prov_Infec.RowSource = Nothing
   Set Dtc_Prov_Infec.RowSource = DB.Execute(strConsultaP, , adCmdText)
   Dtc_Prov_Infec.ReFill
   Dtc_Prov_Infec.BoundText = Empty

   Set Dtc_Prov_Infec_Des.RowSource = Nothing
   Set Dtc_Prov_Infec_Des.RowSource = DB.Execute(strConsultaP, , adCmdText)
   Dtc_Prov_Infec_Des.ReFill
   Dtc_Prov_Infec_Des.BoundText = Empty
End Sub


Private Sub pMunicipio_infec(CodProv As String)
   Dim strConsultaM As String

   strConsultaM = "select * from gc_Municipio where prov_codigo='" & CodProv & "'"

   Set Dtc_Munic_Infec.RowSource = Nothing
   Set Dtc_Munic_Infec.RowSource = DB.Execute(strConsultaM, , adCmdText)
   Dtc_Munic_Infec.ReFill
   Dtc_Munic_Infec.BoundText = Empty

   Set Dtc_Munic_Infec_Des.RowSource = Nothing
   Set Dtc_Munic_Infec_Des.RowSource = DB.Execute(strConsultaM, , adCmdText)
   Dtc_Munic_Infec_Des.ReFill
   Dtc_Munic_Infec_Des.BoundText = Empty
End Sub


Private Sub pComunidad_infec(CodMunic As String)
   Dim strConsultaC As String

   strConsultaC = "select * from t_comunidad where codmunicip='" & CodMunic & "'"

   Set Dtc_Comu_Infec.RowSource = Nothing
   Set Dtc_Comu_Infec.RowSource = DB.Execute(strConsultaC, , adCmdText)
   Dtc_Comu_Infec.ReFill
   Dtc_Comu_Infec.BoundText = Empty

   Set Dtc_Comu_Infec_Des.RowSource = Nothing
   Set Dtc_Comu_Infec_Des.RowSource = DB.Execute(strConsultaC, , adCmdText)
   Dtc_Comu_Infec_Des.ReFill
   Dtc_Comu_Infec_Des.BoundText = Empty
End Sub

Private Sub Dtc_Comu_Infec_LostFocus()
    If Dtc_local_cod = Dtc_Comu_Infec Then
        Dtc_tipo_casoCod = 2
        Dtc_tipo_caso = "Autoctono"
    Else
        Dtc_tipo_casoCod = 1
        Dtc_tipo_caso = "Importado"
    End If
End Sub


Private Sub pRedSalud(CodDepto As String)
'   Dim strConsultaP As String
'
'   strConsultaP = "select * from t_area where depto_codigo='" & CodDepto & "'"
'
'   Set Dtc_RedSalud.RowSource = Nothing
'   Set Dtc_RedSalud.RowSource = db.Execute(strConsultaP, , adCmdText)
'   Dtc_RedSalud.ReFill
'   Dtc_RedSalud.BoundText = Empty
'
'   Set Dtc_RedSalud_des.RowSource = Nothing
'   Set Dtc_RedSalud_des.RowSource = db.Execute(strConsultaP, , adCmdText)
'   Dtc_RedSalud_des.ReFill
'   Dtc_RedSalud_des.BoundText = Empty
End Sub

Private Sub Dtc_RedSalud_Click(Area As Integer)
    Dtc_RedSalud_des.BoundText = Dtc_RedSalud.BoundText
    Call pEstabl(Dtc_RedSalud_des.BoundText)
End Sub

Private Sub Dtc_RedSalud_des_Click(Area As Integer)
    Dtc_RedSalud.BoundText = Dtc_RedSalud_des.BoundText
    Call pEstabl(Dtc_RedSalud.BoundText)
End Sub

Private Sub pEstabl(CodRedSal As String)
   Dim strConsultaC As String

   strConsultaC = "select * from t_estabGest where codarea='" & CodRedSal & "'"

   Set Dtc_Establ.RowSource = Nothing
   Set Dtc_Establ.RowSource = DB.Execute(strConsultaC, , adCmdText)
   Dtc_Establ.ReFill
   Dtc_Establ.BoundText = Empty

   Set Dtc_Establ_Des.RowSource = Nothing
   Set Dtc_Establ_Des.RowSource = DB.Execute(strConsultaC, , adCmdText)
   Dtc_Establ_Des.ReFill
   Dtc_Establ_Des.BoundText = Empty
End Sub

'Private Sub Dtc_prov2_Click(Area As Integer)
'    DtcRespNom.BoundText = Dtc_depto2.BoundText
'    DtcRespProf.BoundText = Dtc_prov2.BoundText
'    DtcRespCargo.BoundText = Dtc_prov2.BoundText
'    DtcRespRenca.BoundText = Dtc_prov2.BoundText
'    DtcRespDom.BoundText = Dtc_prov2.BoundText
'    DtcRespTelf.BoundText = Dtc_prov2.BoundText
'    DtcRespEmail.BoundText = Dtc_prov2.BoundText
'    DtcRespId.BoundText = Dtc_prov2.BoundText
'    Dtc_depto2.BoundText = Dtc_prov2.BoundText
'    Dtc_munic2.BoundText = Dtc_prov2.BoundText
'    Dtc_Comu2.BoundText = Dtc_prov2.BoundText
'End Sub

'Private Sub pMunicipio2(CodProv As String)
'   Dim strConsultaM As String
'
'   strConsultaM = "select * from gc_Municipio where prov_codigo='" & CodProv & "'"
'
'   Set Dtc_munic_cod2.RowSource = Nothing
'   Set Dtc_munic_cod2.RowSource = DB.Execute(strConsultaM, , adCmdText)
'   Dtc_munic_cod2.ReFill
'   Dtc_munic_cod2.BoundText = Empty
'
'   Set Dtc_munic2.RowSource = Nothing
'   Set Dtc_munic2.RowSource = DB.Execute(strConsultaM, , adCmdText)
'   Dtc_munic2.ReFill
'   Dtc_munic2.BoundText = Empty
'End Sub

Private Sub Dtc_prov_cod2_Click(Area As Integer)
    Dtc_prov2.BoundText = Dtc_prov_cod2.BoundText
    Call pMunicipio2(Dtc_prov2.BoundText)
End Sub

Private Sub Dtc_prov2_Click(Area As Integer)
    Dtc_prov_cod2.BoundText = Dtc_prov2.BoundText
    Call pMunicipio2(Dtc_prov_cod2.BoundText)
End Sub

Private Sub pMunicipio2(CodProv2 As String)
   Dim strConsultaM2 As String

   strConsultaM2 = "select * from gc_Municipio where prov_codigo='" & CodProv2 & "'"

   Set Dtc_munic_cod2.RowSource = Nothing
   Set Dtc_munic_cod2.RowSource = DB.Execute(strConsultaM2, , adCmdText)
   Dtc_munic_cod2.ReFill
   Dtc_munic_cod2.BoundText = Empty

   Set Dtc_munic2.RowSource = Nothing
   Set Dtc_munic2.RowSource = DB.Execute(strConsultaM2, , adCmdText)
   Dtc_munic2.ReFill
   Dtc_munic2.BoundText = Empty
End Sub

Private Sub Dtc_SitioTopo_Click(Area As Integer)
    Dtc_SitioTopoDes.BoundText = Dtc_SitioTopo.BoundText
End Sub

Private Sub Dtc_SitioTopoDes_Click(Area As Integer)
    Dtc_SitioTopo.BoundText = Dtc_SitioTopoDes.BoundText
End Sub

Private Sub DtCForm_LostFocus()
    var_form = IIf(IsNull(DtCForm.Text), "FA", DtCForm.Text)
    DtCFormDes.BoundText = DtCForm.BoundText
End Sub

Private Sub DtCFormDes_Click(Area As Integer)
    DtCForm.BoundText = DtCFormDes.BoundText
End Sub

Private Sub DtcRespCargo_Click(Area As Integer)
    DtcRespId.BoundText = DtcRespCargo.BoundText
    DtcRespNom.BoundText = DtcRespCargo.BoundText
'    DtcRespProf.BoundText = DtcRespCargo.BoundText
    DtcRespRenca.BoundText = DtcRespCargo.BoundText
'    DtcRespDom.BoundText = DtcRespCargo.BoundText
'    DtcRespTelf.BoundText = DtcRespCargo.BoundText
'    DtcRespEmail.BoundText = DtcRespCargo.BoundText
'    Dtc_depto2.BoundText = DtcRespCargo.BoundText
'    Dtc_prov2.BoundText = DtcRespCargo.BoundText
'    Dtc_munic2.BoundText = DtcRespCargo.BoundText
'    Dtc_Comu2.BoundText = DtcRespCargo.BoundText
End Sub

Private Sub DtcRespId_Click(Area As Integer)
    DtcRespNom.BoundText = DtcRespId.BoundText
'    DtcRespProf.BoundText = DtcRespId.BoundText
    DtcRespCargo.BoundText = DtcRespId.BoundText
    DtcRespRenca.BoundText = DtcRespId.BoundText
'    DtcRespDom.BoundText = DtcRespId.BoundText
'    DtcRespTelf.BoundText = DtcRespId.BoundText
'    DtcRespEmail.BoundText = DtcRespId.BoundText
'    Dtc_depto2.BoundText = DtcRespId.BoundText
'    Dtc_prov2.BoundText = DtcRespId.BoundText
'    Dtc_munic2.BoundText = DtcRespId.BoundText
'    Dtc_Comu2.BoundText = DtcRespId.BoundText
End Sub

Private Sub DtcRespNom_Click(Area As Integer)
    DtcRespId.BoundText = DtcRespNom.BoundText
'    DtcRespProf.BoundText = DtcRespNom.BoundText
    DtcRespCargo.BoundText = DtcRespNom.BoundText
    DtcRespRenca.BoundText = DtcRespNom.BoundText
'    DtcRespDom.BoundText = DtcRespNom.BoundText
'    DtcRespTelf.BoundText = DtcRespNom.BoundText
'    DtcRespEmail.BoundText = DtcRespNom.BoundText
'    Dtc_depto2.BoundText = DtcRespNom.BoundText
'    Dtc_prov2.BoundText = DtcRespNom.BoundText
'    Dtc_munic2.BoundText = DtcRespNom.BoundText
'    Dtc_Comu2.BoundText = DtcRespNom.BoundText
End Sub

Private Sub DtcRespProf_Click(Area As Integer)
'    DtcRespId.BoundText = DtcRespProf.BoundText
'    DtcRespNom.BoundText = DtcRespProf.BoundText
'    DtcRespCargo.BoundText = DtcRespProf.BoundText
'    DtcRespRenca.BoundText = DtcRespProf.BoundText
''    DtcRespDom.BoundText = DtcRespProf.BoundText
''    DtcRespTelf.BoundText = DtcRespProf.BoundText
''    DtcRespEmail.BoundText = DtcRespProf.BoundText
''    Dtc_depto2.BoundText = DtcRespProf.BoundText
''    Dtc_prov2.BoundText = DtcRespProf.BoundText
''    Dtc_munic2.BoundText = DtcRespProf.BoundText
''    Dtc_Comu2.BoundText = DtcRespProf.BoundText
End Sub

Private Sub DtcRespRenca_Click(Area As Integer)
    DtcRespId.BoundText = DtcRespRenca.BoundText
    DtcRespNom.BoundText = DtcRespRenca.BoundText
'    DtcRespProf.BoundText = DtcRespRenca.BoundText
    DtcRespCargo.BoundText = DtcRespRenca.BoundText
'    DtcRespDom.BoundText = DtcRespRenca.BoundText
'    DtcRespTelf.BoundText = DtcRespRenca.BoundText
'    DtcRespEmail.BoundText = DtcRespRenca.BoundText
'    Dtc_depto2.BoundText = DtcRespRenca.BoundText
'    Dtc_prov2.BoundText = DtcRespRenca.BoundText
'    Dtc_munic2.BoundText = DtcRespRenca.BoundText
'    Dtc_Comu2.BoundText = DtcRespRenca.BoundText
End Sub

Private Sub DtcRespTelf_Click(Area As Integer)
'    DtcRespId.BoundText = DtcRespTelf.BoundText
'    DtcRespNom.BoundText = DtcRespTelf.BoundText
'    DtcRespProf.BoundText = DtcRespTelf.BoundText
'    DtcRespCargo.BoundText = DtcRespTelf.BoundText
'    DtcRespRenca.BoundText = DtcRespTelf.BoundText
''    DtcRespDom.BoundText = DtcRespTelf.BoundText
''    DtcRespEmail.BoundText = DtcRespTelf.BoundText
''    Dtc_depto2.BoundText = DtcRespTelf.BoundText
''    Dtc_prov2.BoundText = DtcRespTelf.BoundText
''    Dtc_munic2.BoundText = DtcRespTelf.BoundText
''    Dtc_Comu2.BoundText = DtcRespTelf.BoundText
End Sub




Private Sub DtgProyUbica_LostFocus()
   If ADO_M1.Recordset.RecordCount > 0 And Not IsNull(DtgProyUbica.Columns("comun").Value) And (DtgProyUbica.Columns("comun").Value) <> "" Then
      If ADO_M1.Recordset!estado_registro = "N" Then
        marca1 = ADO_M1.Recordset.Bookmark
        VARB = DtgProyUbica.Columns("comun_codigo").Value
        VARBD = DtgProyUbica.Columns("depto_codigo").Value
        VARG = DtgProyUbica.Columns("prov_codigo").Value
        VARS = DtgProyUbica.Columns("munic_codigo").Value
        VARU = DtgProyUbica.Columns("depto").Value
        VARPU = DtgProyUbica.Columns("altitud_snm").Value
        VAR10 = DtgProyUbica.Columns("prov").Value
        VAR11 = DtgProyUbica.Columns("munic").Value
        VAR12 = DtgProyUbica.Columns("comun").Value
        VAR13 = DtgProyUbica.Columns("latitud").Value
        VAR14 = DtgProyUbica.Columns("longitud").Value
'        MarcaB = adoao_solicitud_bien.Recordset.Bookmark
'        Call Abre_Sol_Bien
'        'MarcaB = rs_ao_solicitud_bien.Bookmark
'        adoao_solicitud_bien.Recordset.Bookmark = MarcaB
        rs_ProyUbic!comun_codigo = VARB
        rs_ProyUbic!depto_codigo = VARBD
        rs_ProyUbic!prov_codigo = VARG
        rs_ProyUbic!munic_codigo = VARS
        rs_ProyUbic!depto = VARU
        rs_ProyUbic!altitud_snm = VARPU
        rs_ProyUbic!prov = VAR10
        rs_ProyUbic!munic = VAR11
        rs_ProyUbic!comun = VAR12
        rs_ProyUbic!latitud = VAR13
        rs_ProyUbic!longitud = VAR14
        rs_ProyUbic.Update
        'Call Abre_Sol_Bien
        rs_ProyUbic.MoveLast
'        Call OptFilGral1_Click
        'adosolicitud.Recordset.BookMark = marca1
        'adosolicitud.Refresh
        'swgrabar = 2
      Else
         MsgBox "No se puede modificar un registro APROBADO ", vbInformation, "Formulario 1"
      End If
   Else
         MsgBox "Verifique los datos para continuar ... ", vbInformation, "Formulario 1"
   End If
End Sub


Private Sub Form_Load()
'     Command1.Caption = " Eliminar "
'	Call SeguridadSet(Me)
End Sub

Private Sub abrirtabla_maestra()
  Set rs_M1 = New Recordset
  If rs_M1.State = 1 Then rs_M1.Close
  queryinicial = "select * from gc_beneficiario "
  rs_M1.Open queryinicial, DB, adOpenKeyset, adLockOptimistic
  rs_M1.Sort = "denominacion_beneficiario"
  Set ADO_M1.Recordset = rs_M1
  Set DtG_M1.DataSource = ADO_M1.Recordset
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
    With rs_M1
      If Not (.BOF And .EOF) Then
        mvBookMark = .Bookmark
      End If
      GlSW = "ADD"
      .AddNew
      lblStatus.Caption = "Adicionando Nuevo Registro ..."
      FraCabecera.Enabled = True
      FraEmpresa.Enabled = True
      FraProyecto.Enabled = True
'      FraSitio.Enabled = True
'      FraDescripProy.Enabled = True
'      FraAlterTec.Enabled = True
'      FraInversion.Enabled = True
'      FraActividades.Enabled = True
'      FraRRHH.Enabled = True
'      FraRRNN.Enabled = True
'      FraMatPrima.Enabled = True
'      FraDesecho.Enabled = True
'      FraRuido.Enabled = True
'      FraAlmInsumo.Enabled = True
'      FraTranspInsumo.Enabled = True
'      FraContingencia.Enabled = True
'      FraImpactos.Enabled = True
      FraDeclarJur.Enabled = True
      fraOpciones.Visible = False
      FraGrabarCancelar.Visible = True
      DtCForm.Text = "FA"
      txtParam.Text = glGestion
      Dtc_depto_cod.Text = glDepto
      Dtc_prov_cod.Text = glProvi
      Dtc_munic_cod.Text = glMunic
      DTP_FechaLectura.Value = Date
      DtG_M1.Enabled = False
      TxtObs.Visible = False
      LblObs.Visible = False
'      cmdAdd3.Enabled = True
      CmdMod3.Enabled = False
      Command1.Enabled = False
    Command2.Enabled = False
'    Command3.Enabled = False
'    Command4.Enabled = False
'    Command5.Enabled = False
'    Command6.Enabled = False
'    Command7.Enabled = False
'    Command8.Enabled = False
'    Command9.Enabled = False
'    Command10.Enabled = False
'    Command11.Enabled = False
'    Command12.Enabled = False
'    Command13.Enabled = False
'    Command14.Enabled = False
'    Command15.Enabled = False
'    Command16.Enabled = False
'      cmdDel3.Enabled = False
'      cmdGraba3.Enabled = True
'      trv.Enabled = False
    End With
  
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  rs_M1.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub



'Private Sub SetButtons(bVal As Boolean)
'  CmdAdd.Visible = bVal
'  CmdMod.Visible = bVal
'  CmdGrabar.Visible = Not bVal
'  CmdCancelar.Visible = Not bVal
'  CmdDel.Visible = bVal
'  CmdSal.Visible = bVal
'  cmdRefresh.Visible = bVal
''  cmdNext.Enabled = bVal
''  cmdFirst.Enabled = bVal
''  cmdLast.Enabled = bVal
''  cmdPrevious.Enabled = bVal
'End Sub

Private Sub TDBComun_DropDownClose()
    DtgProyUbica.Columns("Gestion").Value = TDBComun.Columns("Gestion").Value
    DtgProyUbica.Columns("Numero_FA").Value = TDBComun.Columns("Numero_FA").Value
    DtgProyUbica.Columns("comun_codigo").Value = TDBComun.Columns("comun_codigo").Value
    DtgProyUbica.Columns("depto_codigo").Value = TDBComun.Columns("depto_codigo").Value
    DtgProyUbica.Columns("prov_codigo").Value = TDBComun.Columns("prov_codigo").Value
    DtgProyUbica.Columns("munic_codigo").Value = TDBComun.Columns("munic_codigo").Value
    DtgProyUbica.Columns("depto").Value = TDBComun.Columns("depto_descripcion").Value
    DtgProyUbica.Columns("prov").Value = TDBComun.Columns("prov_descripcion").Value
    DtgProyUbica.Columns("munic").Value = TDBComun.Columns("munic_descripcion").Value
    DtgProyUbica.Columns("comun").Value = TDBComun.Columns("comun_descripcion").Value
    DtgProyUbica.Columns("latitud").Value = TDBComun.Columns("latitud").Value
    DtgProyUbica.Columns("longitud").Value = TDBComun.Columns("longitud").Value
    DtgProyUbica.Columns("altitud_snm").Value = TDBComun.Columns("altitud_snm").Value
End Sub

Private Sub LlenaArbol()
  Dim Nodo As Node
    Set rs_TramiteC = New ADODB.Recordset
    'Set RsDet = New ADODB.Recordset
    'GlSqlAux = "SELECT CodGrupo, DescGrupo " & _
    '           "FROM ALCLGrupo " & _
    '           "WHERE   cast(codgrupo as int) >100 OR cast(codgrupo as int) <5" & _
    '           "ORDER BY cast(CodGrupo as int)"
               
     GlSqlAux = "SELECT contenido_cod, contenido_descripcion, FORM " & _
               "FROM gc_tramite_contenido WHERE tipo_formulario='F05'  " & _
               "ORDER BY cast(contenido_cod as int)"
    rs_TramiteC.Open GlSqlAux, DB, adOpenStatic
    If rs_TramiteC.RecordCount > 0 Then
        Set Nodo = trv.Nodes.Add(, , "M", "Ficha Ambiental", "Raiz")
        Nodo.Bold = True
        Nodo.Expanded = True
        While Not rs_TramiteC.EOF
            'Set Nodo = trv.Nodes.Add("M", tvwChild, "G" & rs_TramiteC!contenido_cod, rs_TramiteC!contenido_cod & " - " & rs_TramiteC!contenido_descripcion, rs_TramiteC!Form, "Cerrado", "Abierto")
            Set Nodo = trv.Nodes.Add("M", tvwChild, "G" & rs_TramiteC!contenido_cod, rs_TramiteC!contenido_cod & " - " & rs_TramiteC!contenido_descripcion, "Cerrado", "Abierto")
            rs_TramiteC.MoveNext
        Wend
        
'        GlSqlAux = "SELECT CodGrupo, CodDetalle, DescDetalle " & _
'                   "FROM ALCLDetalle " & _
'                   "WHERE Estado = 1 and cast(codgrupo as int) >100 OR cast(codgrupo as int) <5" & _
'                   "ORDER BY CodGrupo, CodDetalle"
'                   '"WHERE Estado = 1 AND CodGrupo IN ()" & _'
'        RsDet.Open GlSqlAux, db, adOpenStatic
'        While Not RsDet.EOF
'            Set Nodo = trv.Nodes.Add("G" & RsDet!CodGrupo, tvwChild, "D" & RsDet!CodGrupo & "-" & RsDet!codDetalle, RsDet!CodGrupo & "-" & RsDet!codDetalle & " : " & RsDet!descdetalle, "Detalle")
'            RsDet.MoveNext
'        Wend
        'Cmdaceptar.Enabled = True
    Else
        Set Nodo = trv.Nodes.Add(, , "M", "No Existe Contenido")
        Nodo.Bold = True
        Nodo.Expanded = True
        'Cmdaceptar.Enabled = False
    End If
End Sub

Private Sub trv_NodeClick(ByVal Node As MSComctlLib.Node)
    'If InStr(Node.Key, "D") > 0 Then
    If Node.Key <> "" Then
       QCodigo = Mid(Node.Key, 2)
       QItem = Mid(Node.Text, InStr(Node.Text, ":") + 1)
        'QItem2 = Mid(Node.Text, InStr(Node.Text, ":") + 1)
        
'        If ADO_M1.Recordset.RecordCount > 0 Then
'          marca1 = ADO_M1.Recordset.Bookmark
'          var_gest = ADO_M1.Recordset!gestion
'          var_FA = ADO_M1.Recordset!Numero_FA
'          ADO_M1.Recordset.Move marca1 - 1
'        Else
'          MsgBox "No Existen Registros ", vbInformation, "Formulario FA"
'        End If
        Select Case QCodigo
            Case "1"
                If ADO_M1.Recordset.RecordCount > 0 Then
                  'marca1 = ADO_M1.Recordset.Bookmark
                  SSTab1.Tab = 0
                  SSTab1.TabEnabled(0) = True
                  SSTab1.TabEnabled(1) = False
                  SSTab1.TabEnabled(2) = False
                  SSTab1.TabEnabled(3) = False
                  SSTab1.TabEnabled(4) = False
                  SSTab1.TabEnabled(5) = False
'                  SSTab1.TabEnabled(6) = False
'                  SSTab1.TabEnabled(7) = False
'                  SSTab1.TabEnabled(8) = False
'                  SSTab1.TabEnabled(9) = False
'                  SSTab1.TabEnabled(10) = False
'                  SSTab1.TabEnabled(11) = False
'                  SSTab1.TabEnabled(12) = False
'                  SSTab1.TabEnabled(13) = False
'                  SSTab1.TabEnabled(14) = False
'                  SSTab1.TabEnabled(15) = False
'                  SSTab1.TabEnabled(16) = False
'                  SSTab1.TabEnabled(17) = False
                    
                  'frmfo_FA01.lbl_Gestion.Caption = ADO_M1.Recordset!gestion
                  'frmfo_FA01.lbl_FA = ADO_M1.Recordset!Numero_FA
                  ''SSTab1.Tab = 0
                  'frmfo_FA01.Show 'vbModal
                  'ADO_M1.Refresh
                  'ADO_M1.Recordset.Move marca1 - 1
                Else
                  MsgBox "No Existen Registros ", vbInformation, "-FA-"
                End If
            Case "2"
                'Frm_FA02.Show
                  SSTab1.Tab = 1
                  SSTab1.TabEnabled(0) = False
                  SSTab1.TabEnabled(1) = True
                  SSTab1.TabEnabled(2) = False
                  SSTab1.TabEnabled(3) = False
                  SSTab1.TabEnabled(4) = False
                  SSTab1.TabEnabled(5) = False
'                  SSTab1.TabEnabled(6) = False
'                  SSTab1.TabEnabled(7) = False
'                  SSTab1.TabEnabled(8) = False
'                  SSTab1.TabEnabled(9) = False
'                  SSTab1.TabEnabled(10) = False
'                  SSTab1.TabEnabled(11) = False
'                  SSTab1.TabEnabled(12) = False
'                  SSTab1.TabEnabled(13) = False
'                  SSTab1.TabEnabled(14) = False
'                  SSTab1.TabEnabled(15) = False
'                  SSTab1.TabEnabled(16) = False
'                  SSTab1.TabEnabled(17) = False
            Case "3"
                'Frm_FA03.Show
                  SSTab1.Tab = 2
                  SSTab1.TabEnabled(0) = False
                  SSTab1.TabEnabled(1) = False
                  SSTab1.TabEnabled(2) = True
                  SSTab1.TabEnabled(3) = False
                  SSTab1.TabEnabled(4) = False
                  SSTab1.TabEnabled(5) = False
'                  SSTab1.TabEnabled(6) = False
'                  SSTab1.TabEnabled(7) = False
'                  SSTab1.TabEnabled(8) = False
'                  SSTab1.TabEnabled(9) = False
'                  SSTab1.TabEnabled(10) = False
'                  SSTab1.TabEnabled(11) = False
'                  SSTab1.TabEnabled(12) = False
'                  SSTab1.TabEnabled(13) = False
'                  SSTab1.TabEnabled(14) = False
'                  SSTab1.TabEnabled(15) = False
'                  SSTab1.TabEnabled(16) = False
'                  SSTab1.TabEnabled(17) = False
            Case "4"
                'Frm_FA04.Show
                  SSTab1.Tab = 3
                  SSTab1.TabEnabled(0) = False
                  SSTab1.TabEnabled(1) = False
                  SSTab1.TabEnabled(2) = False
                  SSTab1.TabEnabled(3) = True
                  SSTab1.TabEnabled(4) = False
                  SSTab1.TabEnabled(5) = False
'                  SSTab1.TabEnabled(6) = False
'                  SSTab1.TabEnabled(7) = False
'                  SSTab1.TabEnabled(8) = False
'                  SSTab1.TabEnabled(9) = False
'                  SSTab1.TabEnabled(10) = False
'                  SSTab1.TabEnabled(11) = False
'                  SSTab1.TabEnabled(12) = False
'                  SSTab1.TabEnabled(13) = False
'                  SSTab1.TabEnabled(14) = False
'                  SSTab1.TabEnabled(15) = False
'                  SSTab1.TabEnabled(16) = False
'                  SSTab1.TabEnabled(17) = False
            Case "5"
                'Frm_FA05.Show
                  SSTab1.Tab = 4
                  SSTab1.TabEnabled(0) = False
                  SSTab1.TabEnabled(1) = False
                  SSTab1.TabEnabled(2) = False
                  SSTab1.TabEnabled(3) = False
                  SSTab1.TabEnabled(4) = True
                  SSTab1.TabEnabled(5) = False
'                  SSTab1.TabEnabled(6) = False
'                  SSTab1.TabEnabled(7) = False
'                  SSTab1.TabEnabled(8) = False
'                  SSTab1.TabEnabled(9) = False
'                  SSTab1.TabEnabled(10) = False
'                  SSTab1.TabEnabled(11) = False
'                  SSTab1.TabEnabled(12) = False
'                  SSTab1.TabEnabled(13) = False
'                  SSTab1.TabEnabled(14) = False
'                  SSTab1.TabEnabled(15) = False
'                  SSTab1.TabEnabled(16) = False
'                  SSTab1.TabEnabled(17) = False
            Case "6"
                'Frm_FA06.Show
                  SSTab1.Tab = 5
                  SSTab1.TabEnabled(0) = False
                  SSTab1.TabEnabled(1) = False
                  SSTab1.TabEnabled(2) = False
                  SSTab1.TabEnabled(3) = False
                  SSTab1.TabEnabled(4) = False
                  SSTab1.TabEnabled(5) = True
'                  SSTab1.TabEnabled(6) = False
'                  SSTab1.TabEnabled(7) = False
'                  SSTab1.TabEnabled(8) = False
'                  SSTab1.TabEnabled(9) = False
'                  SSTab1.TabEnabled(10) = False
'                  SSTab1.TabEnabled(11) = False
'                  SSTab1.TabEnabled(12) = False
'                  SSTab1.TabEnabled(13) = False
'                  SSTab1.TabEnabled(14) = False
'                  SSTab1.TabEnabled(15) = False
'                  SSTab1.TabEnabled(16) = False
'                  SSTab1.TabEnabled(17) = False
            Case "7"
                'Frm_FA07.Show
'                  SSTab1.Tab = 6
                  SSTab1.TabEnabled(0) = False
                  SSTab1.TabEnabled(1) = False
                  SSTab1.TabEnabled(2) = False
                  SSTab1.TabEnabled(3) = False
                  SSTab1.TabEnabled(4) = False
                  SSTab1.TabEnabled(5) = False
'                  SSTab1.TabEnabled(6) = True
'                  SSTab1.TabEnabled(7) = False
'                  SSTab1.TabEnabled(8) = False
'                  SSTab1.TabEnabled(9) = False
'                  SSTab1.TabEnabled(10) = False
'                  SSTab1.TabEnabled(11) = False
'                  SSTab1.TabEnabled(12) = False
'                  SSTab1.TabEnabled(13) = False
'                  SSTab1.TabEnabled(14) = False
'                  SSTab1.TabEnabled(15) = False
'                  SSTab1.TabEnabled(16) = False
'                  SSTab1.TabEnabled(17) = False
            Case "8"
                'Frm_FA08.Show
                  SSTab1.Tab = 7
                  SSTab1.TabEnabled(0) = False
                  SSTab1.TabEnabled(1) = False
                  SSTab1.TabEnabled(2) = False
                  SSTab1.TabEnabled(3) = False
                  SSTab1.TabEnabled(4) = False
                  SSTab1.TabEnabled(5) = False
'                  SSTab1.TabEnabled(6) = False
'                  SSTab1.TabEnabled(7) = True
'                  SSTab1.TabEnabled(8) = False
'                  SSTab1.TabEnabled(9) = False
'                  SSTab1.TabEnabled(10) = False
'                  SSTab1.TabEnabled(11) = False
'                  SSTab1.TabEnabled(12) = False
'                  SSTab1.TabEnabled(13) = False
'                  SSTab1.TabEnabled(14) = False
'                  SSTab1.TabEnabled(15) = False
'                  SSTab1.TabEnabled(16) = False
'                  SSTab1.TabEnabled(17) = False
            
            Case Else
                MsgBox "**no identificado**"
        End Select
    Else
        QCodigo = ""
        QItem = ""
    End If
End Sub


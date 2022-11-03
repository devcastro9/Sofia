VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form FrmExplorador 
   Appearance      =   0  'Flat
   Caption         =   "Explorando"
   ClientHeight    =   5865
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   8490
   Icon            =   "FrmExplora.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   8490
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdSeleccionar 
      Caption         =   "Seleccionar"
      Height          =   405
      Left            =   8490
      TabIndex        =   17
      Top             =   7050
      Width           =   5280
   End
   Begin VB.TextBox TxtRutaNombreArchivo 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   8490
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   1545
      Width           =   5280
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   7515
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Frame FraFecha 
      Height          =   3405
      Left            =   4005
      TabIndex        =   12
      Top             =   5295
      Visible         =   0   'False
      Width           =   3645
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   1245
         TabIndex        =   14
         Top             =   2970
         Width           =   1485
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   600
         TabIndex        =   13
         Top             =   435
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   24641537
         CurrentDate     =   36404
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5955
      Left            =   3465
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5895
      ScaleWidth      =   60
      TabIndex        =   10
      Top             =   1500
      Width           =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   0
      Top             =   15
   End
   Begin VB.Frame FraEdita 
      Height          =   6045
      Left            =   3720
      TabIndex        =   7
      Top             =   1425
      Width           =   4725
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   5685
         Left            =   75
         TabIndex        =   8
         Top             =   165
         Width           =   4590
      End
   End
   Begin VB.Frame FraElige 
      Height          =   6030
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   3240
      Begin VB.DirListBox Dir1 
         DragIcon        =   "FrmExplora.frx":0442
         Height          =   5265
         Left            =   105
         TabIndex        =   6
         Top             =   630
         Width           =   3015
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3000
      End
   End
   Begin MSComctlLib.StatusBar StBExplorador 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5490
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14888
            MinWidth        =   14888
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Picture         =   "FrmExplora.frx":0884
            TextSave        =   "12/08/2000"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Picture         =   "FrmExplora.frx":0CD8
            TextSave        =   "10:15 AM"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClBExplorador 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   2514
      _CBWidth        =   13635
      _CBHeight       =   1425
      _Version        =   "6.0.8169"
      Child1          =   "TlbMenuPrincipal"
      MinHeight1      =   330
      Width1          =   10005
      NewRow1         =   0   'False
      BandTag1        =   "celia"
      Child2          =   "TlbMenuGrafico"
      MinHeight2      =   645
      Width2          =   90000
      NewRow2         =   -1  'True
      Caption3        =   "Dirección"
      MinHeight3      =   330
      Width3          =   71115
      NewRow3         =   -1  'True
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   1125
         TabIndex        =   9
         Top             =   1065
         Width           =   10500
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   11790
         Top             =   420
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   23
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   18
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":0FF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":1448
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":189C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":1CF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":1E04
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":2480
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":2AFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":3178
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":3294
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":33A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":37FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":3C50
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":3D64
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":43E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":4A5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":50D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":5754
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExplora.frx":5868
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar TlbMenuGrafico 
         Height          =   645
         Left            =   165
         TabIndex        =   3
         Top             =   390
         Width           =   13380
         _ExtentX        =   23601
         _ExtentY        =   1138
         ButtonWidth     =   1746
         ButtonHeight    =   1138
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Atrás"
               ImageIndex      =   10
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "uno"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "dos"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Adelante"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Subir"
               ImageIndex      =   12
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cortar"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Copiar"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Pegar"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Deshacer"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Propiedades"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar TlbMenuPrincipal 
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   13380
         _ExtentX        =   23601
         _ExtentY        =   582
         ButtonWidth     =   2011
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Archivo"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   5
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Crear acceso &directo"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "e&liminar"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "&hola"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Edición"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Ver"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Ir a"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Favoritos"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Herrami"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ay&uda"
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   3720
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1500
      Visible         =   0   'False
      Width           =   1935
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "FrmExplora.frx":597C
      Height          =   4815
      Left            =   8475
      OleObjectBlob   =   "FrmExplora.frx":5990
      TabIndex        =   15
      Top             =   2145
      Width           =   5280
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "&Archivo"
      Visible         =   0   'False
      Begin VB.Menu Nuevo 
         Caption         =   "Nuevo"
         Begin VB.Menu Carpeta 
            Caption         =   "Carpeta"
         End
         Begin VB.Menu MnuCreaDirecto 
            Caption         =   "Crear Archivo &directo"
         End
      End
      Begin VB.Menu MnuPC 
         Caption         =   "Mi PC"
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "E&liminar"
      End
      Begin VB.Menu MnuCambiarNombre 
         Caption         =   "Camb&iar Nombre"
      End
      Begin VB.Menu MnuPropiedades 
         Caption         =   "&Propiedades"
      End
      Begin VB.Menu MnuRed 
         Caption         =   "Trabajar sin conexion a la red"
      End
      Begin VB.Menu MnuSepArch1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCerrar 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Visible         =   0   'False
      Begin VB.Menu Cortar 
         Caption         =   "Deshacer"
      End
      Begin VB.Menu SeparadorEd1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCortar 
         Caption         =   "Co&rtar"
         Shortcut        =   ^X
      End
      Begin VB.Menu MnuCopiar 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuPegar 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu SeparadorEd2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSelecciona 
         Caption         =   "Selecc&ionar Todo"
      End
      Begin VB.Menu MnuInvierte 
         Caption         =   "In&vertir Todo"
      End
   End
   Begin VB.Menu MnuVer 
      Caption         =   "Ver"
      Visible         =   0   'False
      Begin VB.Menu MnuBarraHerra 
         Caption         =   "Barra de Herramientas"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuBarraEstado 
         Caption         =   "Barra de Estado"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuBarraExplorador 
         Caption         =   "Barra de Explorador"
      End
      Begin VB.Menu MnuBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrandes 
         Caption         =   "Iconos Grandes"
      End
      Begin VB.Menu MnuPequenos 
         Caption         =   "Iconos Pequeños"
      End
      Begin VB.Menu MnuDetalles 
         Caption         =   "Detalles"
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "Ay&uda"
      Visible         =   0   'False
      Begin VB.Menu MnuTemas 
         Caption         =   "&Temas de Ayuda"
      End
      Begin VB.Menu MnuSepAyuda1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAcerca 
         Caption         =   "&Acerca"
      End
   End
   Begin VB.Menu MnuElige 
      Caption         =   "Elige"
      Visible         =   0   'False
      Begin VB.Menu MnuExplorar 
         Caption         =   "E&xplorar"
      End
      Begin VB.Menu MnuAbrir 
         Caption         =   "&brir"
      End
      Begin VB.Menu MnuBuscar 
         Caption         =   "&Buscar..."
      End
      Begin VB.Menu MnuSepElige1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFormato 
         Caption         =   "Dar f&ormato"
      End
      Begin VB.Menu MnuDirecto 
         Caption         =   "Crear acceso &directo"
      End
      Begin VB.Menu MnuSepElige2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPropie 
         Caption         =   "P&ropiedades"
      End
   End
   Begin VB.Menu MnuAbre 
      Caption         =   "Abre Archivos"
      Visible         =   0   'False
      Begin VB.Menu MnuAbrirCon 
         Caption         =   "Abrir C&on"
      End
      Begin VB.Menu MnuSepAbre1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEnviar 
         Caption         =   "En&viar a"
      End
      Begin VB.Menu MnuSepAbre2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCortarAbre 
         Caption         =   "C&ortar"
      End
      Begin VB.Menu MnuCopiarAbre 
         Caption         =   "&Copiar"
      End
      Begin VB.Menu MnuSepAbre3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAccesoDirecto 
         Caption         =   "Crear acceso &directo"
      End
      Begin VB.Menu MnuEliminarAbre 
         Caption         =   "E&liminar"
      End
      Begin VB.Menu MnuCambiaNombreAbre 
         Caption         =   "Camb&iar Nombre"
      End
      Begin VB.Menu MnuSepAbre4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPropiedadesAbre 
         Caption         =   "P&ropiedades"
      End
   End
End
Attribute VB_Name = "FrmExplorador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objeto As Object
Dim fso As New FileSystemObject
Private Sub Command1_Click()
    Dim RetVal
    Dim RetVal1
    RetVal = Shell("C:\WINDOWS\CALC.EXE", 1)   ' Ejecuta Calculadora.
    RetVal1 = Shell("C:\Archivos de programa\Microsoft Office\Office\excel.EXE", 1)   ' Ejecuta Calculadora.
End Sub

Private Sub CmdAceptar_Click()
    FraFecha.Visible = False
End Sub

Private Sub CmdSeleccionar_Click()
    FrmCapturaDatosBanco.Show
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
   Select Case Button
    Case "Copiar"
    Case "Pegar"
        If File1.FileName <> "" Then
            FileCopy ArchOrigen, ArchDestino & File1.FileName
            ArchDestino = Dir1.Path & "\" & File1.FileName
        End If
   End Select
 End Sub
Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MnuElige
    End If
End Sub
Private Sub Drive1_Change()
On Error GoTo valida_drive:
Dir1.Path = Drive1.Drive
valida_drive:
    If err_number = 51 Then
       MsgBox "No existe disquette"
    End If
End Sub
Private Sub Drive2_Change()
On Error GoTo valida_drive:
Dir1.Path = Drive2.Drive
valida_drive:
    If err_number = 51 Then
       MsgBox "No existe disquette"
    End If
End Sub

Private Sub File1_Click()
Dim fs, f
Set fs = CreateObject("Scripting.FileSystemObject")
Dim i As Long
Dim tem As String
    'Datos del archivo
    Dim s, X
    StBExplorador.Panels(1) = Dir1.Path & "\" & File1.FileName
    X = Dir1.Path & "\" & File1.FileName
    Set f = fs.GetFile(X)
    TxtRutaNombreArchivo.Text = f
    s = f.Path & "-"
    s = s & "Último acceso: " & f.DateLastAccessed & "-"
    s = s & "Última modificación: " & f.DateLastModified
    StBExplorador.Panels(1) = s
TEMP = UCase(Right$(File1.FileName, 3))
If File1.FileName <> "" And TEMP = "XLS" Or TEMP = "xls" Then
  Data1.Connect = "Excel 8.0"
  Data1.DatabaseName = f
  Data1.RecordSource = "Hoja1$"
  TDBGrid1.DataSource = Data1
  TDBGrid1.ReBind
  TDBGrid1.Refresh
Else
    MsgBox "No tiene el formato de EXCEL", vbCritical + vbDefaultButton1, "Validación de Datos"
End If
End Sub

Private Sub File1_DblClick()
Dim i As Long

i = ShellExecute(FrmExplorador.hwnd, "open", File1.Path & "\" & File1.FileName, vbNullString, vbNullString, SW_SHOWNORMAL)

StBExplorador.Panels(1) = File1.FileName
imagen_no_valida:
    If Err.Number = 148 Then
      Image1.Picture = LoadPicture("")
       Exit Sub
    End If
    If Err.Number = 481 Then
        MsgBox Err.Description
    End If
    If Err.Number = 53 Then
        MsgBox Err.Description
    End If
End Sub



Private Sub Form_Load()
'SetWindowText Form1.hwnd, "Bienvenidos  a VB"

	Call SeguridadSet(Me)
End Sub
Private Sub Form_Resize()
'If ScaleWidth > FraElige.Width And ScaleHeight > FraElige.Height Then
'        FraElige.Move 120, 1440, ScaleWidth - 7270, ScaleHeight - 1900
'        Dir1.Move 30, 700, ScaleWidth - 7470, ScaleHeight - 2500
'
'        FraEdita.Move 3435, 1440, ScaleWidth - 7270, ScaleHeight - 1900
'        File1.Move 30, 100, ScaleWidth - 7370, ScaleHeight - 2055
'
'        FraImagen.Move 8355, 1440, ScaleWidth - 7270, ScaleHeight - 1900
'        Image1.Move 30, 100, ScaleWidth - 7270, ScaleHeight - 1900
'
'End If
End Sub
Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
' Obtiene las tres últimas letras del nombre del archivo arrastrado.
    TEMP = Right$(File1.FileName, 3)

    ' Si el archivo arrastrado se encuentra en la raíz, agrega el nombre del archivo.
    If Mid$(File1.Path, Len(File1.Path)) = "\" Then
      dropfile = File1.Path & File1.FileName
    ' Si el archivo arrastrado no se encuentra en la raíz, agrega "\" al nombre del archivo.
    Else
      dropfile = File1.Path & "\" & File1.FileName
    End If
      
    Image1.Picture = LoadPicture("")
    Select Case UCase$(Trim$(TEMP))
        Case "TXT"
            X = Shell("Notepad " + dropfile, 1)
        Case "BMP"
            Image1.Picture = LoadPicture(dropfile)
        Case "EXE"
            X = Shell(dropfile, 1)
        Case "HLP"
            X = Shell("WinHelp " + dropfile, 1)
        Case "DOC"
            'X = Shell(windword.Path & "\" & "windword " + dropfile, 1)
            Set objeto = CreateObject("Word.Application")
            objeto.Visible = True
            objeto.Documents.Open (dropfile)
        Case "XLS"
             Set objeto = CreateObject("excel.Application")
             objeto.Visible = True
             objeto.Documents.Open (dropfile)
        Case Else
            msg = "Pruebe con uno de estos tipos de archivos:"
            msg = vbCrLf & msg & vbCrLf & vbCrLf & "     .txt, .bmp, .exe, .hlp"
            MsgBox msg
    End Select

End Sub

Private Sub Image1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Select Case State
    Case 0
        ' Presenta un icono nuevo cuando el origen entra en el área de colocar.
        File1.DragIcon = Dir1.DragIcon
    Case 1
        ' Presenta el DragIcon original cuando el origen sale del área de colocar.
        File1.DragIcon = Drive1.DragIcon
    End Select
End Sub

Private Sub MnuBarraEstado_Click()
    If MnuBarraEstado.Checked = True Then
       MnuBarraEstado.Checked = False
       StBExplorador.Visible = False
    Else
       MnuBarraEstado.Checked = True
       StBExplorador.Visible = True
    End If
End Sub

Private Sub MnuBarraHerra_Click()
    If MnuBarraHerra.Checked = True Then
       MnuBarraHerra.Checked = False
       MnuBarraHerra.Checked = False
       ClBExplorador.Bands(2).Visible = False
    Else
       MnuBarraHerra.Checked = True
       ClBExplorador.Bands(2).Visible = True
    End If
End Sub
Private Sub MnuAcerca_Click()
    FrmAcerca.Show
End Sub
Private Sub MnuTema_Click()
    MsgBox "esta en pleno desarrollo..."
End Sub
Private Sub MnuEliminar_Click()
On Error GoTo error_lectura
     Kill (Dir1.Path & "\" & File1.FileName)
     File1.Refresh
error_lectura:
  MsgBox err_number

End Sub
Private Sub Toolbar1_Click()
    If Toolbar1.Buttons(2) = "Archivo" Then
            PopupMenu mnufile1
    End If

End Sub
Private Sub MnuCerrar_Click()
    Unload Me
    End
End Sub
Private Sub MnuEliminarAbre_Click()
     Kill (Dir1.Path & "\" & File1.FileName)
     File1.Refresh
End Sub

Private Sub MnuPropie_Click()
'Para subdirectorios
Dim Carpeta As Folder
'StBExplorador.Panels(1) = Dir1.Path & "\" & File1.FileName
X = Dir1.Path
If X = "C:\" Then Exit Sub
Set Carpeta = fso.GetFolder(X)
s = Carpeta.DateCreated
FrmPropiedades.LblCreado = s
s = Carpeta.Name
FrmPropiedades.LblNOmbre = X
s = Carpeta.DateCreated
FrmPropiedades.LblUbicacion = s
s = Carpeta.Size
FrmPropiedades.LblTamano = s & " KBytes"
s = Carpeta.Type
FrmPropiedades.LblTipo = s
'If Carpeta.ReadOnly = True Then
    FrmPropiedades.Check1 = 1
'End If
'If Carpeta.Hidden = True Then
    FrmPropiedades.Check2.Value = 1
'End If
'If Carpeta.System = True Then
    FrmPropiedades.Check3.Value = 1
'End If
'If Carpeta.Archive = True Then
    FrmPropiedades.Check4.Value = 1
'End If

FrmPropiedades.Show
End Sub

Private Sub MnuPropiedadesAbre_Click()
Dim fs, f
Set fs = CreateObject("Scripting.FileSystemObject")
Dim i As Long

    'Datos del archivo
    Dim s, X
    StBExplorador.Panels(1) = Dir1.Path & "\" & File1.FileName
    X = Dir1.Path & "\" & File1.FileName
    FrmPropiedades.LblNOmbre = s
    If File1.FileName <> "" Then
        Set f = fs.GetFile(X)
        s = f.Path
        FrmPropiedades.LblUbicacion = s
        s = f.Type
        FrmPropiedades.LblTipo = s
        s = f.DateCreated
        FrmPropiedades.LblCreado = s
        s = f.DateLastModified
        FrmPropiedades.LblContiene = s
        s = f.Size
        FrmPropiedades.LblTamano = s & "Kbytes"
        FrmPropiedades.Check1 = 1
        FrmPropiedades.Check2.Value = 1
        FrmPropiedades.Check3.Value = 1
        FrmPropiedades.Check4.Value = 1

        FrmPropiedades.Show
    End If
End Sub

Private Sub MnuTemas_Click()
      dlgDialog.HelpFile = "c:\windows\system\access.hlp"
      dlgDialog.HelpCommand = cdlHelpContents
      dlgDialog.ShowHelp
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Tam As Integer
        Tam = FraElige.Width + FraEdita.Width
        Picture1.Left = Picture1.Left + X
        FraElige.Width = FraElige.Width + X
        Drive1.Width = Drive1.Width + X
        Dir1.Width = Dir1.Width + X
        FraEdita.Left = FraEdita.Left + X
        FraEdita.Width = Tam - FraElige.Width
End Sub
Private Sub StBExplorador_PanelDblClick(ByVal Panel As MSComctlLib.Panel)

  If Panel.Index = 2 Then
        FraFecha.Visible = True
  End If
End Sub

Private Sub TlbMenuGrafico_ButtonClick(ByVal Button As MSComctlLib.Button)
'Declaración de variables
Dim guardafile As String
Dim ArchOrigen, ArchDestino

On Error GoTo error_lectura
'Cambiando el gráfico de los botones pulados del toolbar
        
        If Button.Index = 1 And TlbMenuGrafico.Buttons(1).Image <> 1 Then
                TlbMenuGrafico.Buttons(1).Image = 1
                dlgDialog.ShowOpen   ' presentar el cuadro de diálogo común
        Else
                TlbMenuGrafico.Buttons(1).Image = 10
        End If
        If Button.Index = 2 And TlbMenuGrafico.Buttons(2).Image <> 2 Then
                TlbMenuGrafico.Buttons(2).Image = 2
                dlgDialog.ShowOpen   ' presentar el cuadro de diálogo común
        Else
                TlbMenuGrafico.Buttons(2).Image = 11
        End If
        If Button.Index = 3 And TlbMenuGrafico.Buttons(3).Image <> 3 Then
                TlbMenuGrafico.Buttons(3).Image = 3
                dlgDialog.ShowOpen   ' presentar el cuadro de diálogo común
        Else
                TlbMenuGrafico.Buttons(3).Image = 12
        End If
        If Button.Index = 5 And TlbMenuGrafico.Buttons(5).Image <> 4 Then
                TlbMenuGrafico.Buttons(5).Image = 4
        Else
                TlbMenuGrafico.Buttons(5).Image = 13
        End If
        If Button.Index = 6 And TlbMenuGrafico.Buttons(6).Image <> 5 Then
                TlbMenuGrafico.Buttons(6).Image = 5
        Else
                TlbMenuGrafico.Buttons(6).Image = 14
        End If
        If Button.Index = 7 And TlbMenuGrafico.Buttons(7).Image <> 6 Then
                TlbMenuGrafico.Buttons(7).Image = 6
        Else
                TlbMenuGrafico.Buttons(7).Image = 15
        End If
        If Button.Index = 8 And TlbMenuGrafico.Buttons(8).Image <> 7 Then
                TlbMenuGrafico.Buttons(8).Image = 7
        Else
                TlbMenuGrafico.Buttons(8).Image = 16
        End If
        If Button.Index = 9 And TlbMenuGrafico.Buttons(8).Image <> 8 Then
                TlbMenuGrafico.Buttons(9).Image = 7
        Else
                TlbMenuGrafico.Buttons(9).Image = 16
        End If
        
        If Button.Index = 11 And TlbMenuGrafico.Buttons(10).Image <> 9 Then
                TlbMenuGrafico.Buttons(10).Image = 8
        Else
                TlbMenuGrafico.Buttons(10).Image = 17
        End If
    Select Case Button
    Case "Copiar"
        If File1.FileName <> "" Then
            ArchOrigen = Dir1.Path & "\" & File1.FileName
            ArchDestino = "c:\basurero\" & File1.FileName
            FileCopy ArchOrigen, ArchDestino
'            ArchDestino = Dir1.Path & "\" & File1.FileName
            Text1.Text = File1.FileName
        End If
    Case "Pegar"
            ArchOrigen = "c:\basurero\" & Trim(Text1.Text)
            ArchDestino = Dir1.Path & "\" & Text1.Text
            FileCopy ArchOrigen, ArchDestino
            File1.Refresh
    Case "Cortar"
        If File1.FileName <> "" Then
           guardafile = Dir1.Path & "\" & File1.FileName
           Text1.Text = File1.FileName
'           FileCopy guardafile, "c:\basurero" & file.FileName
           Kill (Dir1.Path & "\" & File1.FileName)
           File1.Refresh
        End If
    Case "Eliminar"
        If File1.FileName <> "" Then
            If File1.FileName <> "" Then
                guardafile = Dir1.Path & "\" & File1.FileName
                Kill (Dir1.Path & "\" & File1.FileName)
                File1.Refresh
            End If
        Else
            MsgBox "Hacer click sobre el archivo a eliminarse", vbCritical, "Mensaje"
            
        End If
    End Select
error_lectura:
  If err_number = 1 Then
      MsgBox "Introducir destino", vbOKCancel
      
  End If
  
End Sub
Private Sub TlbMenuPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button = "&Archivo" Then
      PopupMenu MnuArchivo
End If
If Button = "&Edición" Then
    PopupMenu mnuEdicion
End If
If Button = "&Ver" Then
    PopupMenu MnuVer
End If
If Button = "Ay&uda" Then
    PopupMenu MnuAyuda
End If
End Sub

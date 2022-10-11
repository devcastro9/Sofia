VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form FrmExplorador 
   Caption         =   "Captura de Datos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmExplorador.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   4620
      TabIndex        =   13
      Top             =   0
      Width           =   4680
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDAD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   60
         TabIndex        =   18
         Top             =   675
         Width           =   1110
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1245
         TabIndex        =   17
         Top             =   705
         Width           =   2460
      End
      Begin VB.Label Label6 
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   9210
         TabIndex        =   16
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   15
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATOS DE HOJA ELECTRONICA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   360
         Left            =   3705
         TabIndex        =   14
         Top             =   135
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   0
         Picture         =   "FrmExplorador.frx":0ECA
         Top             =   0
         Width           =   11640
      End
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   3435
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1230
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   5955
      Left            =   3135
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5895
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   1170
      Width           =   120
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   7185
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.TextBox TxtRutaNombreArchivo 
      Appearance      =   0  'Flat
      Height          =   555
      Left            =   7440
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1185
      Width           =   4425
   End
   Begin VB.CommandButton CmdSeleccionar 
      Caption         =   "Seleccionar"
      Height          =   450
      Left            =   7440
      TabIndex        =   2
      Top             =   6705
      Width           =   1485
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   450
      Left            =   10410
      TabIndex        =   1
      Top             =   6705
      Width           =   1455
   End
   Begin VB.CommandButton CmdOperaciones 
      Caption         =   "Operar"
      Height          =   450
      Left            =   8925
      TabIndex        =   0
      Top             =   6705
      Width           =   1485
   End
   Begin MSComctlLib.StatusBar StBExplorador 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2820
      Width           =   4680
      _ExtentX        =   8255
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
            Picture         =   "FrmExplorador.frx":27F3A
            TextSave        =   "01/12/2002"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Picture         =   "FrmExplorador.frx":2838E
            TextSave        =   "12:23 PM"
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "FrmExplorador.frx":286AA
      Height          =   4890
      Left            =   7440
      OleObjectBlob   =   "FrmExplorador.frx":286BE
      TabIndex        =   12
      Top             =   1785
      Width           =   4425
   End
   Begin VB.Frame FraElige 
      Height          =   6030
      Left            =   45
      TabIndex        =   7
      Top             =   1095
      Width           =   3045
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2865
      End
      Begin VB.DirListBox Dir1 
         DragIcon        =   "FrmExplorador.frx":2ADC0
         Height          =   5265
         Left            =   120
         TabIndex        =   8
         Top             =   615
         Width           =   2850
      End
   End
   Begin VB.Frame FraEdita 
      Height          =   6045
      Left            =   3360
      TabIndex        =   5
      Top             =   1095
      Width           =   4050
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   5685
         Left            =   60
         TabIndex        =   6
         Top             =   150
         Width           =   3885
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

Private Sub cmdAceptar_Click()
    FraFecha.Visible = False
End Sub

Private Sub CmdOperaciones_Click()
  FrmExtractoBancario.Show
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
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
    If err.Number = 148 Then
      Image1.Picture = LoadPicture("")
       Exit Sub
    End If
    If err.Number = 481 Then
        MsgBox err.Description
    End If
    If err.Number = 53 Then
        MsgBox err.Description
    End If
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
FrmPropiedades.Check1 = 1
FrmPropiedades.Check2.Value = 1
FrmPropiedades.Check3.Value = 1
FrmPropiedades.Check4.Value = 1
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
'
'  If Panel.Index = 2 Then
'        FraFecha.Visible = True
'  End If
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


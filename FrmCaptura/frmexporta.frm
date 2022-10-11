VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Frmexporta 
   BackColor       =   &H8000000A&
   Caption         =   "Copia de Archivos"
   ClientHeight    =   5685
   ClientLeft      =   3870
   ClientTop       =   3330
   ClientWidth     =   7755
   ControlBox      =   0   'False
   Icon            =   "frmexporta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7755
   Begin VB.CommandButton BtnGrabar 
      BackColor       =   &H80000010&
      Caption         =   "Copiar"
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
      Height          =   720
      Left            =   3240
      Picture         =   "frmexporta.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3960
      Width           =   1245
   End
   Begin VB.CommandButton BtnSalir 
      BackColor       =   &H80000010&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3240
      Picture         =   "frmexporta.frx":0E44
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Width           =   1245
   End
   Begin VB.FileListBox File2 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   4560
      TabIndex        =   12
      Top             =   3960
      Width           =   3015
   End
   Begin VB.DirListBox DirDestino2 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   4560
      TabIndex        =   11
      Top             =   2160
      Width           =   3015
   End
   Begin VB.DirListBox DirDestino 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   4560
      TabIndex        =   10
      Top             =   480
      Width           =   3015
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   240
      TabIndex        =   9
      Top             =   3960
      Width           =   2910
   End
   Begin VB.DirListBox DirOrigen 
      Height          =   1890
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   1320
      Width           =   2910
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   2895
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Guardando archivo de Exportación"
      Filter          =   "*.txt"
   End
   Begin MSAdodcLib.Adodc adosolicitud 
      Height          =   330
      Left            =   1920
      Top             =   5520
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Adodc1"
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
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TMPBANCOTXT"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label LblFA 
      Caption         =   "Label6"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3600
      Picture         =   "frmexporta.frx":104E
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3600
      Picture         =   "frmexporta.frx":1358
      Top             =   2280
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Carpeta Origen:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo Origen:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Disco Origen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Carpeta Destino Local:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo Destino:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Carpeta Destino Servidor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label LblFormname 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   5655
   End
End
Attribute VB_Name = "Frmexporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstAo_solicitud As New ADODB.Recordset
Dim fs As FileSystemObject      'Variable de tipo file System Object
'Dim fso As FileSystemObject      'Variable de tipo file System Object
Dim a
Dim sino As String

Private Sub BtnSalir_Click()
   Unload Me
End Sub

Private Sub BtnGrabar_Click()
 On Error GoTo Error_Sub
   Set fs = New FileSystemObject   'Creamos la Nueva referencia Fso
   'Set fs = CreateObject("Scripting.FileSystemObject")
   
'   sino = MsgBox("Desea Borrar los datos copiados anteriormente en -->" & DirDestino, vbYesNo + vbQuestion, "Atención")
'   If sino = vbYes And Err.Number = 0 Then
'        'fs.DeleteFile DirDestino & "\*.*", True
' ************************* PARAMETRIZAR ********************************
    If GlArch = "DEDF" Then
      fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & aw_p_ao_negociacion_cabecera.Ado_datos.Recordset!archivo_respaldo
     'gw_edificaciones.Ado_datos.Recordset!ARCHIVO_Foto = gw_edificaciones.Ado_datos.Recordset!ARCHIVO_Foto
     aw_p_ao_negociacion_cabecera.Ado_datos.Recordset!archivo_respaldo_cargado = "S"
'     If GlServidor = "SRVPRO" Then
'        fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino2 & "\" & frmBeneficiario.adoLista.Recordset!ARCHIVO_F
'     End If
    End If
    If GlArch = "DED2" Then
      fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & mw_solicitud.Ado_datos.Recordset!archivo_respaldo
     'gw_edificaciones.Ado_datos.Recordset!ARCHIVO_Foto = gw_edificaciones.Ado_datos.Recordset!ARCHIVO_Foto
     mw_solicitud.Ado_datos.Recordset!archivo_respaldo_cargado = "S"
'     If GlServidor = "SRVPRO" Then
'        fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino2 & "\" & frmBeneficiario.adoLista.Recordset!ARCHIVO_F
'     End If
    End If
   If GlArch = "Q_R" Then
'     fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & Frm_ao_ventas_cobranzas.Ado_datos.Recordset!ARCHIVO_Foto
'     frmBeneficiario_Admin.Adolista.Recordset!archivo_foto_cargado = "S"
'     'frmBeneficiario_Admin.Ado_datos.Recordset!ARCHIVO_Foto = frmBeneficiario_Admin.Ado_datos.Recordset!ARCHIVO_Foto
'     LblFormname.Caption = DirOrigen & "\" & File1.FileName
''     If GlServidor = "SRVPRO" Then
''        fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino2 & "\" & frmBeneficiario.adoLista.Recordset!ARCHIVO_F
''     End If
   End If
   If GlArch = "FOTE" Then
'     fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & frmBeneficiarioEmp.adoLista.Recordset!ARCHIVO_F
'     frmBeneficiarioEmp.adoLista.Recordset!ARCHIVO_FOTO = frmBeneficiarioEmp.adoLista.Recordset!ARCHIVO_F
''     If GlServidor = "SRVPRO" Then
''        fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino2 & "\" & frmBeneficiario.adoLista.Recordset!ARCHIVO_F
''     End If
   End If
   If GlArch = "FOTB" Then
'     fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & AlFrmCreaMaterial.AdoArt.Recordset!ARCHIVO_F
'     AlFrmCreaMaterial.AdoArt.Recordset!ARCHIVO_FOTO = AlFrmCreaMaterial.AdoArt.Recordset!ARCHIVO_F
''     If GlServidor = "SRVPRO" Then
''        fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino2 & "\" & frmBeneficiario.adoLista.Recordset!ARCHIVO_F
''     End If
   End If
   If GlArch = "FEDF" Then
     fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & gw_edificaciones.Ado_datos.Recordset!ARCHIVO_Foto
     'gw_edificaciones.Ado_datos.Recordset!ARCHIVO_Foto = gw_edificaciones.Ado_datos.Recordset!ARCHIVO_Foto
     gw_edificaciones.Ado_datos.Recordset!archivo_foto_cargado = "S"
'     If GlServidor = "SRVPRO" Then
'        fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino2 & "\" & frmBeneficiario.adoLista.Recordset!ARCHIVO_F
'     End If
   End If
   If GlArch = "FOT1" Then
'     fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & frmBeneficiario_Admin.Adolista.Recordset!ARCHIVO_Foto
'     'fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & frmBeneficiario_Admin.adoLista.Recordset!ARCHIVO_Foto
'     frmBeneficiario_Admin.Adolista.Recordset!archivo_foto_cargado = "S"
'     'frmBeneficiario_Admin.Ado_datos.Recordset!ARCHIVO_Foto = frmBeneficiario_Admin.Ado_datos.Recordset!ARCHIVO_Foto
'     'LblFormname.Caption = DirOrigen & "\" & File1.FileName
''     If GlServidor = "SRVPRO" Then
''        fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino2 & "\" & frmBeneficiario.adoLista.Recordset!ARCHIVO_F
''     End If
   End If
   If GlArch = "FED2" Then
     fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & mw_solicitud.Ado_detalle1.Recordset!ARCHIVO_Foto
     mw_solicitud.Ado_detalle1.Recordset!archivo_foto_cargado = "S"
'     If GlServidor = "SRVPRO" Then
'        fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino2 & "\" & frmBeneficiario.adoLista.Recordset!ARCHIVO_F
'     End If
   End If
   If GlArch = "C_V" Then
'     fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & frmBeneficiario_Admin.Adolista.Recordset!archivo_hojavida
'     frmBeneficiario_Admin.Adolista.Recordset!archivo_hojavida_cargado = "S"
'     'frmBeneficiario.adoLista.Recordset!archivo_hojavida = frmBeneficiario.adoLista.Recordset!ARCHIVO_HV
''     If GlServidor = "SRVPRO" Then
''        fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino2 & "\" & frmBeneficiario.adoLista.Recordset!ARCHIVO_HV
''     End If
   End If
   If GlArch = "D_R" Then
'     fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & frmBeneficiario.adoLista.Recordset!ARCHIVO_RESP
'     frmBeneficiario.adoLista.Recordset!ARCHIVO_RESPALDO = frmBeneficiario.adoLista.Recordset!ARCHIVO_RESP
''     If GlServidor = "SRVPRO" Then
''        fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino2 & "\" & frmBeneficiario.adoLista.Recordset!ARCHIVO_RESP
''     End If
   End If
   If GlArch = "CTO" Then
'     fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & Ac_Personal_Contrato.Ado_Auxiliar.Recordset!ARCHIVO_NOMB
'     Ac_Personal_Contrato.Ado_Auxiliar.Recordset!ARCHIVO = Ac_Personal_Contrato.Ado_Auxiliar.Recordset!ARCHIVO_NOMB
   End If
   
   If GlArch = "ASIS" Then
     Dim ExtFile As String
     ExtFile = Mid(File1.FileName, InStrRev(File1.FileName, ".") + 1, Len(File1.FileName))
     
     GlExtension = ExtFile
     ' Asigna nombre de archivo a variable global.
     'GLCarpeta = Replace(File1.FileName, ".xls", "")
     GLCarpeta = Replace(File1.FileName, "." + ExtFile, "")
     
     ' GLCarpeta2 contiene el nombre del archivo.
     ' fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & GLCarpeta2 & ".xls"
     fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & GLCarpeta2 & "." + ExtFile
     Unload Me
    
   End If
 
   If GlArch = "EXEL" Then
     Dim ExtFile2 As String
     ExtFile2 = Mid(File1.FileName, InStrRev(File1.FileName, ".") + 1, Len(File1.FileName))
     
     GlExtension = ExtFile2
     ' Asigna nombre de archivo a variable global.
     'GLCarpeta = Replace(File1.FileName, ".xls", "")
     GLCarpeta = Replace(File1.FileName, "." + ExtFile2, "")
     
     ' GLCarpeta2 contiene el nombre del archivo.
     ' fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & GLCarpeta2 & ".xls"
     fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & GLCarpeta2 & "." + ExtFile2
     Unload Me
    
   End If
    
   If GlArch = "EXBCO" Then
     Dim ExtFile3 As String
     ExtFile3 = Mid(File1.FileName, InStrRev(File1.FileName, ".") + 1, Len(File1.FileName))
     
     GlExtension = ExtFile3
     ' Asigna nombre de archivo a variable global.
     'GLCarpeta = Replace(File1.FileName, ".xls", "")
     GLCarpeta = Replace(File1.FileName, "." + ExtFile3, "")
     
     ' GLCarpeta2 contiene el nombre del archivo.
     ' fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & GLCarpeta2 & ".xls"
     fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino & "\" & GLCarpeta2 & "." + ExtFile3
     Unload Me
    
   End If
    
'   End If
'   sino = MsgBox("Desea Borrar los datos copiados anteriormente en -->" & DirDestino2, vbYesNo + vbQuestion, "Atención")
'   If sino = vbYes And Err.Number = 0 Then
'        'fs.DeleteFile DirDestino2 & "\*.*", True
'        'fs.CopyFile DirOrigen & "\*.*", DirDestino2 & "\" & Ac_Personal_Contrato.Ado_Auxiliar.Recordset!ARCHIVO_NOMB
'SERVIDOR
'       fs.CopyFile DirOrigen & "\" & File1.FileName, DirDestino2 & "\" & Ac_Personal_Contrato.Ado_Auxiliar.Recordset!ARCHIVO_NOMB
'   End If
   
   'fs.CopyFile DirOrigen & "\" & File1.FileName & "*", DirDestino
   File2.Path = " "
   'File2.Path = DirDestino.Path
   File2.Refresh
   BtnGrabar.Enabled = False
Exit Sub
Error_Sub:
 MsgBox Err.Description, vbCritical
End Sub

Private Sub DirDestino_Change()
   File2.Path = DirDestino.Path
   File2.Refresh
End Sub

Private Sub DirOrigen_Change()
   File1.Path = " "
   File1.Path = DirOrigen.Path
End Sub

Private Sub Drive1_Change()
    'DirDestino.Path = Drive1.Drive
    'DirDestino.Refresh
    DirOrigen.Path = Drive1.Drive
    DirOrigen.Refresh
End Sub

Private Sub File1_Click()
   If Len(File1.FileName) > 0 Then
        BtnGrabar.Enabled = True
   Else
        BtnGrabar.Enabled = False
   End If
End Sub

Private Sub Form_Activate()
    DirDestino.Enabled = False
    BtnGrabar.Enabled = True
End Sub

'Private Sub Form_Load()
'    Set rstAo_solicitud = New ADODB.Recordset
'    If rstAo_solicitud.State = 1 Then rstAo_solicitud.Close
'    rstAo_solicitud.Open "select * from Ao_solicitud ", db, adOpenKeyset, adLockOptimistic    'where tipo_formulario = 'F01'
'    Set adosolicitud.Recordset = rstAo_solicitud
'   adosolicitud.Refresh
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set a = fs.CreateTextFile(App.Path & "\TMPBANCOTXT.txt", True)   'fs.CreateTextFile("c:\greco\captura\TMPBANCOTXT.txt", True)
''    a.WriteLine ("Pruebita de grabar texto")
'    a.Close
'    Data1.DatabaseName = " "
'    Data1.DatabaseName = App.Path & "\pragma.mdb"  '"C:\greco\Captura\pragma.mdb"
'    Data1.RecordSource = "TMPBANCOTXT"
'    Data1.Refresh
'    If (Not adosolicitud.Recordset.BOF) And (Not adosolicitud.Recordset.EOF) Then adosolicitud.Recordset.MoveFirst
'    While Not adosolicitud.Recordset.EOF
'        Data1.Recordset.AddNew
'        Data1.Recordset("campo1") = adosolicitud.Recordset("ges_gestion")
'        Data1.Recordset("campo2") = adosolicitud.Recordset("tipo_formulario")
'        Data1.Recordset("campo3") = adosolicitud.Recordset("codigo_unidad")
'        Data1.Recordset("campo4") = adosolicitud.Recordset("justificacion_solicitud")
'        Data1.Recordset("campo5") = adosolicitud.Recordset("ci")
'        Data1.Recordset("campo6") = adosolicitud.Recordset("Codigo_puesto")
'        Data1.Recordset("campo7") = adosolicitud.Recordset("CI_aprueba")
'        Data1.Recordset("campo8") = adosolicitud.Recordset("estado_aprobacion")
'        Data1.Recordset("campo9") = adosolicitud.Recordset("codigo_poa")
'        Data1.Recordset("campo10") = adosolicitud.Recordset("tipo_moneda")
'
'        Data1.Recordset("campo11") = adosolicitud.Recordset("codigo_solicitud")
'        Data1.Recordset("campo12") = adosolicitud.Recordset("importe_bolivianos")
'        Data1.Recordset("campo13") = adosolicitud.Recordset("Importe_dolares")
'        Data1.Recordset("campo14") = adosolicitud.Recordset("Tipo_cambio")
'        Data1.Recordset("campo15") = adosolicitud.Recordset("duracion_estimada_numero")
'        Data1.Recordset("campo15") = adosolicitud.Recordset("por_tiempo")
'
'        Data1.Recordset.Update
'        adosolicitud.Recordset.MoveNext
'    Wend
''  fs.
'   DirOrigen.Path = App.Path   '"c:\greco\captura"   '"\\Sersis\saf\copia"
'   File1.Path = " "
'   File1.Path = DirOrigen.Path

'   DirOrigen.Path = App.Path & "\FA-2010-00" & File1.FileName
'   File1.Path = " "
'   File1.Path = DirOrigen.Path
'   fs.CopyFile DirOrigen & "\" & File1.FileName
'End Sub


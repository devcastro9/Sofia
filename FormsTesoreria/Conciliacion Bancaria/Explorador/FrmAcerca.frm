VERSION 5.00
Begin VB.Form FrmAcerca 
   Caption         =   "Acerca de Windows"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdInformacionSistema 
      Caption         =   "Informacion Sistema"
      Height          =   390
      Left            =   3840
      TabIndex        =   4
      Top             =   2295
      Width           =   1170
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   405
      Left            =   3840
      TabIndex        =   0
      Top             =   2820
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   45
      Picture         =   "FrmAcerca.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   990
   End
   Begin VB.Label LblAcerca3 
      BackStyle       =   0  'Transparent
      Height          =   420
      Left            =   1185
      TabIndex        =   3
      Top             =   2265
      Width           =   4005
   End
   Begin VB.Label LblAcerca2 
      BackStyle       =   0  'Transparent
      Height          =   825
      Left            =   1185
      TabIndex        =   2
      Top             =   1245
      Width           =   3930
   End
   Begin VB.Label LblAcerca1 
      BackStyle       =   0  'Transparent
      Height          =   1035
      Left            =   1185
      TabIndex        =   1
      Top             =   180
      Width           =   3930
   End
End
Attribute VB_Name = "FrmAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Opciones de seguridad de clave del Registro...
Const KEY_ALL_ACCESS = &H2003F
                                          

' Tipos ROOT de claves del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' cadena terminada en valor nulo Unicode
Const REG_DWORD = 4                      ' número de 32 bits


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub CmdAceptar_Click()
    Unload Me
End Sub

Private Sub CmdInformacionSistema_Click()
  InformacionSistema
End Sub

Private Sub Form_Load()
    LblAcerca1.Caption = "Microsoft (R) Window" & _
    "Windows 95" & _
    "(C) 1981 - 1987 Microsoft Coporation"
    
    LblAcerca2.Caption = "Se autoriza el uso de este producto a" & _
    "." & _
    ".."
    
    LblAcerca3.Caption = "Memoria Física disponible para window:  32.260 KB" & _
    "Recursos del sistema:   56% disponible"
    
	Call SeguridadSet(Me)
End Sub
Public Sub InformacionSistema()
    On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        

        ' Intentar obtener el nombre y la ruta del programa en el Registro...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Intentar obtener sólo la ruta del programa en el Registro...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' Validar la existencia de versión conocida de 32 bits de archivo
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' Error: no se encuentra el archivo...
                Else
                        GoTo SysInfoErr
                End If
        ' Error: no se encuentra la entrada del Registro...
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "La información del sistema no está disponible en este momento", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Contador de bucle
        Dim rc As Long                                          ' Código de retorno
        Dim hKey As Long                                        ' Controlador a una clave de Registro abierta
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Tipo de datos de una clave del Registro
        Dim tmpVal As String                                    ' Almacenamiento temporal para un valor de clave del Registro
        Dim KeyValSize As Long                                  ' Tamaño de variable de clave del Registro
        '------------------------------------------------------------
        ' Abrir RegKey bajo KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abrir la clave del Registro
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar error...
        

        tmpVal = String$(1024, 0)                             ' Asignar espacio de variable
        KeyValSize = 1024                                       ' Marcar tamaño de variable
        

        '------------------------------------------------------------
        ' Obtener valor de clave del Registro...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar errores
        

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' Determinar el tipo de valor de clave para conversión...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Buscar tipos de datos...
        Case REG_SZ                                             ' Tipo de datos String de clave del Registro
                KeyVal = tmpVal                                     ' Copiar valor String
        Case REG_DWORD                                          ' Tipo de datos Double Word de clave del Registro
                For i = Len(tmpVal) To 1 Step -1                    ' Convertir cada bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Generar valor carácter a carácter
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word a String
        End Select
        

        GetKeyValue = True                                      ' Operación realizada correctamente
        rc = RegCloseKey(hKey)                                  ' Cerrar clave del Registro
        Exit Function                                           ' Salir
        

GetKeyError:    ' Limpiar después de que se produzca un error...
        KeyVal = ""                                             ' Establecer el valor de retonor a la cadena vacía
        GetKeyValue = False                                     ' La operación no se ha realizado correctamente
        rc = RegCloseKey(hKey)                                  ' Cerrar clave del Registro
End Function



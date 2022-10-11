Attribute VB_Name = "Module1"
Option Explicit
Public Const SW_SHOWNORMAL = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim ArchOrigen, ArchDestino, a, b As String

Sub main()
    FrmExplorador.Show
End Sub

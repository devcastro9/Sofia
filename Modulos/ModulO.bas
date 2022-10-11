Attribute VB_Name = "Module1"
Option Explicit
  
' Función para almacenar el gráfico
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Guardar_Imagen(db As ADODB.Connection, _
                               sql As String, _
                               Campo_Imagen As String, _
                               Optional Path_Imagen As String) As Boolean
  
Dim rs As Recordset
Dim Stream As ADODB.Stream
  
On Error GoTo Error_Sub
  
    Set rs = New ADODB.Recordset
       
       
    ' Abre el recordset
    rs.Open sql, db, adOpenKeyset, adLockOptimistic
       
    'Nuevo objeto ADODB Stream
    Set Stream = New ADODB.Stream
       
    ' dato de tipo binario
    Stream.Type = adTypeBinary
       
    Stream.Open
        ' verifica que la ruta del gráfico no sea una cadena vacía
        If Len(Path_Imagen) <> 0 Then
               
            ' lee la imagen desde el path
            Stream.LoadFromFile Path_Imagen
            ' La guarda en campo
            rs.Fields(Campo_Imagen).Value = Stream.Read
               
            ' actualiza los cambios
            rs.Update
        End If
    ' cierra el recordset y elimina la referencia
    If rs.State = adStateOpen Then
        rs.Close
    End If
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    'Cierra el objeto ADODBStream y elimina la ref.
    If Stream.State = adStateOpen Then
        Stream.Close
    End If
    If Not Stream Is Nothing Then
        Set Stream = Nothing
    End If
    ' Retorno
    Guardar_Imagen = True
       
Exit Function
  
Error_Sub:
  If Err.Number <> 0 Then
    MsgBox CStr(Err) & "  " & Error, vbExclamation
  End If
     
End Function
  
'Función para Obtener la imagen de la bd
  
'''''''''''''''''''''''''''''''''''''''''''
Public Function Leer_Imagen(ADO_Connection As ADODB.Connection, _
                            sql As String, _
                            Campo_Imagen As String) As Picture
  
On Error GoTo error_Function
  
Dim rs As Recordset
Dim Stream As ADODB.Stream
  
  
    Set rs = New Recordset
  
  
    ' Llena el recordset
    rs.Open sql, ADO_Connection, adOpenKeyset, adLockOptimistic
       
    ' Si no hay registros sale de la función y retorna como _
     resultado un valor Nothing, es decir ninguna imagen
  
    If rs.RecordCount = 0 Then
       Set Leer_Imagen = Nothing
       rs.Close
       Set rs = Nothing
       Exit Function
    End If
       
    ' Nuevo objeto Stream para poder leer el campo de imagen
    Set Stream = New ADODB.Stream
       
    ' Especifica el tipo de datos ( binario )
    Stream.Type = adTypeBinary
    Stream.Open
       
    ' verifica con la función IsNull que el campo no tenga _
     un valor Nulo ya que si no da error, en ese caso sale de la función
    If IsNull(rs.Fields(Campo_Imagen).Value) Then
        GoTo error_Function
    End If
    ' Graba los datos en el objeto stream
    Stream.Write rs.Fields(Campo_Imagen).Value
       
    ' este método graba un  archivo temporal  en disco _
     ( en el app.path que luego se elimina )
    Stream.SaveToFile App.Path & "\temp", adSaveCreateOverWrite
       
    ' Retorna la imagen a la función
    Set Leer_Imagen = LoadPicture(App.Path & "\temp")
       
       
    ' Elimina el archivo temporal
    Kill App.Path & "\temp"
       
       
    'Cierra el recordset y el objeto Stream
    If rs.State = adStateOpen Then
        rs.Close
    End If
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
       
    If Stream.State = adStateOpen Then
        Stream.Close
    End If
    If Not Stream Is Nothing Then
        Set Stream = Nothing
    End If
  
Exit Function
  
error_Function:
 If Err.Number <> 0 Then
    MsgBox CStr(Err) & "  " & Error, vbExclamation
    ' elimina el temporal
    If Len(Dir(App.Path & "\temp")) Then
       Kill App.Path & "\temp"
    End If
 End If
  
End Function


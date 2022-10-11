Attribute VB_Name = "mod_librerias"
Public Function GetValorGeneral(cadena_query As String) As String
Dim ret As String
  Datos.dbo_apGeneralSearching cadena_query
  With Datos.rsdbo_apGeneralSearching
    If Not Datos.rsdbo_apGeneralSearching.EOF Then
      ret = Datos.rsdbo_apGeneralSearching!retorno
    End If
   .Close
  End With
 GetValorGeneral = ret
End Function

Public Function GetValor(tabla, col_buscada, col_busqueda, valor As String) As String
Dim ret As String
  Datos.dbo_apGeneralSearching "SELECT " & col_buscada & " as retorno From " & tabla & " where " & col_busqueda & " = '" & valor & "'"
  With Datos.rsdbo_apGeneralSearching
    If Not Datos.rsdbo_apGeneralSearching.EOF Then
      ret = Datos.rsdbo_apGeneralSearching!retorno
    End If
   .Close
  End With
 GetValor = ret
End Function

Function GetArg(ByVal linea As String, numero_arg As Integer) As String
Dim ret, c As String, i, actual_arg As Integer
  actual_arg = 1
  For i = 1 To Len(linea)
    c = Mid(linea, i, 1)
    If c = " " Or c = vbTab Then
      If actual_arg = numero_arg Then
        Exit For
      Else
        ret = ""
        actual_arg = actual_arg + 1
      End If
    End If
    If c <> " " Then
      ret = ret & c
    End If
  Next i
  If actual_arg = numero_arg Then
    GetArg = ret
  End If
End Function

Public Function getNumber(mensaje, titulo As String) As Double
Dim ret As Double, retInput As String, ok As Boolean
  ok = True
  retInput = ""
  While ok
    If retInput = "" Then
      retInput = InputBox(mensaje, titulo)
      If esNumero(retInput) Then
        ret = retInput
        ok = False
      Else
        retInput = ""
      End If
    End If
  Wend
  getNumber = retInput
End Function

Public Function esNumero(valor As String) As Boolean
Dim numero As Double
  On Error GoTo ControlError
    numero = CDbl(valor)
    esNumero = True
    Exit Function
ControlError:
  esNumero = False
End Function

Public Function literal(Cadena As String) As String
Dim SW As Integer
Dim sw1 As Integer
Dim swc, X, K, c, i As Integer
Dim VEC(20) As Long
Dim Z, D, S, A, j As String
SW = 0
      '*********PARTE DECIMAL*********
            Cadena = Round(Cadena, 2)
             X = Len(Cadena)
              For K = 1 To X
                  Z = Mid(Cadena, K, 1)
                  If (Z = ".") Or SW = 1 Then
                    D = D + Mid(Cadena, K, 1)
                    SW = 1
                  End If
              Next K
              
              D = Mid(D, 2, Len(D))
              
              'Para la parte decimal del monto
              If D = "00" Or D = "" Then
                 D = D & " 00/100"
              Else
                 If D >= 0 And D <= 9 And Len(D) = 1 Then
                    D = " " & D & "0" & "/100"
                 Else
                    D = " " & D & "/100 "
                 End If
              End If
      '*********PARTE ENTERA*********
 If Cadena <> "" Then
    Cadena = Int(Cadena)
 Else
    MsgBox "No existe monto"
 End If
   S = ""
   Z = ""
   c = 0
   K = 0
   sw1 = 0
   swc = 0
   
   
   X = Len(Cadena)
   For i = 1 To X
       A = Mid(Cadena, i, 1)
       VEC(i) = Mid(Cadena, i, 1)
   Next i
j = X
While j <> 0
K = K + 1
If K <> 8 Then
  If c <> 3 Then
       c = c + 1
      
       If c = 1 And (VEC(j - 1) <> 1 And VEC(j - 1) <> 2) Then
            Select Case VEC(j)
                Case 0: S = " " + S
                Case 1:
                   If sw1 <> 1 Then
                      S = "UNO " + Z + S
                   End If
                   If sw1 = 1 Then
                      S = "UN " + Z + S
                   End If
                   
                Case 2: S = "DOS " + Z + S
                Case 3: S = "TRES " + Z + S
                Case 4: S = "CUATRO " + Z + S
                Case 5: S = "CINCO " + Z + S
                Case 6: S = "SEIS " + Z + S
                Case 7: S = "SIETE " + Z + S
                Case 8: S = "OCHO " + Z + S
                Case 9: S = "NUEVE " + Z + S
          End Select
          
           'If J + 1 <> "" And sw1 <> 1 And VEC(J - 1) <> 0 And VEC(J) <> 0 Then
           If VEC(j - 1) <> 0 And VEC(j) <> 0 Then
                 S = "Y " + S
           End If
           
        End If
        
         If c = 2 And VEC(j) = 1 Then
               swc = 1
                Select Case VEC(j + 1)
                      Case 0: S = "DIEZ " + Z + S
                      Case 1: S = "ONCE " + Z + S
                      Case 2: S = "DOCE " + Z + S
                      Case 3: S = "TRECE " + Z + S
                      Case 4: S = "CATORCE " + Z + S
                      Case 5: S = "QUINCE " + Z + S
                      Case 6: S = "DIECISEIS " + Z + S
                      Case 7: S = "DIECISIETE " + Z + S
                      Case 8: S = "DIECIOCHO " + Z + S
                      Case 9: S = "DIECINUEVE " + Z + S
                End Select
          End If
          
          If c = 2 And VEC(j) = 2 Then
                Select Case VEC(j + 1)
                      Case 0: S = "VEINTE " + Z + S
                      Case 1: S = "VEINTIUNO " + Z + S
                      Case 2: S = "VEINTIDOS " + Z + S
                      Case 3: S = "VEINTITRES " + Z + S
                      Case 4: S = "VEINTICUATRO " + Z + S
                      Case 5: S = "VEINTICINCO " + Z + S
                      Case 6: S = "VEINTISEIS " + Z + S
                      Case 7: S = "VEINTISIETE " + Z + S
                      Case 8: S = "VEINTIOCHO " + Z + S
                      Case 9: S = "VEINTINUEVE " + Z + S
                End Select
          End If
   
        If c = 2 Then
            Select Case VEC(j)
                Case 3: S = "TREINTA " + Z + S
                Case 4: S = "CUARENTA " + Z + S
                Case 5: S = "CINCUENTA " + Z + S
                Case 6: S = "SESENTA " + Z + S
                Case 7: S = "SETENTA " + Z + S
                Case 8: S = "OCHENTA " + Z + S
                Case 9: S = "NOVENTA " + Z + S
            End Select
            
        End If
        
        If c = 3 Then
            Select Case VEC(j)
                Case 1:
                If j = 1 Then
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       S = "CIEN " + Z + S
                    Else
                       S = "CIENTO " + Z + S
                    End If
                Else
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       S = "CIEN " + Z + S
                    Else
                       S = "CIENTO " + Z + S
                    End If
                       'S = "CIENTO " + z + S
                End If
                Case 2: S = "DOSCIENTOS " + Z + S
                Case 3: S = "TRESCIENTOS " + Z + S
                Case 4: S = "CUATROCIENTOS " + Z + S
                Case 5: S = "QUINIENTOS " + Z + S
                Case 6: S = "SEISCIENTOS " + Z + S
                Case 7: S = "SETECIENTOS " + Z + S
                Case 8: S = "OCHOCIENTOS " + Z + S
                Case 9: S = "NOVECIENTOS " + Z + S
            End Select
        End If
   Else
     If j >= 3 Then
            If VEC(j) = 0 And VEC(j - 1) = 0 And VEC(j - 2) = 0 Then
            Else
              S = "MIL " + S
            End If
    Else
              S = "MIL " + S
    End If
        j = j + 1
        c = 0
        sw1 = 1
   End If
 Else
    If VEC(j) <> 1 Then
      S = "MILLONES " + S
    Else
'      If K > 7 Then
      If K <> 8 Then
        S = "MILLONES " + S
      Else
        S = "MILLON " + S
      End If
    End If
      j = j + 1
      c = 0
      sw1 = 1
 End If
   j = j - 1
   
Wend

literal = S + D
End Function


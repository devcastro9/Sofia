Attribute VB_Name = "CodigoControl"
'Public Function EncontrarVerhoeff(numero As Integer, digitos As Integer)
'Dim Tmp As Integer
'    Tmp = numero
'    While digitos > 0
'        Tmp = calsum(Tmp)
'        digitos = digitos
'    Wend
'End Function
Private Function Modulo(num As Double, Divi As Double) As Long
Dim R As Long
   R = Int(num / Divi)
   Modulo = num - (R * Divi)
End Function

Public Function Verhoeff(Numero As String) As Integer
Dim i As Integer, j As Integer, k As Integer
Dim Cadena As String
Dim Mul(0 To 9, 0 To 9) As Currency
Dim Per(0 To 7, 0 To 9) As Currency
Dim Inv(0 To 9) As Variant
Dim NumInv() As String, Chek As Integer

   Cadena = "0,1,2,3,4,5,6,7,8,9;" & _
            "1,2,3,4,0,6,7,8,9,5;" & _
            "2,3,4,0,1,7,8,9,5,6;" & _
            "3,4,0,1,2,8,9,5,6,7;" & _
            "4,0,1,2,3,9,5,6,7,8;" & _
            "5,9,8,7,6,0,4,3,2,1;" & _
            "6,5,9,8,7,1,0,4,3,2;" & _
            "7,6,5,9,8,2,1,0,4,3;" & _
            "8,7,6,5,9,3,2,1,0,4;" & _
            "9,8,7,6,5,4,3,2,1,0;"
   
   k = 1
   For i = 0 To 9
      j = -1
      Do While Mid(Cadena, k, 1) <> ";"
         If Mid(Cadena, k, 1) <> "," Then
            j = j + 1
            Mul(i, j) = Mid(Cadena, k, 1)
         End If
         k = k + 1
      Loop
      k = k + 1
   Next i
   
   Cadena = ""
      
   Cadena = "0,1,2,3,4,5,6,7,8,9;" & _
            "1,5,7,6,2,8,3,0,9,4;" & _
            "5,8,0,3,7,9,6,1,4,2;" & _
            "8,9,1,6,0,4,3,5,2,7;" & _
            "9,4,5,3,1,2,6,8,7,0;" & _
            "4,2,8,6,5,7,3,9,0,1;" & _
            "2,7,9,3,8,0,6,4,1,5;" & _
            "7,0,4,6,9,1,3,2,5,8;"
            
   k = 1
   For i = 0 To 7
      j = -1
      Do While Mid(Cadena, k, 1) <> ";"
         If Mid(Cadena, k, 1) <> "," Then
            j = j + 1
            Per(i, j) = Mid(Cadena, k, 1)
         End If
         k = k + 1
      Loop
      k = k + 1
   Next i
   
   Cadena = ""
      
   Inv(0) = 0
   Inv(1) = 4
   Inv(2) = 3
   Inv(3) = 2
   Inv(4) = 1
   Inv(5) = 5
   Inv(6) = 6
   Inv(7) = 7
   Inv(8) = 8
   Inv(9) = 9
   
   Cadena = Trim(StrReverse(Numero))
   
   ReDim NumInv(0 To Len((Cadena)) - 1)
   
   For k = 0 To Len(Cadena) - 1
       NumInv(k) = Mid(Cadena, k + 1, 1)
   Next k
   
   Chek = 0
   
   For i = 0 To Len(Cadena) - 1
      Chek = Mul(Chek, Per(((i + 1) Mod 8), Val(NumInv(i))))
   Next i
   Verhoeff = Inv(Chek)
End Function

Public Function Base64(num As String) As String
Dim Diccionario() As Variant
Dim Cociente As Double, Resto As Integer
Dim Resultado As String
   Diccionario = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", _
                "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", _
                "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", _
                "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", _
                "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
                "y", "z", "+", "/")
   Cociente = 1
   Resultado = ""
   
   Do While Int(Cociente) > 0
      Cociente = Int(CDbl(num) / 64)
      Resto = CDbl(num) Mod 64
      Resultado = Diccionario(Resto) & Resultado
      num = Cociente
   Loop
   Base64 = Resultado
End Function

Private Sub SWAP(ByRef a As Integer, ByRef B As Integer)
Dim Aux As String
   Aux = a
   a = B
   B = Aux
End Sub

Public Function AllegedRC4(Cadena As String, Key As String) As String
Dim State(0 To 255) As Integer, Llave() As String, Mensaje() As String
Dim X As Integer, Y As Integer
Dim Index1 As Integer, Index2 As Integer
Dim Cod As Integer, Cifrado As String
   
   Index1 = 0
   Index2 = 0
   X = 0
   Y = 0
   Cifrado = ""
   For i = 0 To Len(Key) - 1
      ReDim Preserve Llave(0 To i)
      Llave(i) = Mid(Key, i + 1, 1)
   Next i
   
   For i = 0 To Len(Cadena) - 1
      ReDim Preserve Mensaje(0 To i)
      Mensaje(i) = Mid(Cadena, i + 1, 1)
   Next i
   
   For i = 0 To 255
      State(i) = i
   Next i
   
   For i = 0 To 255
      Index2 = (Asc(Llave(Index1)) + State(i) + Index2) Mod 256
      SWAP State(i), State(Index2) 'intercambiando valores
      Index1 = (Index1 + 1) Mod Len(Key)
   Next i
   
   For i = 0 To Len(Cadena) - 1
      X = (X + 1) Mod 256
      Y = (State(X) + Y) Mod 256
      SWAP State(X), State(Y)
      Cod = Asc(Mensaje(i)) Xor State((State(X) + State(Y)) Mod 256)
      Cifrado = Cifrado & String(2 - Len(Trim(Hex(Cod))), "0") & Hex(Cod)
   Next i
   
   AllegedRC4 = Trim(Mid(Cifrado, 1, Len(Cifrado)))
End Function



Attribute VB_Name = "modcnn"
'-------------------------
'updated 2000/08/17 by EMA
'-------------------------

Option Explicit
Public Const cnnString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=SAF2001;Data Source=sersis;"
Public db As ADODB.Connection
Public glusuario As String
Public GlFechaProceso As Date
Public glProceso As String
Public GlNombFor$
Sub main()
Set db = New ADODB.Connection
DE.Edson.ConnectionString = cnnString
DE.Historico.ConnectionString = cnnString
DE.Pagos.ConnectionString = cnnString
If InStr(UCase(cnnString), "SAFPRUEBA") > 0 Then
    MsgBox Chr(13) & Chr(13) & "A T E N C I O N" & Chr(13) & "estoy en la base de datos de prueba"
End If

db.CursorLocation = adUseClient
db.Open cnnString
glusuario = "XXX"
glProceso = "CONSULTORIA"
GlFechaProceso = Date
ac_main_c.Show vbModal
End
End Sub

Public Function fHoraValida(xHora As String) As Boolean
Dim h%, m%
h = Val(Mid(xHora, 1, 2))
m = Val(Mid(xHora, 4, 2))
If h >= 0 And h <= 24 Then
    If m >= 0 And m <= 60 Then
        fHoraValida = True
    Else
        fHoraValida = False
    End If
Else
    fHoraValida = False
End If
End Function

Public Function meses(nMes As Integer) As String
'funcion que devuelve el mes
Select Case nMes
Case 1
    meses = "enero"
Case 2
    meses = "febrero"
Case 3
    meses = "marzo"
Case 4
    meses = "abril"
Case 5
    meses = "mayo"
Case 6
    meses = "junio"
Case 7
    meses = "julio"
Case 8
    meses = "agosto"
Case 9
    meses = "septiembre"
Case 10
    meses = "octubre"
Case 11
    meses = "noviembre"
Case 12
    meses = "diciembre"
Case Else
    meses = "**no identificado**"
End Select
End Function


Public Function Literal(ByVal NroReal As Currency) As String
  ' Declaración de variables Locales
  Dim N        As Long        ' Número para división
  Dim NRef     As Long        ' Número de referencia
  Dim AntResto As Long        ' Resto anterior
  Dim Resto    As Long        ' Resto actual
  Dim C1       As Integer     ' Contador de 3 dígitos del Num.
  Dim c2       As Integer     ' Contador general de dígitos del Num.
  Dim NumLite1 As String      ' Cadena por cada 3 dígitos
  Dim NumLite2 As String      ' Cadena general
  Dim CAux     As String      ' Cadena auxiliar
  Dim Numero   As Long
  Dim Dec      As Integer
  Dim N1(20, 5) As String     'Para obtener el literal de un número
  
  Numero = Int(NroReal)
  ' N1 Matriz de Nombres de los números
        N1(0, 1) = "Un"
        N1(1, 1) = "Uno"
        N1(2, 1) = "Dos"
        N1(3, 1) = "Tres"
        N1(4, 1) = "Cuatro"
        N1(5, 1) = "Cinco"
        N1(6, 1) = "Seis"
        N1(7, 1) = "Siete"
        N1(8, 1) = "Ocho"
        N1(9, 1) = "Nueve"
        '
        N1(0, 2) = "Diez"
        N1(1, 2) = "Once"
        N1(2, 2) = "Doce"
        N1(3, 2) = "Trece"
        N1(4, 2) = "Catorce"
        N1(5, 2) = "Quince"
        N1(6, 2) = "Dieciseis"
        N1(7, 2) = "Diecisiete"
        N1(8, 2) = "Dieciocho"
        N1(9, 2) = "Diecinueve"
        '
        N1(0, 3) = "Veinte"
        N1(1, 3) = "Veintiuno"
        N1(2, 3) = "Veintidos"
        N1(3, 3) = "Veintitres"
        N1(4, 3) = "Veinticuatro"
        N1(5, 3) = "Veinticinco"
        N1(6, 3) = "Veintiseis"
        N1(7, 3) = "Veintisiete"
        N1(8, 3) = "Veintiocho"
        N1(9, 3) = "Veintinueve"
        '
        N1(0, 4) = ""
        N1(1, 4) = "Diez"
        N1(2, 4) = "Veinte"
        N1(3, 4) = "Treinta"
        N1(4, 4) = "Cuarenta"
        N1(5, 4) = "Cincuenta"
        N1(6, 4) = "Sesenta"
        N1(7, 4) = "Setenta"
        N1(8, 4) = "Ochenta"
        N1(9, 4) = "Noventa"
        '
        N1(0, 5) = ""
        N1(1, 5) = "Ciento"
        N1(2, 5) = "Doscientos"
        N1(3, 5) = "Trescientos"
        N1(4, 5) = "Cuatrocientos"
        N1(5, 5) = "Quinientos"
        N1(6, 5) = "Seiscientos"
        N1(7, 5) = "Setecientos"
        N1(8, 5) = "Ochocientos"
        N1(9, 5) = "Novecientos"
  ' Inicio de Variables
  CAux = ""
  NumLite1 = ""
  NumLite2 = ""
  C1 = 0
  c2 = 0
  Resto = 0
  AntResto = 0
  N = Numero
  NRef = Numero
  ' Realizar Mientras el Número sea > 0
  While N > 0
    Do
      AntResto = Resto
      N = Int(N / 10)
      Resto = Numero Mod 10
      '
      C1 = C1 + 1
      c2 = c2 + 1
      If (C1 = 1) And (Resto <> 0) Then
        ' Colocar Uno solo si es Mil o tiene 0 y 0 en los miles Ej: 111 001 000 o si son unidades
        If (Resto = 1 And NRef >= 1000 And NRef <= 1999) Or (Resto = 1 And NRef >= 1000 And (N Mod 10 = 0) And (Int(N / 10) Mod 10 = 0) And c2 = 4) Then
        Else
          ' Colocar Un en Lugar de Uno si es 1
          ' Si es Mil o Millón
          If (c2 = 4 And Resto = 1) Or (c2 = 7 And Resto = 1) Then
            NumLite1 = N1(0, 1) + NumLite1
          Else
            NumLite1 = N1(Resto, 1) + NumLite1
          End If
        End If
      End If
      If (C1 = 2) And (Resto <> 0) Then
        If AntResto = 0 Then
          NumLite1 = N1(Resto, 4)
        Else
          NumLite1 = N1(Resto, 4) + " y " + NumLite1
        End If
      End If
      If (C1 = 2) And (Resto = 1) Then
        NumLite1 = N1(AntResto, 2)
      End If
      If (C1 = 2) And (Resto = 2) Then
        If (c2 > 4) And (AntResto = 1) Then
          NumLite1 = "Veintiun"
        Else
          NumLite1 = N1(AntResto, 3)
        End If
      End If
      If (C1 = 3) And (Resto <> 0) Then
        If Resto = 1 And NumLite1 = "" Then
          NumLite1 = "Cien"
        Else
          NumLite1 = N1(Resto, 5) + " " + NumLite1
        End If
      End If
      '
      CAux = Str(Resto) + CAux
      Numero = N
    Loop Until C1 = 3 Or N = 0
    C1 = 0
    ' Agregar Mil
    If c2 = 4 Or c2 = 5 Or c2 = 6 Then
       If Val(CAux) >= 1000 Then
         NumLite1 = NumLite1 + " Mil "
       End If
    End If
    ' Agregar Millón
    If c2 = 7 Or c2 = 8 Or c2 = 9 Then
      If NRef >= 1000000 And NRef <= 1999999 Then
        NumLite1 = NumLite1 + " Millón "
      Else
        NumLite1 = NumLite1 + " Millones "
      End If
    End If
    '
    NumLite2 = NumLite1 + NumLite2
    NumLite1 = ""
  Wend
  
  'Para la parte decimal del monto
  Dec = (NroReal - Int(NroReal)) * 100
  If Dec = 0 Then
     Literal = NumLite2 & " 00/100"
     'Literal = NumLite2
  Else
     If Dec >= 1 And Dec <= 9 Then
        Literal = NumLite2 & " 0" & Dec & "/100"
        'Literal = NumLite2
     Else
        Literal = NumLite2 & " " & Dec & "/100"
        'Literal = NumLite2
     End If
  End If
End Function


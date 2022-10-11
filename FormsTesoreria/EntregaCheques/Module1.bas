Attribute VB_Name = "Module1"
Public db As Connection
Public rsRegularizacion As New Recordset
Public rsDetalle As New Recordset
Public LiteralCry As String
Public NrosChequeImprimir As String

Public NombreUsuario As String
Public Cont_Comp As Long

Public recSetAuxActualizar1 As New ADODB.Recordset
Public recsetAdicion As New ADODB.Recordset
Public recSetAuxActualizar As New ADODB.Recordset
Public recSetPartida As New ADODB.Recordset
Public recSetGenera As New ADODB.Recordset

'GROVER
Public recSetBusqueda As New ADODB.Recordset
Public recSetAuxcomp As New ADODB.Recordset
Public recSetAuxcomp1 As New ADODB.Recordset
Public recSetPartida1 As New ADODB.Recordset

Public recSetAuxbenefi1 As New ADODB.Recordset
Public recSetPartid1 As New ADODB.Recordset

'Freddy
Public GlUsuario As String

'PARA FINES DE IMPRESION
Public swMes As Integer
Public swFecha As Integer
Sub Main()
   Set db = New Connection
   db.CursorLocation = adUseClient
   'db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=c:\fbase\pragma.mdb;"
   'db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=\\sersis\saf\Celia-No Borrar\udapre.mdb;"
   db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=\\sersis\saf\udapre.mdb;"
   ''''''db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=\\sersis\saf\udapre08062000.mdb;"
   'db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=D:\Saf-1\udapre.mdb;"
   'db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=d:\18Abril2000\udapre.mdb;"
   'db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=d:\saf-1\udapre.mdb;"
   'db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=c:\desa\db\udapre.mdb;"
   ''''''db.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=SAF2000;Data Source=sersis"
   'FrmComprobanteTrans.Show
   'FrmCP.Show
   'frmMain.Show
   '''''frmLogin.Show
   'FrmChequesCuenta.Show
   'FrmImprimirComprobante.Show
   'FrmActivacionCheques.Show
   'FrmActivacionCheques.Show
   'FrmTransferencia.Show
   'FrmActivacionCheques.Show
   'FrmDevoluciones.Show
   'FrmColaImpresion.Show
   'FrmOperacionCheques.Show
   'FrmCuentaBancaria.Show
   'FrmPagosRealizados.Show
   'FrmPagosTotal.Show
   ' Form1.Show
   'FrmActivacionCheques.Show
   MDIForm1.Show
   'FrmTributosFiscales.Show
   ' FrmPagadero.Show
End Sub

'LITERAL DE CELIA TARQUINO
Public Function Literal(CADENA As String) As String
Dim sw As Integer
Dim sw1 As Integer
Dim swc As Integer
Dim VEC(20) As Long
sw = 0
      '*********PARTE DECIMAL*********
            CADENA = Round(CADENA, 2)
             x = Len(CADENA)
              For K = 1 To x
                  z = Mid(CADENA, K, 1)
                  If (z = ".") Or sw = 1 Then
                    D = D + Mid(CADENA, K, 1)
                    sw = 1
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
 If CADENA <> "" Then
    CADENA = Int(CADENA)
 Else
    MsgBox "No existe monto"
 End If
   S = ""
   z = ""
   C = 0
   K = 0
   sw1 = 0
   swc = 0
   
   
   x = Len(CADENA)
   For i = 1 To x
       A = Mid(CADENA, i, 1)
       VEC(i) = Mid(CADENA, i, 1)
   Next i
j = x
While j <> 0
K = K + 1
If K <> 8 Then
  If C <> 3 Then
       C = C + 1
      
       If C = 1 And (VEC(j - 1) <> 1 And VEC(j - 1) <> 2) Then
            Select Case VEC(j)
                Case 0: S = " " + S
                Case 1:
                   If sw1 <> 1 Then
                      S = "UNO " + z + S
                   End If
                   If sw1 = 1 Then
                      S = "UN " + z + S
                   End If
                   
                Case 2: S = "DOS " + z + S
                Case 3: S = "TRES " + z + S
                Case 4: S = "CUATRO " + z + S
                Case 5: S = "CINCO " + z + S
                Case 6: S = "SEIS " + z + S
                Case 7: S = "SIETE " + z + S
                Case 8: S = "OCHO " + z + S
                Case 9: S = "NUEVE " + z + S
          End Select
          
           'If J + 1 <> "" And sw1 <> 1 And VEC(J - 1) <> 0 And VEC(J) <> 0 Then
           If VEC(j - 1) <> 0 And VEC(j) <> 0 Then
                 S = "Y " + S
           End If
           
        End If
        
         If C = 2 And VEC(j) = 1 Then
               swc = 1
                Select Case VEC(j + 1)
                      Case 0: S = "DIEZ " + z + S
                      Case 1: S = "ONCE " + z + S
                      Case 2: S = "DOCE " + z + S
                      Case 3: S = "TRECE " + z + S
                      Case 4: S = "CATORCE " + z + S
                      Case 5: S = "QUINCE " + z + S
                      Case 6: S = "DIECISEIS " + z + S
                      Case 7: S = "DIECISIETE " + z + S
                      Case 8: S = "DIECIOCHO " + z + S
                      Case 9: S = "DIECINUEVE " + z + S
                End Select
          End If
          
          If C = 2 And VEC(j) = 2 Then
                Select Case VEC(j + 1)
                      Case 0: S = "VEINTE " + z + S
                      Case 1: S = "VEINTIUNO " + z + S
                      Case 2: S = "VEINTIDOS " + z + S
                      Case 3: S = "VEINTITRES " + z + S
                      Case 4: S = "VEINTICUATRO " + z + S
                      Case 5: S = "VEINTICINCO " + z + S
                      Case 6: S = "VEINTISEIS " + z + S
                      Case 7: S = "VEINTISIETE " + z + S
                      Case 8: S = "VEINTIOCHO " + z + S
                      Case 9: S = "VEINTINUEVE " + z + S
                End Select
          End If
   
        If C = 2 Then
            Select Case VEC(j)
                Case 3: S = "TREINTA " + z + S
                Case 4: S = "CUARENTA " + z + S
                Case 5: S = "CINCUENTA " + z + S
                Case 6: S = "SESENTA " + z + S
                Case 7: S = "SETENTA " + z + S
                Case 8: S = "OCHENTA " + z + S
                Case 9: S = "NOVENTA " + z + S
            End Select
            
        End If
        
        If C = 3 Then
            Select Case VEC(j)
                Case 1:
                If j = 1 Then
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       S = "CIEN " + z + S
                    Else
                       S = "CIENTO " + z + S
                    End If
                Else
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       S = "CIEN " + z + S
                    Else
                       S = "CIENTO " + z + S
                    End If
                       'S = "CIENTO " + z + S
                End If
                Case 2: S = "DOSCIENTOS " + z + S
                Case 3: S = "TRESCIENTOS " + z + S
                Case 4: S = "CUATROCIENTOS " + z + S
                Case 5: S = "QUINIENTOS " + z + S
                Case 6: S = "SEISCIENTOS " + z + S
                Case 7: S = "SETECIENTOS " + z + S
                Case 8: S = "OCHOCIENTOS " + z + S
                Case 9: S = "NOVECIENTOS " + z + S
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
        C = 0
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
      C = 0
      sw1 = 1
 End If
   j = j - 1
   
Wend

Literal = S + D
End Function

'LITERAL DE ERICK MUÑOZ
'Public Function Literal(ByVal NroReal As Currency) As String
'  ' Declaración de variables Locales
'  Dim N        As Long        ' Número para división
'  Dim NRef     As Long        ' Número de referencia
'  Dim AntResto As Long        ' Resto anterior
'  Dim Resto    As Long        ' Resto actual
'  Dim C1       As Integer     ' Contador de 3 dígitos del Num.
'  Dim c2       As Integer     ' Contador general de dígitos del Num.
'  Dim NumLite1 As String      ' Cadena por cada 3 dígitos
'  Dim NumLite2 As String      ' Cadena general
'  Dim CAux     As String      ' Cadena auxiliar
'  Dim Numero   As Long
'  Dim dec      As Integer
'  Dim N1(20, 5) As String     'Para obtener el literal de un número
'
'  Numero = Int(NroReal)
'  ' N1 Matriz de Nombres de los números
'        N1(0, 1) = "Un"
'        N1(1, 1) = "Uno"
'        N1(2, 1) = "Dos"
'        N1(3, 1) = "Tres"
'        N1(4, 1) = "Cuatro"
'        N1(5, 1) = "Cinco"
'        N1(6, 1) = "Seis"
'        N1(7, 1) = "Siete"
'        N1(8, 1) = "Ocho"
'        N1(9, 1) = "Nueve"
'        '
'        N1(0, 2) = "Diez"
'        N1(1, 2) = "Once"
'        N1(2, 2) = "Doce"
'        N1(3, 2) = "Trece"
'        N1(4, 2) = "Catorce"
'        N1(5, 2) = "Quince"
'        N1(6, 2) = "Dieciseis"
'        N1(7, 2) = "Diecisiete"
'        N1(8, 2) = "Dieciocho"
'        N1(9, 2) = "Diecinueve"
'        '
'        N1(0, 3) = "Veinte"
'        N1(1, 3) = "Veintiuno"
'        N1(2, 3) = "Veintidos"
'        N1(3, 3) = "Veintitres"
'        N1(4, 3) = "Veinticuatro"
'        N1(5, 3) = "Veinticinco"
'        N1(6, 3) = "Veintiseis"
'        N1(7, 3) = "Veintisiete"
'        N1(8, 3) = "Veintiocho"
'        N1(9, 3) = "Veintinueve"
'        '
'        N1(0, 4) = ""
'        N1(1, 4) = "Diez"
'        N1(2, 4) = "Veinte"
'        N1(3, 4) = "Treinta"
'        N1(4, 4) = "Cuarenta"
'        N1(5, 4) = "Cincuenta"
'        N1(6, 4) = "Sesenta"
'        N1(7, 4) = "Setenta"
'        N1(8, 4) = "Ochenta"
'        N1(9, 4) = "Noventa"
'        '
'        N1(0, 5) = ""
'        N1(1, 5) = "Ciento"
'        N1(2, 5) = "Doscientos"
'        N1(3, 5) = "Trescientos"
'        N1(4, 5) = "Cuatrocientos"
'        N1(5, 5) = "Quinientos"
'        N1(6, 5) = "Seiscientos"
'        N1(7, 5) = "Setecientos"
'        N1(8, 5) = "Ochocientos"
'        N1(9, 5) = "Novecientos"
'  ' Inicio de Variables
'  CAux = ""
'  NumLite1 = ""
'  NumLite2 = ""
'  C1 = 0
'  c2 = 0
'  Resto = 0
'  AntResto = 0
'  N = Numero
'  NRef = Numero
'  ' Realizar Mientras el Número sea > 0
'  While N > 0
'    Do
'      AntResto = Resto
'      N = Int(N / 10)
'      Resto = Numero Mod 10
'      '
'      C1 = C1 + 1
'      c2 = c2 + 1
'      If (C1 = 1) And (Resto <> 0) Then
'        ' Colocar Uno solo si es Mil o tiene 0 y 0 en los miles Ej: 111 001 000 o si son unidades
'        If (Resto = 1 And NRef >= 1000 And NRef <= 1999) Or (Resto = 1 And NRef >= 1000 And (N Mod 10 = 0) And (Int(N / 10) Mod 10 = 0) And c2 = 4) Then
'             If c2 = 4 Or c2 = 7 Or c2 = 11 Then
'               NumLite1 = N1(0, 1) + NumLite1
''             Else
''               NumLite1 = N1(1, 1) + NumLite1
'             End If
'        Else
'          ' Colocar Un en Lugar de Uno si es 1
'          ' Si es Mil o Millón
'          If (c2 = 4 And Resto = 1) Or (c2 = 7 And Resto = 1) Then
'            NumLite1 = N1(0, 1) + NumLite1
'          Else
'            NumLite1 = N1(Resto, 1) + NumLite1
'          End If
'        End If
'      End If
'      If (C1 = 2) And (Resto <> 0) Then
'        If AntResto = 0 Then
'          NumLite1 = N1(Resto, 4)
'        Else
'          NumLite1 = N1(Resto, 4) + " y " + NumLite1
'        End If
'      End If
'      If (C1 = 2) And (Resto = 1) Then
'        NumLite1 = N1(AntResto, 2)
'      End If
'      If (C1 = 2) And (Resto = 2) Then
'        If (c2 > 4) And (AntResto = 1) Then
'          NumLite1 = "Veintiun"
'        Else
'          NumLite1 = N1(AntResto, 3)
'        End If
'      End If
'      If (C1 = 3) And (Resto <> 0) Then
'        If Resto = 1 And NumLite1 = "" Then
'          NumLite1 = "Cien"
'        Else
'          NumLite1 = N1(Resto, 5) + " " + NumLite1
'        End If
'      End If
'      '
'      CAux = Str(Resto) + CAux
'      Numero = N
'    Loop Until C1 = 3 Or N = 0
'    C1 = 0
'    ' Agregar Mil
'    If c2 = 4 Or c2 = 5 Or c2 = 6 Then
'       If Val(CAux) >= 1000 Then
'         If NumLite1 = "" Then
'            NumLite1 = NumLite1 + "Un Mil "
'         Else
'            NumLite1 = NumLite1 + " Mil "
'         End If
'       End If
'    End If
'    ' Agregar Millón
'    If c2 = 7 Or c2 = 8 Or c2 = 9 Then
'      If NRef >= 1000000 And NRef <= 1999999 Then
'        NumLite1 = NumLite1 + " Millón "
'      Else
'        NumLite1 = NumLite1 + " Millones "
'      End If
'    End If
'    '
'    NumLite2 = NumLite1 + NumLite2
'    NumLite1 = ""
'  Wend
'  Literal = NumLite2
'
'
'
'  'Para la parte decimal del monto
'  dec = (NroReal - Int(NroReal)) * 100
'  If dec = 0 Then
'     Literal = NumLite2 & " 00/100"
'  Else
'     If dec >= 1 And dec <= 9 Then
'        Literal = NumLite2 & " 0" & dec & "/100"
'     Else
'        Literal = NumLite2 & " " & dec & "/100"
'     End If
'  End If
'End Function


Public Function Buscar(atrib1 As String, atrib2 As String, atrib3 As String, atrib4 As String, atrib5 As String, atrib6 As String) As Boolean
    Set recSetBusqueda = New ADODB.Recordset
    recSetBusqueda.CursorLocation = adUseClient
    If recSetBusqueda.State = 1 Then recSetBusqueda.Close
    recSetBusqueda.Open atrib1 & _
    " where   Cod_Trans='" & atrib2 & "' and Org_Codigo='" & atrib3 & "' " & _
    " and Ges_Gestion='" & atrib4 & "' and Tipo_comp='" & atrib5 & "' and Cod_Trans_Detalle='" & atrib6 & "'", db, adOpenDynamic, adLockOptimistic, adCmdText
    If recSetBusqueda.RecordCount > 0 Then
    Buscar = True
    Else
    Buscar = False
    End If

End Function



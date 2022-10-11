VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClBuscaSecuencialEnRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Permite realizar búsquedas de Tipo Secuenciales en un Recordset. La Búsqueda puede ser hasta en 3 campos."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'variables locales para almacenar los valores de las propiedades
Private mvarRecordset As Object 'copia local
Private mvarCampo1 As String 'copia local
Private mvarValorCampo1 As String 'copia local
Private mvarCampo2 As String 'copia local
Private mvarValorCampo2 As String 'copia local
Private mvarCampo3 As String 'copia local
Private mvarValorCampo3 As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private CuantosCampos As Byte 'copia local


Public Function Ejecutar() As Boolean
Attribute Ejecutar.VB_Description = "Ejecuta la busqueda, devolviendo True si encontro o False de otro modo."
  CuantosCampos = 0
  If mvarCampo1 <> "" Then CuantosCampos = 1
  If mvarCampo2 <> "" Then CuantosCampos = 2
  If mvarCampo3 <> "" Then CuantosCampos = 3
  Ejecutar = BuscaSecuencial(mvarRecordset, mvarCampo1, mvarValorCampo1, mvarCampo2, mvarValorCampo2, mvarCampo3, mvarValorCampo3, CuantosCampos)
End Function


Public Property Let ValorCampo1(ByVal vData As String)
Attribute ValorCampo1.VB_Description = "Valor 1 buscado en el campo 1."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ValorCampo1 = 5
    mvarValorCampo1 = vData
End Property

Public Property Get ValorCampo1() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ValorCampo1
    ValorCampo1 = mvarValorCampo1
End Property


Public Property Let Campo1(ByVal vData As String)
Attribute Campo1.VB_Description = "Campo 1 por el que se busca."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Campo1 = 5
    mvarCampo1 = vData
End Property

Public Property Get Campo1() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Campo1
    Campo1 = mvarCampo1
End Property


Public Property Let ValorCampo2(ByVal vData As String)
Attribute ValorCampo2.VB_Description = "Valor 2 buscado en el campo 2."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ValorCampo1 = 5
    mvarValorCampo2 = vData
End Property

Public Property Get ValorCampo2() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ValorCampo1
    ValorCampo2 = mvarValorCampo2
End Property


Public Property Let Campo2(ByVal vData As String)
Attribute Campo2.VB_Description = "Campo 2 por el que se busca."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Campo1 = 5
    mvarCampo2 = vData
End Property

Public Property Get Campo2() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Campo1
    Campo2 = mvarCampo2
End Property


Public Property Let ValorCampo3(ByVal vData As String)
Attribute ValorCampo3.VB_Description = "Valor 3 buscado en el campo 3."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ValorCampo1 = 5
    mvarValorCampo3 = vData
End Property

Public Property Get ValorCampo3() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ValorCampo1
    ValorCampo3 = mvarValorCampo3
End Property


Public Property Let Campo3(ByVal vData As String)
Attribute Campo3.VB_Description = "Campo 3 por el que se busca."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Campo1 = 5
    mvarCampo3 = vData
End Property

Public Property Get Campo3() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Campo1
    Campo3 = mvarCampo3
End Property


Public Property Set Recordset(ByVal vData As Object)
Attribute Recordset.VB_Description = "Recordset sobre el que se realizará la búsqueda."
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Recordset = Form1
    Set mvarRecordset = vData
End Property


Public Property Get Recordset() As Object
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Recordset
    Set Recordset = mvarRecordset
End Property


Private Function BuscaSecuencial(rsBusca As ADODB.Recordset, _
                                Campo1 As String, valor1 As String, _
                                Campo2 As String, valor2 As String, _
                                Campo3 As String, valor3 As String, _
                                CuantosCampos As Byte) As Boolean
'Cuantos   : Por Cuantos Busca, 2 o 3
On Error GoTo QError
Dim Encontro As Boolean
Dim Marca As Variant
  Screen.MousePointer = vbHourglass
  Encontro = False
  If rsBusca.RecordCount = 0 Then
    BuscaSecuencial = False
    Exit Function
  End If
  Marca = rsBusca.Bookmark
  rsBusca.MoveFirst
  While (Not rsBusca.EOF) And (Not Encontro)
    Select Case CuantosCampos
      Case 1
        If rsBusca(Campo1).Type = adDBTimeStamp Then
          If CDate(rsBusca(Campo1)) = CDate(valor1) Then Encontro = True Else rsBusca.MoveNext
        Else
          If rsBusca(Campo1) = valor1 Then Encontro = True Else rsBusca.MoveNext
        End If
      Case 2
        If rsBusca(Campo1).Type = adDBTimeStamp Then
          If CDate(rsBusca(Campo1)) = CDate(valor1) Then
            If rsBusca(Campo2).Type = adDBTimeStamp Then
              If CDate(rsBusca(Campo2)) = CDate(valor2) Then Encontro = True Else rsBusca.MoveNext
            Else
               If (rsBusca(Campo2) = valor2) Then Encontro = True Else rsBusca.MoveNext
            End If
          Else
             rsBusca.MoveNext
          End If
        Else
          If rsBusca(Campo1) = valor1 Then
            If rsBusca(Campo2).Type = adDBTimeStamp Then
              If CDate(rsBusca(Campo2)) = CDate(valor2) Then Encontro = True Else rsBusca.MoveNext
            Else
               If (rsBusca(Campo2) = valor2) Then Encontro = True Else rsBusca.MoveNext
            End If
          Else
             rsBusca.MoveNext
          End If
        End If
      Case 3
        If rsBusca(Campo1).Type = adDBTimeStamp Then
          If CDate(rsBusca(Campo1)) = CDate(valor1) Then
            If rsBusca(Campo2).Type = adDBTimeStamp Then
              If CDate(rsBusca(Campo2)) = CDate(valor2) Then
                If rsBusca(Campo3).Type = adDBTimeStamp Then
                  If CDate(rsBusca(Campo3)) = CDate(valor3) Then Encontro = True Else rsBusca.MoveNext
                Else
                  If (rsBusca(Campo3) = valor3) Then Encontro = True Else rsBusca.MoveNext
                End If
              Else
                rsBusca.MoveNext
              End If
            Else
              If (rsBusca(Campo2) = valor2) Then
                If rsBusca(Campo3).Type = adDBTimeStamp Then
                  If CDate(rsBusca(Campo3)) = CDate(valor3) Then Encontro = True Else rsBusca.MoveNext
                Else
                  If (rsBusca(Campo3) = valor3) Then Encontro = True Else rsBusca.MoveNext
                End If
              End If
            End If
          Else
             rsBusca.MoveNext
          End If
        Else
          If rsBusca(Campo1) = valor1 Then
            If rsBusca(Campo2).Type = adDBTimeStamp Then
              If CDate(rsBusca(Campo2)) = CDate(valor2) Then
                If rsBusca(Campo3).Type = adDBTimeStamp Then
                  If CDate(rsBusca(Campo3)) = CDate(valor3) Then Encontro = True Else rsBusca.MoveNext
                Else
                  If (rsBusca(Campo3) = valor3) Then Encontro = True Else rsBusca.MoveNext
                End If
              Else
                rsBusca.MoveNext
              End If
            Else
              If (rsBusca(Campo2) = valor2) Then
                If rsBusca(Campo3).Type = adDBTimeStamp Then
                  If CDate(rsBusca(Campo3)) = CDate(valor3) Then Encontro = True Else rsBusca.MoveNext
                Else
                  If (rsBusca(Campo3) = valor3) Then Encontro = True Else rsBusca.MoveNext
                End If
              Else
                rsBusca.MoveNext
              End If
            End If
          Else
             rsBusca.MoveNext
          End If
        End If
    End Select
  Wend
  If Not Encontro Then rsBusca.Bookmark = Marca
  BuscaSecuencial = Encontro
  Screen.MousePointer = vbDefault
  Exit Function
QError:
  Screen.MousePointer = vbDefault
  BuscaSecuencial = False
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Atención"
End Function


Private Sub Class_Initialize()
  mvarCampo1 = ""
  mvarValorCampo1 = ""
  mvarCampo2 = ""
  mvarValorCampo2 = ""
  mvarCampo3 = ""
  mvarValorCampo3 = ""
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClBuscaBinariaEnRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Permite realizar búsquedas de Tipo binarias en un Recordset. La Búsqueda solo es sobre un campo de tipo único."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'variables locales para almacenar los valores de las propiedades
Private mvarRecordset As Object 'copia local
Private mvarCampo As String 'copia local
Private mvarValorCampo As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarPosicionar As Boolean 'copia local

Public Function Ejecutar() As Long
Attribute Ejecutar.VB_Description = "Ejecutar la búsqueda."
Dim Resp As Long
  Resp = BuscaBinaria(mvarRecordset, mvarCampo, mvarValorCampo)
  If Resp <> 0 Then
    If mvarPosicionar Then
      mvarRecordset.MoveFirst
      mvarRecordset.Move Resp - 1
    End If
  End If
  Ejecutar = Resp
End Function

Public Property Let Posicionar(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Posicionar = 5
    mvarPosicionar = vData
End Property

Public Property Get Posicionar() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Posicionar
    Posicionar = mvarPosicionar
End Property

Public Property Let ValorCampo(ByVal vData As String)
Attribute ValorCampo.VB_Description = "Valor que se busca en el campo especificado."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ValorCampo1 = 5
    mvarValorCampo = vData
End Property

Public Property Get ValorCampo() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ValorCampo1
    ValorCampo = mvarValorCampo
End Property


Public Property Let Campo(ByVal vData As String)
Attribute Campo.VB_Description = "Campo por el que se quiere buscar dentro del recordset."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Campo1 = 5
    mvarCampo = vData
End Property

Public Property Get Campo() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Campo1
    Campo = mvarCampo
End Property


Public Property Set Recordset(ByVal vData As Object)
Attribute Recordset.VB_Description = "Recordset sobre el que se quiere realizar la búsqueda."
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Recordset = Form1
    Set mvarRecordset = vData
End Property


Public Property Get Recordset() As Object
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Recordset
    Set Recordset = mvarRecordset
End Property

Private Function BuscaBinaria(rsBusca As ADODB.Recordset, Campo As String, valor As String) As Long
On Error GoTo QError
Dim min As Long
Dim max As Long
Dim Middle As Long
Dim AntMiddle As Long
Dim Sentido As Integer
Dim rsClone As ADODB.Recordset
  Screen.MousePointer = vbHourglass
  If rsBusca.CursorLocation <> adUseClient Then
    Screen.MousePointer = vbDefault
    MsgBox "El Recordset debe ser un cursor del lado del cliente...", vbInformation + vbOKOnly, "Atención"
    BuscaBinaria = 0
    Exit Function
  End If
  rsBusca.Sort = Campo & " ASC"
  Set rsClone = New ADODB.Recordset
  Set rsClone = rsBusca.Clone
  rsClone.Sort = Campo & " ASC"
    ' During the search the target's index will be
    ' between Min and Max: Min <= target index <= Max
    min = 1
    AntMiddle = 0
    rsClone.MoveFirst
    max = rsClone.RecordCount   'NumItems
    Sentido = 1
    Do While min <= max
        Middle = (max + min) / 2
        If AntMiddle < Middle Then
          If AntMiddle = 0 Then
            rsClone.Move Sentido * (Middle - AntMiddle - 1)
          Else
            rsClone.Move Sentido * (Middle - AntMiddle)
          End If
        Else
          rsClone.Move Sentido * (AntMiddle - Middle)
        End If
        If valor = Trim(rsClone(Campo)) Then     ' We have found it!
            BuscaBinaria = Middle
            Screen.MousePointer = vbDefault
            Exit Function
        ElseIf valor < rsClone(Campo) Then ' Search the left half.
            AntMiddle = Middle
            max = Middle - 1
            Sentido = -1
        Else
            ' Search the right half.
            AntMiddle = Middle
            min = Middle + 1
            Sentido = 1
        End If
    Loop
    ' If we get here the target is not in the list.
    BuscaBinaria = 0
    Screen.MousePointer = vbDefault
    Exit Function
QError:
  Screen.MousePointer = vbDefault
  BuscaBinaria = 0
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Atención"
End Function

Private Sub Class_Initialize()
  mvarCampo = ""
  mvarValorCampo = ""
  mvarPosicionar = False
End Sub

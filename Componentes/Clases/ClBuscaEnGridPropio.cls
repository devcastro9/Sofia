VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClBuscaEnGridPropio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Habilita herramientas de búsqueda y filtrado, en una ventana que contiene un Grid Propio."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarElegidoCol1 As String 'copia local
Private mvarElegidoCol2 As String 'copia local
Private mvarElegidoCol3 As String 'copia local
Private mvarElegidoCol4 As String 'copia local
Private mvarQueryUtilizado As String 'copia local
Private mvarCampoDefecto As String 'copia local
Private mvarOcultarPrimero As Boolean 'copia local
Private mvarTamañoCampos As String 'copia local
Private mvarTítulo As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarConexión As Object 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarFiltrosMultiples As Boolean 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarAliasColumnas As String 'copia local


Public Sub Ejecutar()
  GrFrmBuscaEnGridPropio.Elige mvarConexión, mvarQueryUtilizado, mvarTítulo, _
                               mvarCampoDefecto, mvarOcultarPrimero, mvarTamañoCampos, _
                               mvarFiltrosMultiples, mvarAliasColumnas
  mvarElegidoCol1 = GrFrmBuscaEnGridPropio.CodBuscado
  mvarElegidoCol2 = GrFrmBuscaEnGridPropio.CodBuscado1
  mvarElegidoCol3 = GrFrmBuscaEnGridPropio.CodBuscado2
  mvarElegidoCol4 = GrFrmBuscaEnGridPropio.CodBuscado3
End Sub


Public Property Let AliasColumnas(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.AliasColumnas = 5
    mvarAliasColumnas = vData
End Property

Public Property Get AliasColumnas() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.AliasColumnas
    AliasColumnas = mvarAliasColumnas
End Property


Public Property Let FiltrosMultiples(ByVal vData As Boolean)
Attribute FiltrosMultiples.VB_Description = "Permite realizar Filtros Multiples o simples. Es decir, si filtra siempre sobre el conjunto de registros inicial o sobre el subconjunto originado por el último filtro (defecto: True)."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FiltrosMultiples = 5
    mvarFiltrosMultiples = vData
End Property

Public Property Get FiltrosMultiples() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FiltrosMultiples
    FiltrosMultiples = mvarFiltrosMultiples
End Property


Public Property Set Conexión(ByVal vData As Object)
Attribute Conexión.VB_Description = "La Conexión que va a usar para crear el recordset."
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Conexión = Form1
    Set mvarConexión = vData
End Property

Public Property Get Conexión() As Object
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Conexión
    Set Conexión = mvarConexión
End Property


Public Property Let Título(ByVal vData As String)
Attribute Título.VB_Description = "Título de la ventana en la que se permitirá el trabajo."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Título = 5
    mvarTítulo = vData
End Property

Public Property Get Título() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Título
    Título = mvarTítulo
End Property


Public Property Let TamañoCampos(ByVal vData As String)
Attribute TamañoCampos.VB_Description = "Cadena que define el ancho de cada campo en el grid (Ejem: ""1-2-3-0.5-1"", donde cada número entre guiones, equivale al ancho de un campo en relación al ancho por defecto)."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.TamañoCampos = 5
    mvarTamañoCampos = vData
End Property

Public Property Get TamañoCampos() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.TamañoCampos
    TamañoCampos = mvarTamañoCampos
End Property


Public Property Let OcultarPrimero(ByVal vData As Boolean)
Attribute OcultarPrimero.VB_Description = "Bandera que indica si debe ocultar el primer campo de el query (util cuando no se quiere mostrar el código)."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.OcultarPrimero = 5
    mvarOcultarPrimero = vData
End Property

Public Property Get OcultarPrimero() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.OcultarPrimero
    OcultarPrimero = mvarOcultarPrimero
End Property


Public Property Let CampoDefecto(ByVal vData As String)
Attribute CampoDefecto.VB_Description = "Nombre deel campo por defecto que aparecerá en la lista de búsqueda."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CampoDefecto = 5
    mvarCampoDefecto = vData
End Property

Public Property Get CampoDefecto() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CampoDefecto
    CampoDefecto = mvarCampoDefecto
End Property


Public Property Let QueryUtilizado(ByVal vData As String)
Attribute QueryUtilizado.VB_Description = "Es el Query con el que se va a trabajar."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QueryUtilizado = 5
    mvarQueryUtilizado = vData
End Property

Public Property Get QueryUtilizado() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QueryUtilizado
    QueryUtilizado = mvarQueryUtilizado
End Property


Public Property Get ElegidoCol4() As String
Attribute ElegidoCol4.VB_Description = "Devuelve el valor de la cuarta columna del registro elegido, o vacio si no se eligio ninguno."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ElegidoCol4
    ElegidoCol4 = mvarElegidoCol4
End Property


Public Property Get ElegidoCol3() As String
Attribute ElegidoCol3.VB_Description = "Devuelve el valor de la tercera columna del registro elegido, o vacio si no se eligio ninguno."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ElegidoCol3
    ElegidoCol3 = mvarElegidoCol3
End Property


Public Property Get ElegidoCol2() As String
Attribute ElegidoCol2.VB_Description = "Devuelve el valor de la segunda columna del registro elegido, o vacio si no se eligio ninguno."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ElegidoCol2
    ElegidoCol2 = mvarElegidoCol2
End Property


Public Property Get ElegidoCol1() As String
Attribute ElegidoCol1.VB_Description = "Devuelve el valor de la primera columna del registro elegido, o vacio si no se eligio ninguno."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ElegidoCol1
    ElegidoCol1 = mvarElegidoCol1
End Property

Private Sub Class_Initialize()
  mvarElegidoCol1 = ""
  mvarElegidoCol2 = ""
  mvarElegidoCol3 = ""
  mvarElegidoCol4 = ""
  mvarQueryUtilizado = ""
  mvarCampoDefecto = ""
  mvarOcultarPrimero = False
  mvarTamañoCampos = ""
  mvarTítulo = "Elija su Opción"
  mvarFiltrosMultiples = True
  mvarAliasColumnas = ""
End Sub

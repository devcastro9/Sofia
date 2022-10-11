VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClBuscaEnGridExterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Habilita herramientas de búsqueda y filtrado, en una ventana que utiliza de base de búsqueda un Grid de un formulario primario que lo invoca."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarQueryUtilizado As String 'copia local
Private mvarCampoDefecto As String 'copia local
Private mvarTítulo As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarConexión As Object 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarGridTrabajo As Object 'copia local
Private mvarRecordsetTrabajo As Object 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarEsTdbGrid As Boolean 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarCamposVisibles As String 'copia local

Private TipoRs As Integer
Private TipoBloqueo As Integer
'Para activar este evento, use RaiseEvent con la siguiente sintaxis:
'RaiseEvent RefrescarGrid[(arg1, arg2, ... , argn)]
Public Event RefrescarGrid()
Attribute RefrescarGrid.VB_Description = "Evento en el que se debe ubicar código para refrescar el Grid con el código."

Private Sub EventoRefrescar()
  RaiseEvent RefrescarGrid
End Sub

Public Sub Ejecutar()
On Error GoTo QueError
  GrFrmBuscaEnGridExterno.GrPrincipal mvarConexión, mvarRecordsetTrabajo, _
                                      TipoRs, TipoBloqueo, _
                                      mvarQueryUtilizado, _
                                      mvarGridTrabajo, mvarEsTdbGrid, _
                                      mvarTítulo, mvarCamposVisibles, mvarCampoDefecto
  Exit Sub
QueError:
  MsgBox vbObjectError + 100
End Sub

Public Property Let CamposVisibles(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CamposVisibles = 5
    mvarCamposVisibles = vData
End Property

Public Property Get CamposVisibles() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CamposVisibles
    CamposVisibles = mvarCamposVisibles
End Property

Public Property Let EsTdbGrid(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.EsTdbGrid = 5
    mvarEsTdbGrid = vData
End Property

Public Property Get EsTdbGrid() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.EsTdbGrid
    EsTdbGrid = mvarEsTdbGrid
End Property

Public Property Set RecordsetTrabajo(ByVal vData As Object)
Attribute RecordsetTrabajo.VB_Description = "Nombre Del recordset con el que se trabaja y que alimenta al Grid."
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.RecordsetTrabajo = Form1
    Set mvarRecordsetTrabajo = vData
    TipoRs = mvarRecordsetTrabajo.CursorType
    TipoBloqueo = mvarRecordsetTrabajo.LockType
End Property

Public Property Get RecordsetTrabajo() As Object
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.RecordsetTrabajo
    Set RecordsetTrabajo = mvarRecordsetTrabajo
End Property

Public Property Set GridTrabajo(ByVal vData As Object)
Attribute GridTrabajo.VB_Description = "Grid que tiene la información con la que se quiere trabajar."
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.GridTrabajo = Form1
    Set mvarGridTrabajo = vData
End Property

Public Property Get GridTrabajo() As Object
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.GridTrabajo
    Set GridTrabajo = mvarGridTrabajo
End Property

Public Property Set Conexión(ByVal vData As Object)
Attribute Conexión.VB_Description = "La Conexión que va a usar para refrescar el recordset."
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

Public Property Let CampoDefecto(ByVal vData As String)
Attribute CampoDefecto.VB_Description = "Nombre del campo por defecto que aparecerá en la lista de búsqueda."
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
Attribute QueryUtilizado.VB_Description = "Es el Query con el que se va a trabajar. Es el mismo que tiene el recordset base al momento de crearlo."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QueryUtilizado = 5
    mvarQueryUtilizado = vData
End Property

Public Property Get QueryUtilizado() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QueryUtilizado
    QueryUtilizado = mvarQueryUtilizado
End Property

Private Sub Class_Initialize()
  mvarQueryUtilizado = ""
  mvarEsTdbGrid = True
  mvarCampoDefecto = ""
  mvarTítulo = "Realice su Elección"
  mvarCamposVisibles = ""
End Sub

Private Sub Class_Terminate()
  Unload GrFrmBuscaEnGridExterno
End Sub

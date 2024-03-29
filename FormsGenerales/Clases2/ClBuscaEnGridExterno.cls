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
Attribute VB_Description = "Habilita herramientas de b�squeda y filtrado, en una ventana que utiliza de base de b�squeda un Grid de un formulario primario que lo invoca."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarQueryUtilizado As String 'copia local
Private mvarCampoDefecto As String 'copia local
Private mvarT�tulo As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarConexi�n As Object 'copia local
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
Attribute RefrescarGrid.VB_Description = "Evento en el que se debe ubicar c�digo para refrescar el Grid con el c�digo."

Private Sub EventoRefrescar()
  RaiseEvent RefrescarGrid
End Sub

Public Sub Ejecutar()
On Error GoTo QueError
  GrFrmBuscaEnGridExterno.GrPrincipal mvarConexi�n, mvarRecordsetTrabajo, _
                                      TipoRs, TipoBloqueo, _
                                      mvarQueryUtilizado, _
                                      mvarGridTrabajo, mvarEsTdbGrid, _
                                      mvarT�tulo, mvarCamposVisibles, mvarCampoDefecto
  Exit Sub
QueError:
  MsgBox vbObjectError + 100
End Sub

Public Property Let CamposVisibles(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.CamposVisibles = 5
    mvarCamposVisibles = vData
End Property

Public Property Get CamposVisibles() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.CamposVisibles
    CamposVisibles = mvarCamposVisibles
End Property

Public Property Let EsTdbGrid(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.EsTdbGrid = 5
    mvarEsTdbGrid = vData
End Property

Public Property Get EsTdbGrid() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.EsTdbGrid
    EsTdbGrid = mvarEsTdbGrid
End Property

Public Property Set RecordsetTrabajo(ByVal vData As Object)
Attribute RecordsetTrabajo.VB_Description = "Nombre Del recordset con el que se trabaja y que alimenta al Grid."
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucci�n Set.
'Syntax: Set x.RecordsetTrabajo = Form1
    Set mvarRecordsetTrabajo = vData
    TipoRs = mvarRecordsetTrabajo.CursorType
    TipoBloqueo = mvarRecordsetTrabajo.LockType
End Property

Public Property Get RecordsetTrabajo() As Object
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.RecordsetTrabajo
    Set RecordsetTrabajo = mvarRecordsetTrabajo
End Property

Public Property Set GridTrabajo(ByVal vData As Object)
Attribute GridTrabajo.VB_Description = "Grid que tiene la informaci�n con la que se quiere trabajar."
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucci�n Set.
'Syntax: Set x.GridTrabajo = Form1
    Set mvarGridTrabajo = vData
End Property

Public Property Get GridTrabajo() As Object
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.GridTrabajo
    Set GridTrabajo = mvarGridTrabajo
End Property

Public Property Set Conexi�n(ByVal vData As Object)
Attribute Conexi�n.VB_Description = "La Conexi�n que va a usar para refrescar el recordset."
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucci�n Set.
'Syntax: Set x.Conexi�n = Form1
    Set mvarConexi�n = vData
End Property

Public Property Get Conexi�n() As Object
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.Conexi�n
    Set Conexi�n = mvarConexi�n
End Property

Public Property Let T�tulo(ByVal vData As String)
Attribute T�tulo.VB_Description = "T�tulo de la ventana en la que se permitir� el trabajo."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.T�tulo = 5
    mvarT�tulo = vData
End Property

Public Property Get T�tulo() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.T�tulo
    T�tulo = mvarT�tulo
End Property

Public Property Let CampoDefecto(ByVal vData As String)
Attribute CampoDefecto.VB_Description = "Nombre del campo por defecto que aparecer� en la lista de b�squeda."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.CampoDefecto = 5
    mvarCampoDefecto = vData
End Property

Public Property Get CampoDefecto() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.CampoDefecto
    CampoDefecto = mvarCampoDefecto
End Property


Public Property Let QueryUtilizado(ByVal vData As String)
Attribute QueryUtilizado.VB_Description = "Es el Query con el que se va a trabajar. Es el mismo que tiene el recordset base al momento de crearlo."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.QueryUtilizado = 5
    mvarQueryUtilizado = vData
End Property

Public Property Get QueryUtilizado() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.QueryUtilizado
    QueryUtilizado = mvarQueryUtilizado
End Property

Private Sub Class_Initialize()
  mvarQueryUtilizado = ""
  mvarEsTdbGrid = True
  mvarCampoDefecto = ""
  mvarT�tulo = "Realice su Elecci�n"
  mvarCamposVisibles = ""
End Sub

Private Sub Class_Terminate()
  Unload GrFrmBuscaEnGridExterno
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Movil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A7FF7A029F"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Public Enum eEstadoMovil
    eDisponible = 1
    eAtendiendo = 2
    eEnReparacion = 3
End Enum

'##ModelId=46A80483004E
Private mId As Long

'##ModelId=46A804B0000F
Private mNombre As String

'##ModelId=46A8E89B0203
Private mPatente As String

'##ModelId=46A8E8A60242
Private mEstado As eEstadoMovil

'##ModelId=46A97774015A
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Movil", "Id Property Get", err.Description)
End Property

'##ModelId=46A97774015B
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Movil", "Id Property Let", err.Description)
End Property

'##ModelId=46A97774015F
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("Movil", "Nombre Property Get", err.Description)
End Property

'##ModelId=46A977740160
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Movil", "Nombre Property Let", err.Description)
End Property

'##ModelId=46A977740164
Public Property Get Patente() As String
    On Error GoTo PatenteErr

    '## Generated default body ...
    Patente = mPatente

    Exit Property
PatenteErr:
    Call ErrorLog("Movil", "Patente Property Get", err.Description)
End Property

'##ModelId=46A977740165
Public Property Let Patente(ByVal vNewValue As String)
    On Error GoTo PatenteErr

    '## Generated default body ...
    mPatente = vNewValue

    Exit Property
PatenteErr:
    Call ErrorLog("Movil", "Patente Property Let", err.Description)
End Property

'##ModelId=46A977740169
Public Property Get Estado() As eEstadoMovil
    On Error GoTo EstadoErr

    '## Generated default body ...
    Estado = mEstado

    Exit Property
EstadoErr:
    Call ErrorLog("Movil", "Estado Property Get", err.Description)
End Property

'##ModelId=46A97774016A
Public Property Let Estado(ByVal vNewValue As eEstadoMovil)
    On Error GoTo EstadoErr

    '## Generated default body ...
    mEstado = vNewValue

    Exit Property
EstadoErr:
    Call ErrorLog("Movil", "Estado Property Let", err.Description)
End Property

Public Function EstadoToString() As String
    EstadoToString = Choose(mEstado, "Disponible", "Atendiendo", "En Reparacion")
End Function

Friend Sub Insert()
    Dim params As New ParametrosManager
    Id = BD.getIdDisponible(tMovil)
    params.Add "Id", Id
    params.Add "Nombre", Nombre
    params.Add "Patente", Patente
    params.Add "Estado", Estado
    
    BD.Insertar tMovil, params
End Sub

Public Sub Update()
    Dim params As New ParametrosManager
    
    params.Add "Id", Id, True
    params.Add "Nombre", Nombre
    params.Add "Patente", Patente
    params.Add "Estado", Estado
    
    BD.Update tMovil, params
End Sub

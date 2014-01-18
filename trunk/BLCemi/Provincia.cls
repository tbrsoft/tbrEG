VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Provincia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A800720261"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A80493034B
Private mId As Long

'##ModelId=46A804A30271
Private mNombre As String

Private mCiudades As CiudadManager

Private mParent As Pais

'##ModelId=46A977E701F6
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Provincia", "Id Property Get", Err.Description)
End Property

'##ModelId=46A977E701F7
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Provincia", "Id Property Let", Err.Description)
End Property

'##ModelId=46A977E701FB
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("Provincia", "Nombre Property Get", Err.Description)
End Property

'##ModelId=46A977E701FC
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Provincia", "Nombre Property Let", Err.Description)
End Property

Public Property Get Ciudades() As CiudadManager
    Set Ciudades = CiudadesLocal.GetByProvincia(Me)
End Property

Public Property Get Parent() As Pais
    Set Parent = mParent
End Property

Friend Property Set Parent(pParent As Pais)
    Set mParent = pParent
End Property

Friend Sub Guardar()
    Dim params As New ParametrosManager
    params.Add "id", mId
    params.Add "nombre", mNombre
    params.Add "idPais", mParent.Id
    BD.Insertar tProvincia, params
End Sub


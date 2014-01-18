VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ciudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A80070007D"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A804910290
Private mId As Long

'##ModelId=46A804A1031C
Private mNombre As String

Private mBarrios As BarrioManager

Private mParent As Provincia

'##ModelId=46A975B80325
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Ciudad", "Id Property Get", Err.Description)
End Property

'##ModelId=46A975B80326
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Ciudad", "Id Property Let", Err.Description)
End Property

'##ModelId=46A975B8032A
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("Ciudad", "Nombre Property Get", Err.Description)
End Property

'##ModelId=46A975B8032B
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Ciudad", "Nombre Property Let", Err.Description)
End Property

Public Property Get Barrios() As BarrioManager
    'fijarse porq esto sirve para cuando la ciudad ya esta creada
    'no se si sirve cuando creo la instancia
    Set Barrios = BarriosLocal.GetByCiudad(Me)
End Property

Public Property Get Parent() As Provincia
    Set Parent = mParent
End Property

Friend Property Set Parent(pParent As Provincia)
    Set mParent = pParent
End Property

Friend Sub Guardar()
    Dim params As New ParametrosManager
    params.Add "id", mId
    params.Add "nombre", mNombre
    params.Add "idProvincia", mParent.Id
    BD.Insertar tCiudad, params
    
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Barrio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A8006D029F"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A804880157
Private mId As Long

'##ModelId=46A8049F0271
Private mNombre As String

Private mParent As Ciudad

'##ModelId=46A9758403DA
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Barrio", "Id Property Get", Err.Description)
End Property

'##ModelId=46A9758403DB
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Barrio", "Id Property Let", Err.Description)
End Property

'##ModelId=46A9758403DF
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("Barrio", "Nombre Property Get", Err.Description)
End Property

'##ModelId=46A9758403E0
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Barrio", "Nombre Property Let", Err.Description)
End Property

Public Property Get Parent() As Ciudad
    Set Parent = mParent
End Property

Friend Property Set Parent(pParent As Ciudad)
    Set mParent = pParent
End Property

Friend Sub Guardar()
    Dim params As New ParametrosManager
    params.Add "id", mId
    params.Add "nombre", mNombre
    params.Add "idCiudad", mParent.Id
    BD.Insertar tBarrio, params
End Sub

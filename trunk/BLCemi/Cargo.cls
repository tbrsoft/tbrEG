VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A7FF79004E"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A804810186
Private mId As Long

'##ModelId=46A804B2032C
Private mNombre As String

'##ModelId=46A975A00306
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Cargo", "Id Property Get", Err.Description)
End Property

'##ModelId=46A975A00307
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Cargo", "Id Property Let", Err.Description)
End Property

'##ModelId=46A975A0030B
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("Cargo", "Nombre Property Get", Err.Description)
End Property

'##ModelId=46A975A0030C
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Cargo", "Nombre Property Let", Err.Description)
End Property

Friend Sub Guardar()
    Dim parametros As New ParametrosManager
    parametros.Add "Id", mId
    parametros.Add "Nombre", mNombre
    BD.Insertar tCargo, parametros
End Sub

'probablemente friend
Public Sub GuardarModificaciones()
    Dim parametros As New ParametrosManager
    parametros.Add "Id", mId, True
    parametros.Add "Nombre", mNombre
    BD.Update tCargo, parametros
End Sub

Public Sub Eliminar()
MsgBox "Implementar!"
End Sub

Public Property Get CanModify() As Boolean
MsgBox "Implementar!"
End Property

Public Property Get CanDelete() As Boolean
MsgBox "Implementar!"
End Property

Public Function Clone() As Cargo
    Dim c As New Cargo
    With c
        .Id = Me.Id
        .Nombre = Me.Nombre
    End With
    Set Clone = c
End Function

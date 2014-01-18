VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lugar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A980A50128"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A980BD00BB
Private mId As Long

'##ModelId=46A980BE00EA
Private mNombre As String

'##ModelId=46A9818302AF
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Lugar", "Id Property Get", err.Description)
End Property

'##ModelId=46A9818302B0
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Lugar", "Id Property Let", err.Description)
End Property

'##ModelId=46A9818302B4
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("Lugar", "Nombre Property Get", err.Description)
End Property

'##ModelId=46A9818302B5
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Lugar", "Nombre Property Let", err.Description)
End Property

Public Sub Guardar()
    
    Dim parametros As New ParametrosManager
    parametros.Add "Id", mId
    parametros.Add "Nombre", mNombre
    BD.Transaccion.AddInsertCommand tLugar, parametros
    BD.Transaccion.Execute
End Sub

Public Sub GuardarModificaciones()
    MsgBox "Implementar!"
End Sub

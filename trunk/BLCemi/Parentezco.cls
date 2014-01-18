VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parentezco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'##ModelId=46A80477002E
Private mId As Long

'##ModelId=46A804C703C8
Private mNombre As String

'##ModelId=46A977A8033F
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Parentezco", "Id Property Get", Err.Description)
End Property

'##ModelId=46A977A80340
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Parentezco", "Id Property Let", Err.Description)
End Property

'##ModelId=46A977A80344
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("Parentezco", "Nombre Property Get", Err.Description)
End Property

'##ModelId=46A977A80345
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Parentezco", "Nombre Property Let", Err.Description)
End Property

Friend Sub Guardar()
    Dim parametros As New ParametrosManager
    parametros.Add "Id", mId
    parametros.Add "Nombre", mNombre
    BD.Insertar tParentezco, parametros
End Sub

Public Sub GuardarModificaciones()
    Dim parametros As New ParametrosManager
    parametros.Add "Id", mId, True
    parametros.Add "Nombre", mNombre
    BD.Update tParentezco, parametros
End Sub


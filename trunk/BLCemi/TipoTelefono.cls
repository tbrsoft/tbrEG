VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TipoTelefono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A80928005D"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A8094202FD
Private mId As Long

'##ModelId=46A80944031C
Private mNombre As String

'##ModelId=46A978AC02C9
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("TipoTelefono", "Id Property Get", Err.Description)
End Property

'##ModelId=46A978AC02CA
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("TipoTelefono", "Id Property Let", Err.Description)
End Property

'##ModelId=46A978AC02CE
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("TipoTelefono", "Nombre Property Get", Err.Description)
End Property

'##ModelId=46A978AC02CF
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("TipoTelefono", "Nombre Property Let", Err.Description)
End Property

Friend Sub Guardar()
    Dim parametros As New ParametrosManager
    parametros.Add "Id", mId
    parametros.Add "Nombre", mNombre
    BD.Insertar tTipoTelefono, parametros
End Sub

Public Sub GuardarModificaciones()
    Dim parametros As New ParametrosManager
    parametros.Add "Id", mId, True
    parametros.Add "Nombre", mNombre
    BD.Update tTipoTelefono, parametros
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnidadCuerpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
    Call ErrorLog("UnidadCuerpo", "Id Property Get", err.Description)
End Property

'##ModelId=46A978AC02CA
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("UnidadCuerpo", "Id Property Let", err.Description)
End Property

'##ModelId=46A978AC02CE
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("UnidadCuerpo", "Nombre Property Get", err.Description)
End Property

'##ModelId=46A978AC02CF
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("UnidadCuerpo", "Nombre Property Let", err.Description)
End Property

Friend Sub Guardar(pIdCuerpo As Long)
    Dim parametros As New ParametrosManager
    
    parametros.Add "Nombre", mNombre
    
    If mId = 0 Then
        mId = BD.getIdDisponible(tUnidadCuerpo)
        parametros.Add "Id", mId, True
        parametros.Add "idCuerpo", pIdCuerpo
        BD.Insertar tUnidadCuerpo, parametros
    Else
        parametros.Add "Id", mId, True
        BD.Update tUnidadCuerpo, parametros
    End If
    
End Sub

'Public Sub GuardarModificaciones()
'    Dim parametros As New ParametrosManager
'    parametros.Add "Id", mId, True
'    parametros.Add "Nombre", mNombre
'    BD.Update tUnidadCuerpo, parametros
'End Sub




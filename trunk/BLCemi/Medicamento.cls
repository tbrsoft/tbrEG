VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Medicamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mId As Long

Private mNombre As String


Public Property Get Id() As Long
    On Error GoTo idErr

    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Medicamento", "Id Property Get", Err.Description)
End Property

Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Medicamento", "Id Property Let", Err.Description)
End Property

Public Property Get Nombre() As String
    On Error GoTo nombreErr

    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("Medicamento", "Nombre Property Get", Err.Description)
End Property

Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Medicamento", "Nombre Property Let", Err.Description)
End Property

Friend Sub Guardar()
    Dim parametros As New ParametrosManager
    parametros.Add "Id", mId
    parametros.Add "Nombre", mNombre
    BD.Insertar tMedicamento, parametros
End Sub

Public Sub GuardarModificaciones()
    Dim parametros As New ParametrosManager
    parametros.Add "Id", mId, True
    parametros.Add "Nombre", mNombre
    BD.Update tMedicamento, parametros
End Sub



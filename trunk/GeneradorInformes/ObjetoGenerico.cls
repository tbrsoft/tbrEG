VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjetoGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim a As Dictionary

Public Function GetProperty(pNombre As String) As String
    GetProperty = a.Item(LCase(pNombre))
End Function

Public Sub AddProperty(pName As String, pValue As Variant)
    a.Add LCase(pName), pValue
End Sub

Private Sub Class_Initialize()
    Set a = New Dictionary
End Sub

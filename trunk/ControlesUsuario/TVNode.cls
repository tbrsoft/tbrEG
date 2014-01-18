VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TVNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'esta clase representa una clase de nodo, no nodos individuales a diferencia del node del treview
Public id As Long
Public BackColor As Long
Public Bold As Boolean
Public Expanded As Boolean 'osea, todos los nodos expandidos
Public ForeColor As Long
Public ChildCollectionField As String
Public TextField As String
Public IdField As String
Public Child As TVNode
Public Parent As TVNode
Private mCol As Object

'Public Property Get ChildCollection(pObj As Object) As Object
'    Set ChildCollection = CallByName(pObj, ChildCollectionField, VbGet)
'End Property

Public Property Get Collection() As Object
    Set Collection = mCol
End Property

Public Property Set Collection(pCol As Object)
    Set mCol = pCol
End Property


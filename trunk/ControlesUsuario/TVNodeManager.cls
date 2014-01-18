VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TVNodeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'variable local para contener colección
Public Event NodeAdded()

Private mCol As Collection

Public Sub AddItem(pTVNode As TVNode)
    mCol.Add pTVNode, "k" + Str(pTVNode.id)
    RaiseEvent NodeAdded
End Sub

Public Function Add(pId As Long, pIdField As String, pTextField As String, pChildCollection As String, pBackColor As Long, pBold As Boolean, pExpanded As Boolean, pForeColor As Long) As TVNode
    'crear un nuevo objeto
    On Error GoTo e
    Dim objNewMember As TVNode
    Set objNewMember = New TVNode

    'establecer las propiedades que se transfieren al método
    objNewMember.BackColor = pBackColor
    objNewMember.Bold = pBold
    objNewMember.ChildCollectionField = pChildCollection
    objNewMember.Expanded = pExpanded
    objNewMember.ForeColor = pForeColor
    objNewMember.IdField = pIdField
    objNewMember.TextField = pTextField
    
100 mCol.AddItem objNewMember ', objNewMember.id   'estaba nombre, lo cambie por miembro para probar...
    
    'devolver el objeto creado
    Set Add = objNewMember
    Set objNewMember = Nothing
    RaiseEvent NodeAdded
    Exit Function
e:
    If Erl = 100 Then
        Err.Raise 457 'la clave ya esta asociada a la coleccion...
    End If
End Function

Public Property Get Item(pKey As String) As TVNode
    On Error GoTo errman
    Set Item = mCol(pKey)
    Exit Property
errman:
    If Err.Number = 5 Then Err.Raise 2011, , "No se encontro el Nodo."
    
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(pId As Long)
    mCol.Remove "k" + Str(pId)
End Sub

Public Sub Clear()
    Set mCol = Nothing
    Set mCol = New Collection
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContactsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mCol As Collection

Public Sub AddItem(pContact As Contact)
    On Error GoTo AddItemErr

    mCol.Add pContact, "k" + Str(pContact.Id)
    
    Exit Sub
AddItemErr:
    
End Sub

Friend Function Add(pId As Long, pName As String) As Contact
Dim c As New Contact
c.Id = pId
c.ContactName = pPathAudio

AddItem c
Set Add = c
Set c = Nothing
End Function

Friend Function NewContact(pName As String) As Contact
    Dim c As New Contact
    c.Id = bd.getIdDisponible(tContact)
    c.ContactName = pName
    c.Save
    Me.AddItem c
    Set NewContact = c
End Function

Friend Sub LoadAll()
On Error GoTo errMaq

Dim rs
Set rs = bd.SelectTabla(etContact)
While Not rs.EOF
        
    Add rs!Id, rs!nombre
    rs.MoveNext
    
Wend
Exit Sub
errMaq:

End Sub

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Property Get Item(pId As Long) As Contact
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

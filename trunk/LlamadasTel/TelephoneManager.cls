VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TelephoneManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mCol As Collection

Public Sub AddItem(pTelephone As Telephone)
    On Error GoTo AddItemErr

    mCol.Add pTelephone, "k" + Str(pTelephone.Id)
    
    Exit Sub
AddItemErr:
    
End Sub

Friend Function Add(pId As Long, pTelNumber As String) As Telephone
Dim c As New Telephone

c.Id = pId
c.TelNumber = pTelNumber

AddItem c
Set Add = c
Set c = Nothing
End Function

Friend Function NewTelephone(pTelNumber As String) As Telephone
    Dim c As New Telephone
    c.Id = bd.getIdDisponible(tTelephone)
    c.TelNumber = pTelNumber
    
    c.Save
    Me.AddItem c
    Set NewTelephone = c
End Function

Friend Sub LoadAll()
On Error GoTo errMaq

Dim rs
Set rs = bd.SelectTabla(tTelephone)
While Not rs.EOF
        
    Add rs!Id, rs!TelNumber
    rs.MoveNext
    
Wend
Exit Sub
errMaq:

End Sub

Friend Function GetByContact(pIdContact As Long)
Dim tm As New TelephoneManager
Dim t As Telephone
For Each t In Me
    
Next
End Function

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

Public Property Get Item(pId As Long) As Telephone
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

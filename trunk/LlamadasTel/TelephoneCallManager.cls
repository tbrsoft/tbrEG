VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TelephoneCallManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mCol As Collection

Public Sub AddItem(pTelephoneCall As TelephoneCall)
    On Error GoTo AddItemErr

    mCol.Add pTelephoneCall, "k" + Str(pTelephoneCall.Id)
    
    Exit Sub
AddItemErr:
    
End Sub

Friend Function Add(pId As Long, pTelNumber As String, pDuration As Integer, pDateTime As Date, pPathAudio As String) As TelephoneCall
Dim c As New TelephoneCall
c.Id = pId
c.PathAudio = pPathAudio
c.CallDuration = pDuration
c.TelNumber = pTelNumber
c.CallDateTime = pDateTime

AddItem c
Set Add = c
Set c = Nothing
End Function

Friend Function NewTelephoneCall(pTelNumber As String, pDuration As Integer, pDateTime As Date, pPathAudio As String) As TelephoneCall
    Dim c As New TelephoneCall
    c.Id = bd.getIdDisponible(tTelephoneCall)
    c.CallDateTime = pDateTime
    c.CallDuration = pDuration
    c.PathAudio = pPathAudio
    c.TelNumber = pTelNumber
    
    c.Save
    Me.AddItem c
    Set NewTelephoneCall = c
End Function

Friend Sub LoadAll()
On Error GoTo errMaq

Dim rs
Set rs = bd.SelectTabla(tTelephoneCall)
While Not rs.EOF
        
    Add rs!Id, rs!TelNumber, rs!duration, rs!CallDateTime, rs!PathAudio
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

Public Property Get Item(pId As Long) As TelephoneCall
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

Public Property Get GetByDate(pDate As Date) As TelephoneCallManager
'    Dim telCalls As New TelephoneCallManager
'    Dim rs
'    Set rs = DB.SelectSQL("Select id from TelephoneCall where uname like " + pUsername)
'    While Not rs.EOF
'        telCalls.AddItem LocalTelCalls.Item(rs!Id)
'        rs.MoveNext
'    Wend
End Property

Public Property Get GetByUser(pUsername As String) As TelephoneCallManager
    Dim telCalls As New TelephoneCallManager
    Dim rs
    Set rs = DB.SelectSQL("Select id from TelephoneCall where uname like " + pUsername)
    While Not rs.EOF
        telCalls.AddItem LocalTelCalls.Item(rs!Id)
        rs.MoveNext
    Wend
End Property


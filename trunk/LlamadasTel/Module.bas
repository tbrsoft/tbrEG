Attribute VB_Name = "Module"
Dim mTelCalls As TelephoneCallManager
Dim mUsers As UserManager
Dim mTelephones As TelephoneManager
Dim mDB As DataBaseL

Sub main()

End Sub

Public Property Get LocalTelCalls() As TelephoneCallManager
    If mTelCalls Is Nothing Then
        Set mTelCalls = New TelephoneCallManager
        mTelCalls.LoadAll
    End If
    Set LocalTelCalls = mTelCalls
End Property

Public Property Get LocalUsers() As UserManager
    If mUsers Is Nothing Then
        Set mUsers = New UserManager
        mUsers.LoadUsers
    End If
    Set LocalUsers = mUsers
End Property

Public Property Get LocalTelephones() As TelephoneManager
    If mTelephones Is Nothing Then
        Set mTelephones = New TelephoneManager
        mTelephones.LoadAll
    End If
    Set LocalTelephones = mTelephones
End Property

Public Property Get DB() As DataBaseL
    If mDB Is Nothing Then
        Set mDB = New DataBaseL
    End If
    Set Localdb = mDB
End Property


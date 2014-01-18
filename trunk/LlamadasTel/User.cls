VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public UserName As String
Public Pass As String

Friend Sub Save()
    On Error Resume Next
    DB.SelectSQL "insert into Users(uname,pass) values (" + UserName + "," + Pass + ")"
End Sub

Public Property Get Calls() As TelephoneCallManager
    Set Calls = LocalTelCalls.GetByUser(UserName)
End Property

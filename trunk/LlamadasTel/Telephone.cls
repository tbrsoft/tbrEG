VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Telephone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mTelNumber As String
Private mTelCalls As TelephoneCallManager
Private mContactId As Long

Public Property Let TelNumber(ByVal vData As String)
    mTelNumber = vData
End Property

Public Property Get TelNumber() As String
    TelNumber = mTelNumber
End Property

Public Property Get Owner() As Contact
    
End Property

Friend Property Let OwnerId(pValue As Long)
    mContactId = pValue
End Property

Public Property Get Calls() As TelephoneCallManager
    Set Calls = mTelCalls
End Property

Public Property Set Calls(pValue As TelephoneCallManager)
    Set mTelCalls = pValue
End Property


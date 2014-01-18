VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TelephoneCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mId As Double
Private mPathAudio As String
Private mTel As Telephone
Private mDuration As Integer
Private mDateTime As Date
Private mCallType As eCallType

Public Enum eCallType
    ectIncoming = 1
    ectOutgoing = 2
End Enum

Public Property Let Telephone(ByVal vData As Telephone)
    mTel = vData
End Property

Public Property Get Telephone() As Telephone
    Telephone = mTel
End Property

Public Property Let PathAudio(ByVal vData As String)
    mPathAudio = vData
End Property

Public Property Get PathAudio() As String
    PathAudio = mPathAudio
End Property

Public Property Let Id(ByVal vData As Double)
    mId = vData
End Property

Public Property Get Id() As Double
    Id = mId
End Property

Public Property Get CallDuration() As Integer
    CallDuration = mDuration
End Property

Public Property Let CallDuration(pValue As Integer)
    mDuration = pValue
End Property

Public Property Get CallDateTime() As Date
    CallDateTime = mDateTime
End Property

Public Property Let CallDateTime(pValue As Date)
    mDateTime = pValue
End Property

Public Property Get CallType() As eCallType
    CallType = mCallType
End Property

Public Property Let CallType(pValue As eCallType)
    mCallType = pValue
End Property

Public Property Get CallTypeToString() As eCallType
    If mCallType = ectIncoming Then
        CallTypeToString = "Incoming"
    Else
        CallTypeToString = "Outgoing"
    End If
End Property

Friend Sub Save()
    DB.SelectSQL "insert into telephonecall (id,telNumber, duration, calldatetime, uname, tipoLlamada, pathaudio) values(" + mId + ",'" + mTelNumber + "'," + Str(mDuration) + ",'" + Str(mDateTime) + "','" + uname + "'," + Str(mCallType) + ",'" + "')"
End Sub

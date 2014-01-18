VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MonthYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Month As Integer
Public Year As Integer

Public Function ToLongString() As String
    ToLongString = MonthName(Month) + " de " + Trim(Str(Year))
End Function

Public Function ToShortString() As String
    ToShortString = MonthName(Month, True) + "-" + Trim(Str(Year))
End Function


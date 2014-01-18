VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Event Error(eError As eDBErrors, pDescription As String)

Public Enum eDBErrors
    eUnknown = 1
    eDBCantFindDB = 2
    eDBConnectionClosed = 3
End Enum

Friend Sub FireErrorEvent(pError As eDBErrors, pDescription As String)
    RaiseEvent Error(pError, pDescription)
End Sub

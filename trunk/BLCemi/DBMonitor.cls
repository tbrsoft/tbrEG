VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'esta clase sirve para monitorear la base de datos.
Public Event Error(eError As eDBErrors, pDescription As String)

Public Enum eDBErrors
    eUnknown = 1
    eDBCantFindDB = 2
    eDBConnectionClosed = 3
End Enum

Dim WithEvents mDbEvents As DBEvents
Attribute mDbEvents.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mDbEvents = BD.DBEventsManager
End Sub

Private Sub mDbEvents_Error(eError As DataBaseLayer.eDBErrors, pDescription As String)
    ErrorLog "DBMonitor", "Evento Error", pDescription
    RaiseEvent Error(eError, pDescription)
End Sub

Public Sub Activate()
'es para q se instancie la clase y empiece a levantar eventos
End Sub
